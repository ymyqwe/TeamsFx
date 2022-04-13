// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import * as path from "path";
import * as fs from "fs-extra";
import { Argv } from "yargs";
import {
  assembleError,
  AzureSolutionSettings,
  Colors,
  ConfigFolderName,
  err,
  FxError,
  InputConfigsFolderName,
  Inputs,
  LogLevel,
  ok,
  Platform,
  ProjectSettings,
  ProjectSettingsFileName,
  Result,
  SystemError,
  UnknownError,
  UserError,
} from "@microsoft/teamsfx-api";
import {
  DepsType,
  FolderName,
  FxCore,
  ITaskDefinition,
  loadTeamsFxDevScript,
  LocalEnvManager,
  ProjectSettingsHelper,
  TaskDefinition,
  ProgrammingLanguage,
  isConfigUnifyEnabled,
  environmentManager,
  DepsManager,
  getSideloadingStatus,
  NodeNotSupportedError,
  isPureExistingApp,
} from "@microsoft/teamsfx-core";

import { YargsCommand } from "../../yargsCommand";
import * as utils from "../../utils";
import * as commonUtils from "./commonUtils";
import * as constants from "./constants";
import { doctorResult } from "./constants";
import cliLogger from "../../commonlib/log";
import * as errors from "./errors";
import activate from "../../activate";
import { Task, TaskResult } from "./task";
import AppStudioTokenInstance from "../../commonlib/appStudioLogin";
import cliTelemetry, { CliTelemetry } from "../../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../../telemetry/cliTelemetryEvents";
import { ServiceLogWriter } from "./serviceLogWriter";
import CLIUIInstance from "../../userInteraction";
import { cliEnvCheckerLogger } from "./depsChecker/cliLogger";
import { cliEnvCheckerTelemetry } from "./depsChecker/cliTelemetry";
import { URL } from "url";
import { CliDepsChecker } from "./depsChecker/cliChecker";
import { isNgrokCheckerEnabled, isTrustDevCertEnabled } from "./depsChecker/cliUtils";
import { signedOut } from "../../commonlib/common/constant";
import { cliSource } from "../../constants";
import { performance } from "perf_hooks";

enum Checker {
  M365Account = "M365 Account",
  LocalCertificate = "Development certificate for localhost",
  Ports = "Ports",
}

const DepsDisplayName = {
  [DepsType.FunctionNode]: "Node.js",
  [DepsType.SpfxNode]: "Node.js",
  [DepsType.AzureNode]: "Node.js",
  [DepsType.Dotnet]: ".NET Core SDK",
  [DepsType.Ngrok]: "Ngrok",
  [DepsType.FuncCoreTools]: "Azure Functions Core Tools",
};

const ProgressMessage: { [key: string]: string } = Object.freeze({
  [Checker.M365Account]: `Checking ${Checker.M365Account}`,
  [Checker.LocalCertificate]: `Checking ${Checker.LocalCertificate}`,
  [Checker.Ports]: `Checking ${Checker.Ports}`,
  [DepsType.FunctionNode]: `Checking ${DepsDisplayName[DepsType.FunctionNode]}`,
  [DepsType.SpfxNode]: `Checking ${DepsDisplayName[DepsType.SpfxNode]}`,
  [DepsType.AzureNode]: `Checking ${DepsDisplayName[DepsType.AzureNode]}`,
  [DepsType.Dotnet]: `Checking and installing ${DepsDisplayName[DepsType.Dotnet]}`,
  [DepsType.Ngrok]: `Checking and installing ${DepsDisplayName[DepsType.Ngrok]}`,
  [DepsType.FuncCoreTools]: `Checking and installing ${DepsDisplayName[DepsType.FuncCoreTools]}`,
});

export default class Preview extends YargsCommand {
  public readonly commandHead = `preview`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Preview the current application.";

  private backgroundTasks: Task[] = [];
  private readonly telemetryProperties: { [key: string]: string } = {};
  private readonly telemetryMeasurements: { [key: string]: number } = {};
  private serviceLogWriter: ServiceLogWriter | undefined;
  private sharepointSiteUrl: string | undefined;
  public builder(yargs: Argv): Argv<any> {
    yargs.option("local", {
      description: "Preview the application from local, exclusive with --remote",
      boolean: true,
      default: false,
    });
    yargs.option("remote", {
      description: "Preview the application from remote, exclusive with --local",
      boolean: true,
      default: false,
    });
    yargs.option("folder", {
      description: "Select root folder of the project",
      string: true,
      default: "./",
    });
    yargs.option("browser", {
      description: "Select browser to open Teams web client",
      string: true,
      choices: [constants.Browser.chrome, constants.Browser.edge, constants.Browser.default],
      default: constants.Browser.default,
    });
    yargs.option("browser-arg", {
      description:
        'Argument to pass to the browser, requires --browser, can be used multiple times (e.g. --browser-args="--guest")',
      string: true,
    });
    yargs.option("sharepoint-site", {
      description:
        "SharePoint site URL, like {your-tenant-name}.sharepoint.com [only for SPFx project remote preview]",
      array: false,
      string: true,
    });
    yargs.option("env", {
      description: "Select an existing env for the project",
      string: true,
    });

    return yargs.version(false);
  }

  public async runCommand(args: {
    [argName: string]: boolean | string | string[] | undefined;
  }): Promise<Result<null, FxError>> {
    try {
      let previewType = "";
      if ((args.local && !args.remote) || (!args.local && !args.remote)) {
        previewType = "local";
      } else if (!args.local && args.remote) {
        previewType = "remote";
      }
      this.telemetryProperties[TelemetryProperty.PreviewType] = previewType;

      const workspaceFolder = path.resolve(args.folder as string);
      this.telemetryProperties[TelemetryProperty.PreviewAppId] = utils.getLocalTeamsAppId(
        workspaceFolder
      ) as string;

      cliTelemetry
        .withRootFolder(workspaceFolder)
        .sendTelemetryEvent(TelemetryEvent.PreviewStart, this.telemetryProperties);

      const browser = args.browser as constants.Browser;
      this.telemetryProperties[TelemetryProperty.PreviewBrowser] = browser;

      const browserArguments: string[] = [];
      if (args["browser-arg"]) {
        if (Array.isArray(args["browser-arg"])) {
          args["browser-arg"].forEach((x) => browserArguments.push(x));
        } else {
          browserArguments.push(args["browser-arg"] as string);
        }
      }

      // parse sharepoint site url to get workbench url
      if (args["sharepoint-site"]) {
        try {
          let spSite = args["sharepoint-site"] as string;
          if (!spSite.startsWith("https")) {
            spSite = `https://${spSite}`;
          }
          const spWorkbenchHttpsUrl = new URL("_layouts/workbench.aspx", spSite);
          this.sharepointSiteUrl = spWorkbenchHttpsUrl.toString();
        } catch (error: any) {
          throw errors.InvalidSharePointSiteURL(error);
        }
      }
      if (args.local && args.remote) {
        throw errors.ExclusiveLocalRemoteOptions();
      }

      let result: Result<null, FxError>;
      if (previewType === "local") {
        if (await this.isExistingApp(workspaceFolder)) {
          result = await this.localPreviewMinimalApp(workspaceFolder, browser, browserArguments);
        } else {
          result = await this.localPreview(workspaceFolder, browser, browserArguments);
        }
      } else {
        result = await this.remotePreview(
          workspaceFolder,
          browser,
          args.env as any,
          browserArguments
        );
      }

      if (result.isErr()) {
        throw result.error;
      }
      cliTelemetry.sendTelemetryEvent(
        TelemetryEvent.Preview,
        {
          ...this.telemetryProperties,
          [TelemetryProperty.Success]: TelemetrySuccess.Yes,
        },
        this.telemetryMeasurements
      );
      return ok(null);
    } catch (error: any) {
      cliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Preview, error, this.telemetryProperties);
      await this.terminateTasks();
      return err(error);
    }
  }

  private async localPreview(
    workspaceFolder: string,
    browser: constants.Browser,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    const coreResult = await activate();
    if (coreResult.isErr()) {
      return err(coreResult.error);
    }
    const core = coreResult.value;

    const skipNgrok = !(await isNgrokCheckerEnabled());
    const trustDevCert = await isTrustDevCertEnabled();
    let ignoreEnvInfo = true;
    if (isConfigUnifyEnabled()) {
      ignoreEnvInfo = false;
    }
    const inputs: Inputs = {
      projectPath: workspaceFolder,
      platform: Platform.CLI,
      ignoreEnvInfo: ignoreEnvInfo, // local debug does not require environments
      checkerInfo: {
        skipNgrok: skipNgrok,
        trustDevCert: trustDevCert,
      },
      env: isConfigUnifyEnabled() ? environmentManager.getLocalEnvName() : undefined,
    };

    const localEnvManager = new LocalEnvManager(cliLogger, CliTelemetry.getReporter());
    const projectSettings = await localEnvManager.getProjectSettings(workspaceFolder);
    let localSettings = undefined;
    let configResult = undefined;
    if (!isConfigUnifyEnabled()) {
      localSettings = await localEnvManager.getLocalSettings(workspaceFolder); // here does not need crypt data
    }
    const includeFrontend = ProjectSettingsHelper.includeFrontend(projectSettings);
    const includeBackend = ProjectSettingsHelper.includeBackend(projectSettings);
    const includeBot = ProjectSettingsHelper.includeBot(projectSettings);
    const includeSpfx = ProjectSettingsHelper.isSpfx(projectSettings);
    const includeSimpleAuth = ProjectSettingsHelper.includeSimpleAuth(projectSettings);
    const includeFuncHostedBot = ProjectSettingsHelper.includeFuncHostedBot(projectSettings);

    // TODO: move path validation to core
    const spfxRoot = path.join(workspaceFolder, FolderName.SPFx);
    if (includeSpfx && !(await fs.pathExists(spfxRoot))) {
      return err(errors.RequiredPathNotExists(spfxRoot));
    }

    const frontendRoot = path.join(workspaceFolder, FolderName.Frontend);
    if (includeFrontend && !(await fs.pathExists(frontendRoot))) {
      return err(errors.RequiredPathNotExists(frontendRoot));
    }

    const backendRoot = path.join(workspaceFolder, FolderName.Function);
    if (includeBackend && !(await fs.pathExists(backendRoot))) {
      return err(errors.RequiredPathNotExists(backendRoot));
    }

    const botRoot = path.join(workspaceFolder, FolderName.Bot);
    if (includeBot && !(await fs.pathExists(botRoot))) {
      return err(errors.RequiredPathNotExists(botRoot));
    }

    if (includeSpfx) {
      return this.spfxPreview(
        workspaceFolder,
        browser,
        "https://localhost:5432/workbench",
        browserArguments
      );
    }

    const start = performance.now();
    const depsManager = new DepsManager(cliEnvCheckerLogger, cliEnvCheckerTelemetry);
    try {
      // check node
      const nodeRes = await this.checkNode(includeBackend, includeFuncHostedBot, depsManager);
      if (nodeRes.isErr()) {
        return err(nodeRes.error);
      }

      // check account
      const accountRes = await this.checkM365Account();
      if (accountRes.isErr()) {
        return err(accountRes.error);
      }

      // check cert
      const certRes = await this.resolveLocalCertificate(localEnvManager);
      if (certRes.isErr()) {
        return err(certRes.error);
      }

      // check deps
      const envCheckerResult = await this.handleDependences(
        projectSettings,
        localEnvManager,
        depsManager
      );
      if (envCheckerResult.isErr()) {
        return err(envCheckerResult.error);
      }

      /* === check ports === */
      const portsRes = await this.checkPorts(workspaceFolder);
      if (portsRes.isErr()) {
        return portsRes;
      }
    } finally {
      // use seconds
      this.telemetryMeasurements[TelemetryProperty.PreviewPrerequisitesCheckTime] = Number(
        ((performance.now() - start) / 1000).toFixed(2)
      );
    }

    // clear background tasks
    this.backgroundTasks = [];
    // init service log writer
    this.serviceLogWriter = new ServiceLogWriter();
    await this.serviceLogWriter.init();

    /* === start ngrok === */
    if (includeBot && !skipNgrok) {
      const result = await this.startNgrok(workspaceFolder, depsManager);
      if (result.isErr()) {
        return result;
      }
    }

    /* === prepare dev env === */
    let result = await this.prepareDevEnv(
      core,
      inputs,
      workspaceFolder,
      includeFrontend,
      includeBackend,
      includeBot,
      depsManager
    );
    if (result.isErr()) {
      return result;
    }

    this.telemetryProperties[TelemetryProperty.PreviewAppId] = utils.getLocalTeamsAppId(
      workspaceFolder
    ) as string;

    /* === start services === */
    const programmingLanguage = projectSettings.programmingLanguage as string;
    if (programmingLanguage === undefined || programmingLanguage.length === 0) {
      return err(errors.MissingProgrammingLanguageSetting());
    }

    result = await this.startServices(
      workspaceFolder,
      programmingLanguage,
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot,
      depsManager,
      includeSimpleAuth
    );
    if (result.isErr()) {
      return result;
    }

    /* === get local teams app id === */
    // re-load local settings
    let tenantId = undefined,
      localTeamsAppId = undefined;
    if (isConfigUnifyEnabled()) {
      configResult = await core.getProjectConfig(inputs);
      if (configResult.isErr()) {
        return err(configResult.error);
      }
      const config = configResult.value;
      tenantId = config?.config
        ?.get(constants.solutionPluginName)
        ?.get(constants.teamsAppTenantIdConfigKey) as string;
      localTeamsAppId = config?.config
        ?.get(constants.appstudioPluginName)
        ?.get(constants.remoteTeamsAppIdConfigKey);
    } else {
      localSettings = await localEnvManager.getLocalSettings(workspaceFolder); // here does not need crypt data

      tenantId = localSettings?.teamsApp?.tenantId as string;
      localTeamsAppId = localSettings?.teamsApp?.teamsAppId as string;
    }

    if (localTeamsAppId === undefined || localTeamsAppId.length === 0) {
      return err(errors.TeamsAppIdNotExists());
    }

    /* === open teams web client === */
    result = await this.openTeamsWebClient(
      tenantId.length === 0 ? undefined : tenantId,
      localTeamsAppId,
      browser,
      browserArguments
    );
    if (result.isErr()) {
      return result;
    }

    cliLogger.necessaryLog(LogLevel.Warning, constants.waitCtrlPlusC);

    return ok(null);
  }

  private async checkPorts(workspaceFolder: string): Promise<Result<null, FxError>> {
    const portsBar = CLIUIInstance.createProgressBar(Checker.Ports, 1);
    await portsBar.start(ProgressMessage[Checker.Ports]);
    await portsBar.next(ProgressMessage[Checker.Ports]);
    const portsInUse = await commonUtils.getPortsInUse(workspaceFolder);
    if (portsInUse.length > 0) {
      await portsBar.end(false);
      return err(errors.PortsAlreadyInUse(portsInUse));
    }
    await portsBar.end(true);
    return ok(null);
  }

  private async spfxPreviewSetup(workspaceFolder: string): Promise<Result<null, FxError>> {
    // init service log writer
    this.serviceLogWriter = new ServiceLogWriter();
    await this.serviceLogWriter.init();

    // run npm install for spfx
    const spfxInstallTask = this.prepareTask(
      TaskDefinition.spfxInstall(workspaceFolder),
      constants.spfxInstallStartMessage
    );

    let result = await spfxInstallTask.task.wait(spfxInstallTask.startCb, spfxInstallTask.stopCb);
    if (result.isErr()) {
      return err(result.error);
    }

    // run gulp trust-dev-cert
    const gulpCertTask = this.prepareTask(
      TaskDefinition.gulpCert(workspaceFolder),
      constants.gulpCertStartMessage
    );

    result = await gulpCertTask.task.wait(gulpCertTask.startCb, gulpCertTask.stopCb);
    if (result.isErr()) {
      return err(result.error);
    }

    // run gulp serve
    const gulpServeTask = this.prepareTask(
      TaskDefinition.gulpServe(workspaceFolder),
      constants.gulpServeStartMessage
    );

    result = await gulpServeTask.task.waitFor(
      constants.gulpServePattern,
      gulpServeTask.startCb,
      gulpServeTask.stopCb,
      undefined,
      this.serviceLogWriter,
      cliLogger
    );
    if (result.isErr()) {
      return err(result.error);
    }
    return ok(null);
  }

  private async openSPFxWebClient(
    browser: constants.Browser,
    url: string,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    cliTelemetry.sendTelemetryEvent(
      TelemetryEvent.PreviewSPFxOpenBrowserStart,
      this.telemetryProperties
    );

    const previewBar = CLIUIInstance.createProgressBar(constants.previewSPFxTitle, 1);
    await previewBar.start(constants.previewSPFxStartMessage);
    await previewBar.next(constants.previewSPFxStartMessage);
    try {
      await commonUtils.openBrowser(browser, url, browserArguments);
    } catch {
      const error = errors.OpeningBrowserFailed(browser);
      cliTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.PreviewSPFxOpenBrowser,
        error,
        this.telemetryProperties
      );
      cliLogger.necessaryLog(LogLevel.Warning, constants.openBrowserHintMessage);
      cliLogger.necessaryLog(LogLevel.Warning, constants.waitCtrlPlusC);
      await previewBar.end(false);
      return ok(null);
    }

    await previewBar.end(true);
    const message = [
      {
        content: `preview url: `,
        color: Colors.WHITE,
      },
      {
        content: url,
        color: Colors.BRIGHT_CYAN,
      },
    ];
    cliLogger.necessaryLog(LogLevel.Info, utils.getColorizedString(message));

    cliTelemetry.sendTelemetryEvent(TelemetryEvent.PreviewSPFxOpenBrowser, {
      ...this.telemetryProperties,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });

    cliLogger.necessaryLog(LogLevel.Warning, constants.waitCtrlPlusC);
    return ok(null);
  }

  private async spfxPreview(
    workspaceFolder: string,
    browser: constants.Browser,
    url: string,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    {
      const result = await this.spfxPreviewSetup(workspaceFolder);
      if (result.isErr()) {
        return err(result.error);
      }
    }
    {
      const result = await this.openSPFxWebClient(browser, url, browserArguments);
      if (result.isErr()) {
        return err(result.error);
      }
    }
    return ok(null);
  }

  private async localPreviewMinimalApp(
    workspaceFolder: string,
    browser: constants.Browser,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    const coreResult = await activate();
    if (coreResult.isErr()) {
      return err(coreResult.error);
    }
    const core = coreResult.value;

    const inputs: Inputs = {
      projectPath: workspaceFolder,
      platform: Platform.CLI,
      env: environmentManager.getLocalEnvName(),
    };

    /* === register teams app === */
    const result = await core.localDebug(inputs);
    if (result.isErr()) {
      return err(result.error);
    }

    return await this.remotePreview(
      workspaceFolder,
      browser,
      environmentManager.getLocalEnvName(),
      browserArguments
    );
  }

  private async isExistingApp(workspacePath: string): Promise<boolean> {
    const projectSettingsPath = path.resolve(
      workspacePath,
      `.${ConfigFolderName}`,
      InputConfigsFolderName,
      ProjectSettingsFileName
    );

    if (await fs.pathExists(projectSettingsPath)) {
      const projectSettings = await fs.readJson(projectSettingsPath);
      return isPureExistingApp(projectSettings);
    } else {
      return false;
    }
  }

  private async remotePreview(
    workspaceFolder: string,
    browser: constants.Browser,
    env: string | undefined,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    /* === get remote teams app id === */
    const coreResult = await activate(workspaceFolder);
    if (coreResult.isErr()) {
      return err(coreResult.error);
    }
    const core = coreResult.value;

    const inputs: Inputs = {
      projectPath: workspaceFolder,
      platform: Platform.CLI,
      env: env,
    };

    const configResult = await core.getProjectConfig(inputs);
    if (configResult.isErr()) {
      return err(configResult.error);
    }
    const config = configResult.value;

    const activeResourcePlugins =
      (config?.settings?.solutionSettings as AzureSolutionSettings)?.activeResourcePlugins ?? [];
    const includeSpfx = activeResourcePlugins.some(
      (pluginName) => pluginName === constants.spfxPluginName
    );
    if (includeSpfx) {
      if (!this.sharepointSiteUrl) {
        return err(errors.NoUrlForSPFxRemotePreview());
      }
      const spfxRoot = path.join(workspaceFolder, FolderName.SPFx);
      return this.spfxPreview(spfxRoot, browser, this.sharepointSiteUrl, browserArguments);
    }

    const tenantId = config?.config
      ?.get(constants.solutionPluginName)
      ?.get(constants.teamsAppTenantIdConfigKey) as string;

    const remoteTeamsAppId: string = config?.config
      ?.get(constants.appstudioPluginName)
      ?.get(constants.remoteTeamsAppIdConfigKey);
    if (remoteTeamsAppId === undefined || remoteTeamsAppId.length === 0) {
      return err(errors.PreviewWithoutProvision());
    }

    /* === open teams web client === */
    const result = await this.openTeamsWebClient(
      tenantId.length === 0 ? undefined : tenantId,
      remoteTeamsAppId,
      browser,
      browserArguments
    );
    if (result.isErr()) {
      return result;
    }

    return ok(null);
  }

  private async startNgrok(
    workspaceFolder: string,
    depsManager: DepsManager
  ): Promise<Result<null, FxError>> {
    // bot npm install
    const botInstallTask = this.prepareTask(
      TaskDefinition.botInstall(workspaceFolder),
      constants.botInstallStartMessage
    );
    let result = await botInstallTask?.task.wait(botInstallTask?.startCb, botInstallTask?.stopCb);
    if (result.isErr()) {
      return err(errors.PreviewCommandFailed([result.error]));
    }

    // start ngrok
    const ngrok = (await depsManager.getStatus([DepsType.Ngrok]))[0];
    const ngrokBinFolders = ngrok.details.binFolders;
    const ngrokStartTask = this.prepareTask(
      TaskDefinition.ngrokStart(workspaceFolder, false, ngrokBinFolders),
      constants.ngrokStartStartMessage
    );
    result = await ngrokStartTask.task.waitFor(
      constants.ngrokStartPattern,
      ngrokStartTask.startCb,
      ngrokStartTask.stopCb,
      undefined,
      this.serviceLogWriter
    );
    if (result.isErr()) {
      return err(errors.PreviewCommandFailed([result.error]));
    }
    return ok(null);
  }

  private async prepareDevEnv(
    core: FxCore,
    inputs: Inputs,
    workspaceFolder: string,
    includeFrontend: boolean,
    includeBackend: boolean,
    includeBot: boolean,
    depsManager: DepsManager
  ): Promise<Result<null, FxError>> {
    const frontendInstallTask = includeFrontend
      ? this.prepareTask(
          TaskDefinition.frontendInstall(workspaceFolder),
          constants.frontendInstallStartMessage
        )
      : undefined;

    const backendInstallTask = includeBackend
      ? this.prepareTask(
          TaskDefinition.backendInstall(workspaceFolder),
          constants.backendInstallStartMessage
        )
      : undefined;

    const dotnet = (await depsManager.getStatus([DepsType.Dotnet]))[0];
    const dotnetExecPath = dotnet.command;
    const backendExtensionsInstallTask = includeBackend
      ? this.prepareTask(
          TaskDefinition.backendExtensionsInstall(workspaceFolder, dotnetExecPath),
          constants.backendExtensionsInstallStartMessage
        )
      : undefined;

    const botInstallTask = includeBot
      ? this.prepareTask(
          TaskDefinition.botInstall(workspaceFolder),
          constants.botInstallStartMessage
        )
      : undefined;

    const results = await Promise.all([
      core.localDebug(inputs),
      frontendInstallTask?.task.wait(frontendInstallTask.startCb, frontendInstallTask.stopCb),
      backendInstallTask?.task.wait(backendInstallTask.startCb, backendInstallTask.stopCb),
      backendExtensionsInstallTask?.task.wait(
        backendExtensionsInstallTask.startCb,
        backendExtensionsInstallTask.stopCb
      ),
      botInstallTask?.task.wait(botInstallTask.startCb, botInstallTask.stopCb),
    ]);
    const fxErrors: FxError[] = [];
    for (const result of results) {
      if (result?.isErr()) {
        fxErrors.push(result.error);
      }
    }
    if (fxErrors.length > 0) {
      return err(errors.PreviewCommandFailed(fxErrors));
    }
    return ok(null);
  }

  /**
   * Create a promise that run tasks sequentially.
   * @param tasks The tasks to run
   * @returns An array of the results if all tasks succeed, or the FxError of the first failed task.
   *          Tasks after the first failed tasks will not be executed
   * Example:
   *  sequatialTasks(t1, t2, t3)
   *
   *  If t1 succeeds and t2 fails, the result is the error of t2, and t3 is never executed.
   *  If they all succeed, the result is [t1Result, t2Result, t3Result].
   */
  public static async sequentialTasks<T>(
    ...tasks: (() => Promise<Result<T, FxError>> | undefined)[]
  ): Promise<Result<(T | undefined)[], FxError>> {
    const results: (T | undefined)[] = [];
    for (const createTask of tasks) {
      const result = await createTask();
      if (result) {
        if (result.isErr()) {
          return err(result.error);
        } else {
          results.push(result.value);
        }
      } else {
        results.push(undefined);
      }
    }
    return ok(results);
  }

  public async createBotTasksForStartServices(
    workspaceFolder: string,
    programmingLanguage: string,
    includeBot: boolean,
    includeFuncHostedBot: boolean,
    localEnv: { [key: string]: string } | undefined,
    funcEnv: { [key: string]: string } | undefined
  ): Promise<(Promise<Result<unknown, FxError>> | undefined)[]> {
    // The following task logic aligns with the vscode extension.
    // Bot task order:
    //  - for legacy bot: botStart
    //  - for func hosted bot: [botWatch (ts only) -> botStart] | botAzurite
    // "|" for concurrent
    // "->" for sequential
    let botTaskPromises: (Promise<Result<unknown, FxError>> | undefined)[] = [];
    if (includeBot) {
      const hasTeamsFxDevScript =
        (await loadTeamsFxDevScript(path.join(workspaceFolder, FolderName.Bot))) !== undefined;
      const botWatchTask =
        includeFuncHostedBot && programmingLanguage === ProgrammingLanguage.typescript
          ? hasTeamsFxDevScript
            ? this.prepareTaskNext(
                TaskDefinition.funcHostedBotWatch(workspaceFolder),
                constants.botWatchStartMessage,
                true
              )
            : this.prepareTask(
                TaskDefinition.funcHostedBotWatch(workspaceFolder),
                constants.botWatchStartMessage,
                commonUtils.getBotLocalEnv(localEnv)
              )
          : undefined;
      const botStartTask = includeFuncHostedBot
        ? // For func hosted bot, always use the new task (prepareTaskNext).
          this.prepareTaskNext(
            TaskDefinition.funcHostedBotStart(workspaceFolder),
            constants.botStartStartMessageNext,
            false,
            funcEnv
          )
        : hasTeamsFxDevScript
        ? this.prepareTaskNext(
            TaskDefinition.botStart(workspaceFolder, programmingLanguage, true),
            constants.botStartStartMessageNext,
            false
          )
        : this.prepareTask(
            TaskDefinition.botStart(workspaceFolder, programmingLanguage, true),
            constants.botStartStartMessage,
            commonUtils.getBotLocalEnv(localEnv)
          );

      const botAzuriteTask = this.prepareTask(
        TaskDefinition.funcHostedBotAzurite(workspaceFolder),
        constants.botWatchStartMessage
      );

      botTaskPromises = [
        Preview.sequentialTasks(
          () =>
            botWatchTask?.task.waitFor(
              constants.funcHostedBotWatchPattern,
              botWatchTask.startCb,
              botWatchTask.stopCb,
              undefined,
              this.serviceLogWriter
            ),
          () =>
            botStartTask?.task.waitFor(
              includeFuncHostedBot
                ? constants.funcHostedBotStartPattern
                : constants.botStartPattern,
              botStartTask.startCb,
              botStartTask.stopCb,
              30000,
              this.serviceLogWriter
            )
        ),
        includeFuncHostedBot
          ? botAzuriteTask?.task.waitFor(
              constants.funcHostedBotAzuritePattern,
              botAzuriteTask?.startCb,
              botAzuriteTask?.stopCb,
              30000,
              this.serviceLogWriter
            )
          : undefined,
      ];
    }

    return botTaskPromises;
  }

  private async startServices(
    workspaceFolder: string,
    programmingLanguage: string,
    includeFrontend: boolean,
    includeBackend: boolean,
    includeBot: boolean,
    includeFuncHostedBot: boolean,
    depsManager: DepsManager,
    includeSimpleAuth?: boolean
  ): Promise<Result<null, FxError>> {
    const localEnv = await commonUtils.getLocalEnv(workspaceFolder);

    const frontendStartTask = includeFrontend
      ? (await loadTeamsFxDevScript(path.join(workspaceFolder, FolderName.Frontend))) !== undefined
        ? this.prepareTaskNext(
            TaskDefinition.frontendStart(workspaceFolder),
            constants.frontendStartStartMessageNext,
            false
          )
        : this.prepareTask(
            TaskDefinition.frontendStart(workspaceFolder),
            constants.frontendStartStartMessage,
            commonUtils.getFrontendLocalEnv(localEnv)
          )
      : undefined;

    const dotnet = (await depsManager.getStatus([DepsType.Dotnet]))[0];
    const authStartTask =
      includeFrontend && includeSimpleAuth
        ? this.prepareTask(
            TaskDefinition.authStart(dotnet.command, commonUtils.getAuthServicePath(localEnv)),
            constants.authStartStartMessage,
            commonUtils.getAuthLocalEnv(localEnv)
          )
        : undefined;

    const func = (await depsManager.getStatus([DepsType.FuncCoreTools]))[0];
    const funcCommand = func.command;
    let funcEnv = undefined;
    if (func.details.binFolders !== undefined) {
      funcEnv = {
        PATH: `${process.env.PATH}${path.delimiter}${func.details.binFolders.join(path.delimiter)}`,
      };
    }
    const backendStartTask = includeBackend
      ? (await loadTeamsFxDevScript(path.join(workspaceFolder, FolderName.Function))) !== undefined
        ? this.prepareTaskNext(
            TaskDefinition.backendStart(workspaceFolder, programmingLanguage, funcCommand, true),
            constants.backendStartStartMessageNext,
            false,
            funcEnv
          )
        : this.prepareTask(
            TaskDefinition.backendStart(workspaceFolder, programmingLanguage, funcCommand, true),
            constants.backendStartStartMessage,
            commonUtils.getBackendLocalEnv(localEnv)
          )
      : undefined;
    const backendWatchTask =
      includeBackend && programmingLanguage === ProgrammingLanguage.typescript
        ? (await loadTeamsFxDevScript(path.join(workspaceFolder, FolderName.Function))) !==
          undefined
          ? this.prepareTaskNext(
              TaskDefinition.backendWatch(workspaceFolder),
              constants.backendWatchStartMessageNext,
              true
            )
          : this.prepareTask(
              TaskDefinition.backendWatch(workspaceFolder),
              constants.backendWatchStartMessage,
              commonUtils.getBackendLocalEnv(localEnv)
            )
        : undefined;

    // For TypeScript projects, backendStart depends on backendWatch.
    // backendStart runs `func start ...` which uses `bot/{funcName}/function.json`,
    //  which refers to the JavaScript files built from TypeScript files.
    // As a result, running backendStart before backendWatch succeeds will result in JavaScript file not found error.
    const backendTaskPromise = Preview.sequentialTasks(
      () =>
        backendStartTask?.task.waitFor(
          constants.backendStartPattern,
          backendStartTask.startCb,
          backendStartTask.stopCb,
          undefined,
          this.serviceLogWriter
        ),
      () =>
        backendWatchTask?.task.waitFor(
          constants.backendWatchPattern,
          backendWatchTask.startCb,
          backendWatchTask.stopCb,
          undefined,
          this.serviceLogWriter
        )
    );

    const botTaskPromises = await this.createBotTasksForStartServices(
      workspaceFolder,
      programmingLanguage,
      includeBot,
      includeFuncHostedBot,
      localEnv,
      funcEnv
    );

    const results = await Promise.all([
      frontendStartTask?.task.waitFor(
        constants.frontendStartPattern,
        frontendStartTask.startCb,
        frontendStartTask.stopCb,
        undefined,
        this.serviceLogWriter
      ),
      authStartTask?.task.waitFor(
        constants.authStartPattern,
        authStartTask.startCb,
        authStartTask.stopCb,
        undefined,
        this.serviceLogWriter
      ),
      backendTaskPromise,
      ...botTaskPromises,
    ]);
    const fxErrors: FxError[] = [];
    for (const result of results) {
      if (result?.isErr()) {
        fxErrors.push(result.error);
      }
    }
    if (fxErrors.length > 0) {
      return err(errors.PreviewCommandFailed(fxErrors));
    }
    return ok(null);
  }

  private async openTeamsWebClient(
    tenantIdFromConfig: string | undefined,
    teamsAppId: string,
    browser: constants.Browser,
    browserArguments: string[] = []
  ): Promise<Result<null, FxError>> {
    cliTelemetry.sendTelemetryEvent(
      TelemetryEvent.PreviewSideloadingStart,
      this.telemetryProperties
    );

    let sideloadingUrl = constants.sideloadingUrl.replace(
      constants.teamsAppIdPlaceholder,
      teamsAppId
    );

    let tenantId, loginHint: string | undefined;
    try {
      const tokenObject = (await AppStudioTokenInstance.getStatus())?.accountInfo;
      if (tokenObject) {
        // user signed in
        tenantId = tokenObject.tid as string;
        loginHint = tokenObject.upn as string;
      } else {
        // no signed user
        tenantId = tenantIdFromConfig;
        loginHint = "login_your_m365_account"; // a workaround that user has the chance to login
      }
    } catch {
      // ignore error
    }

    if (tenantId && loginHint) {
      sideloadingUrl = sideloadingUrl.replace(
        constants.accountHintPlaceholder,
        `appTenantId=${tenantId}&login_hint=${loginHint}`
      );
    } else {
      sideloadingUrl = sideloadingUrl.replace(constants.accountHintPlaceholder, "");
    }

    const previewBar = CLIUIInstance.createProgressBar(constants.previewTitle, 1);
    await previewBar.start(constants.previewStartMessage);
    await previewBar.next(constants.previewStartMessage);
    try {
      await commonUtils.openBrowser(browser, sideloadingUrl, browserArguments);
    } catch {
      const error = errors.OpeningBrowserFailed(browser);
      cliTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.PreviewSideloading,
        error,
        this.telemetryProperties
      );
      cliLogger.necessaryLog(LogLevel.Warning, constants.openBrowserHintMessage);
      await previewBar.end(false);
      return ok(null);
    }
    await previewBar.end(true);
    const message = [
      {
        content: `preview url: `,
        color: Colors.WHITE,
      },
      {
        content: sideloadingUrl,
        color: Colors.BRIGHT_CYAN,
      },
    ];
    cliLogger.necessaryLog(LogLevel.Info, utils.getColorizedString(message));

    cliTelemetry.sendTelemetryEvent(TelemetryEvent.PreviewSideloading, {
      ...this.telemetryProperties,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
    return ok(null);
  }

  private async terminateTasks(): Promise<void> {
    for (const task of this.backgroundTasks) {
      await task.terminate();
    }
    this.backgroundTasks = [];
  }

  private async handleDependences(
    projectSettings: ProjectSettings,
    localEnvManager: LocalEnvManager,
    depsManager: DepsManager
  ): Promise<Result<null, FxError>> {
    let shouldContinue = true;
    const availableDeps = localEnvManager.getActiveDependencies(projectSettings);
    const enabledDeps = await CliDepsChecker.getEnabledDeps(
      availableDeps.filter((dep) => !CliDepsChecker.getNodeDeps().includes(dep))
    );

    for (const dep of enabledDeps) {
      const bar = CLIUIInstance.createProgressBar(DepsDisplayName[dep], 1);
      await bar.start(ProgressMessage[dep]);
      await bar.next(ProgressMessage[dep]);
      const depStatus = (
        await depsManager.ensureDependencies([dep], {
          fastFail: false,
          doctor: true,
        })
      )[0];

      let result;
      let summaryMsg;

      if (depStatus.isInstalled) {
        result = true;
        summaryMsg = depStatus.details.binFolders
          ? `${depStatus.name} (installed at ${depStatus.details.binFolders?.[0]})`
          : DepsDisplayName[dep];
      } else {
        result = false;
        summaryMsg = depStatus.error ? depStatus.error.message : DepsDisplayName[dep];
      }
      shouldContinue = shouldContinue && result;
      await bar.next(summaryMsg);
      await bar.end(result);
      if (!result && depStatus.error && depStatus.error.helpLink) {
        cliLogger.necessaryLog(
          LogLevel.Info,
          doctorResult.HelpLink.split("@Link").join(depStatus.error.helpLink)
        );
      }
    }

    if (!shouldContinue) {
      return err(errors.DependencyCheckerFailed());
    }

    return ok(null);
  }

  private async checkNode(
    hasBackend: boolean,
    hasFuncHostedBot: boolean,
    depsManager: DepsManager
  ): Promise<Result<null, FxError>> {
    const node = hasBackend || hasFuncHostedBot ? DepsType.FunctionNode : DepsType.AzureNode;
    const nodeBar = CLIUIInstance.createProgressBar(DepsDisplayName[node], 1);
    await nodeBar.start(ProgressMessage[node]);

    let nodeStatus;
    let result = true;
    let summaryMsg = doctorResult.NodeSuccess;
    let helpLink = undefined;

    try {
      nodeStatus = (
        await depsManager.ensureDependencies([node], {
          fastFail: false,
          doctor: true,
        })
      )[0];

      if (!nodeStatus.isInstalled) {
        summaryMsg = doctorResult.NodeNotFound;
        result = false;
        if (nodeStatus.error) {
          helpLink = nodeStatus.error.helpLink;
        }
        if (nodeStatus.error instanceof NodeNotSupportedError) {
          const supportedVersions = nodeStatus?.details.supportedVersions
            .map((v) => "v" + v)
            .join(" ,");
          summaryMsg = doctorResult.NodeNotSupported.split("@CurrentVersion")
            .join(nodeStatus?.details.installVersion)
            .split("@SupportedVersions")
            .join(supportedVersions);
        }
      }
    } catch (err) {
      result = false;
      summaryMsg = doctorResult.NodeNotFound;
    }

    await nodeBar.next(summaryMsg);
    await nodeBar.end(result);
    if (!result) {
      cliLogger.necessaryLog(LogLevel.Info, doctorResult.InstallNode);
      return err(errors.PrerequisitesValidationError("Node.js checker failed.", helpLink));
    }

    return ok(null);
  }

  private async checkM365Account(): Promise<Result<null, FxError>> {
    let result = true;
    let summaryMsg = `${Checker.M365Account}`;
    let error = undefined;
    const accountBar = CLIUIInstance.createProgressBar(Checker.M365Account, 1);
    await accountBar.start(ProgressMessage[Checker.M365Account]);
    await accountBar.next(ProgressMessage[Checker.M365Account]);
    let loginHint = undefined;
    try {
      const loginStatus = await AppStudioTokenInstance.getStatus();
      let token = loginStatus.token;
      if (loginStatus.status === signedOut) {
        token = await AppStudioTokenInstance.getAccessToken(true);
      }

      if (token === undefined) {
        result = false;
        summaryMsg = doctorResult.NotSignIn;
      } else {
        const isSideloadingEnabled = await getSideloadingStatus(token);
        if (isSideloadingEnabled === false) {
          // sideloading disabled
          result = false;
          summaryMsg = doctorResult.SideLoadingDisabled;
        }
      }

      const tokenObject = loginStatus.accountInfo;
      if (tokenObject && tokenObject.upn) {
        loginHint = tokenObject.upn;
      }
    } catch (err: any) {
      result = false;
      error = this.assembleError(err, cliSource);
    }

    if (result && loginHint) {
      summaryMsg = doctorResult.SignInSuccess.split("@account").join(`${loginHint}`);
    }
    await accountBar.next(summaryMsg);
    await accountBar.end(result);

    if (!result) {
      return error ? err(error) : err(errors.PrerequisitesValidationError(summaryMsg));
    }
    return ok(null);
  }

  private async resolveLocalCertificate(
    localEnvManager: LocalEnvManager
  ): Promise<Result<null, FxError>> {
    let result = true;
    let summaryMsg;
    let error = undefined;
    const certBar = CLIUIInstance.createProgressBar(Checker.LocalCertificate, 1);
    await certBar.start(ProgressMessage[Checker.LocalCertificate]);
    await certBar.next(ProgressMessage[Checker.LocalCertificate]);
    try {
      const trustDevCert = await isTrustDevCertEnabled();
      const localCertResult = await localEnvManager.resolveLocalCertificate(trustDevCert);
      if (localCertResult.isTrusted === false) {
        result = false;
        error = localCertResult.error;
      } else if (typeof localCertResult.isTrusted === "undefined") {
        summaryMsg = doctorResult.SkipTrustingCert;
      }
    } catch (err: any) {
      result = false;
      error = assembleError(err);
    }

    await certBar.next(summaryMsg);
    await certBar.end(result);
    if (!result && error) {
      return err(error);
    }
    return ok(null);
  }

  private prepareTask(
    taskDefinition: ITaskDefinition,
    startMessage: string,
    env?: { [key: string]: string }
  ): {
    task: Task;
    startCb: (taskTitle: string, background: boolean) => Promise<void>;
    stopCb: (
      taskTitle: string,
      background: boolean,
      result: TaskResult,
      serviceLogWriter?: ServiceLogWriter
    ) => Promise<FxError | null>;
  } {
    const taskEnv = env ?? taskDefinition.env;
    const task = new Task(
      taskDefinition.name,
      taskDefinition.isBackground,
      taskDefinition.command,
      taskDefinition.args,
      {
        shell: taskDefinition.execOptions.needCmd
          ? "cmd.exe"
          : taskDefinition.execOptions.needShell,
        cwd: taskDefinition.cwd,
        env: taskEnv ? commonUtils.mergeProcessEnv(taskEnv) : undefined,
      }
    );
    const bar = CLIUIInstance.createProgressBar(taskDefinition.name, 1);
    const startCb = commonUtils.createTaskStartCb(bar, startMessage, this.telemetryProperties);
    const stopCb = commonUtils.createTaskStopCb(bar, this.telemetryProperties);
    if (taskDefinition.isBackground) {
      this.backgroundTasks.push(task);
    }
    return { task: task, startCb: startCb, stopCb: stopCb };
  }

  private prepareTaskNext(
    taskDefinition: ITaskDefinition,
    startMessage: string,
    isWatchTask: boolean,
    env?: { [key: string]: string }
  ): {
    task: Task;
    startCb: (taskTitle: string, background: boolean) => Promise<void>;
    stopCb: (
      taskTitle: string,
      background: boolean,
      result: TaskResult,
      serviceLogWriter?: ServiceLogWriter
    ) => Promise<FxError | null>;
  } {
    const taskEnv = env ?? taskDefinition.env;
    const task = new Task(
      taskDefinition.name,
      taskDefinition.isBackground,
      isWatchTask ? "npm run watch:teamsfx" : "npm run dev:teamsfx",
      taskDefinition.args,
      {
        shell: taskDefinition.execOptions.needCmd
          ? "cmd.exe"
          : taskDefinition.execOptions.needShell,
        cwd: taskDefinition.cwd,
        env: taskEnv ? commonUtils.mergeProcessEnv(taskEnv) : undefined,
      }
    );
    const bar = CLIUIInstance.createProgressBar(taskDefinition.name, 1);
    const startCb = commonUtils.createTaskStartCb(bar, startMessage, this.telemetryProperties);
    const stopCb = commonUtils.createTaskStopCb(bar, this.telemetryProperties);
    if (taskDefinition.isBackground) {
      this.backgroundTasks.push(task);
    }
    return { task: task, startCb: startCb, stopCb: stopCb };
  }

  private assembleError(e: any, source?: string): FxError {
    if (e instanceof UserError || e instanceof SystemError) return e;
    if (!source) source = "unknown";
    const type = typeof e;
    if (type === "string") {
      return new UnknownError(source, e as string);
    } else if (e instanceof Error) {
      const err = e as Error;
      const fxError = new SystemError(err, source);
      fxError.stack = err.stack;
      return fxError;
    } else {
      return new UnknownError(source, JSON.stringify(e));
    }
  }
}
