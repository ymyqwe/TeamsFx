import {
  AzureSolutionSettings,
  combine,
  err,
  Func,
  FxError,
  Inputs,
  Json,
  ok,
  Platform,
  ProjectSettings,
  Result,
  returnSystemError,
  returnUserError,
  SolutionSettings,
  SystemError,
  TelemetryReporter,
  TokenProvider,
  UserError,
  UserInteraction,
  v2,
  v3,
  Void,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { cloneDeep } from "lodash";
import { Container } from "typedi";
import * as util from "util";
import { isAADEnabled, isAadManifestEnabled } from "../../../../common";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { isVSProject } from "../../../../common/projectSettingsHelper";
import { OperationNotPermittedError } from "../../../../core/error";
import { CoreQuestionNames } from "../../../../core/question";
import { AppStudioPluginV3 } from "../../../resource/appstudio/v3";
import {
  DEFAULT_PERMISSION_REQUEST,
  PluginNames,
  SolutionError,
  SolutionSource,
  SolutionTelemetryComponentName,
  SolutionTelemetryEvent,
  SolutionTelemetryProperty,
  SolutionTelemetrySuccess,
  SOLUTION_PROVISION_SUCCEEDED,
} from "../constants";
import { scaffoldLocalDebugSettings } from "../debug/scaffolding";
import {
  AzureResourceApim,
  AzureResourceFunction,
  AzureResourceKeyVault,
  AzureResourceSQL,
  AzureSolutionQuestionNames,
  BotOptionItem,
  HostTypeOptionAzure,
  MessageExtensionItem,
  SsoItem,
  TabOptionItem,
} from "../question";
import { getAllV2ResourcePluginMap, ResourcePluginsV2 } from "../ResourcePluginContainer";
import { sendErrorTelemetryThenReturnError } from "../utils/util";
import { BuiltInFeaturePluginNames } from "../v3/constants";
import { TeamsAppSolutionNameV2 } from "./constants";
import { generateResourceTemplateForPlugins } from "./generateResourceTemplate";
import { scaffoldByPlugins } from "./scaffolding";
import { getAzureSolutionSettings, setActivatedResourcePluginsV2 } from "./utils";
export async function executeUserTask(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  localSettings: Json,
  envInfo: v2.EnvInfoV2,
  tokenProvider: TokenProvider
): Promise<Result<unknown, FxError>> {
  const namespace = func.namespace;
  const method = func.method;
  const array = namespace.split("/");
  if (method === "addCapability") {
    return addCapability(ctx, inputs, localSettings);
  }
  if (method === "addResource") {
    return addResource(ctx, inputs, localSettings, func, envInfo, tokenProvider);
  }
  if (method === "addSso") {
    return addSso(ctx, inputs, localSettings);
  }
  if (namespace.includes("solution")) {
    if (method === "registerTeamsAppAndAad") {
      // not implemented for now
      return err(
        returnSystemError(
          new Error("Not implemented"),
          SolutionSource,
          SolutionError.FeatureNotSupported
        )
      );
    } else if (method === "VSpublish") {
      // VSpublish means VS calling cli to do publish. It is different than normal cli work flow
      // It's teamsfx init followed by teamsfx  publish without running provision.
      // Using executeUserTask here could bypass the fx project check.
      if (inputs.platform !== "vs") {
        return err(
          returnSystemError(
            new Error(`VS publish is not supposed to run on platform ${inputs.platform}`),
            SolutionSource,
            SolutionError.UnsupportedPlatform
          )
        );
      }
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.publishApplication) {
        return appStudioPlugin.publishApplication(
          ctx,
          inputs,
          envInfo,
          tokenProvider.appStudioToken
        );
      }
    } else if (method === "validateManifest") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return await appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (method === "buildPackage") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return await appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (method === "validateManifest") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (array.length == 2) {
      const pluginName = array[1];
      const pluginMap = getAllV2ResourcePluginMap();
      const plugin = pluginMap.get(pluginName);
      if (plugin && plugin.executeUserTask) {
        return plugin.executeUserTask(ctx, inputs, func, localSettings, envInfo, tokenProvider);
      }
    }
  }

  return err(
    returnUserError(
      new Error(`executeUserTaskRouteFailed:${JSON.stringify(func)}`),
      SolutionSource,
      `executeUserTaskRouteFailed`
    )
  );
}

export function canAddCapability(
  settings: AzureSolutionSettings | undefined,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  if (settings && !(settings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionError.AddCapabilityNotSupport,
      getLocalizedString("core.addCapability.onlySupportAzure"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddCapability, e, telemetryReporter)
    );
  }
  return ok(Void);
}

export function canAddResource(
  projectSetting: ProjectSettings,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  const isVS = isVSProject(projectSetting);
  if (isVS) {
    const e = new UserError(
      SolutionError.AddResourceNotSupport,
      getLocalizedString("core.addResource.notSupportForVSProject"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddResource, e, telemetryReporter)
    );
  }
  const solutionSettings = projectSetting.solutionSettings as AzureSolutionSettings;
  if (!(solutionSettings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionError.AddResourceNotSupport,
      getLocalizedString("core.addResource.onlySupportAzure"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddResource, e, telemetryReporter)
    );
  }
  return ok(Void);
}

export async function addCapability(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json
): Promise<
  Result<{ solutionSettings?: SolutionSettings; solutionConfig?: Record<string, unknown> }, FxError>
> {
  ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddCapabilityStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  // 0. set programming language if it is empty
  const programmingLanguageInputs = inputs[CoreQuestionNames.ProgrammingLanguage];
  if (!ctx.projectSetting.programmingLanguage && programmingLanguageInputs) {
    ctx.projectSetting.programmingLanguage = programmingLanguageInputs;
  }

  // 1. checking addable
  let solutionSettings = getAzureSolutionSettings(ctx);
  if (!solutionSettings) {
    // pure existing app
    solutionSettings = {
      name: TeamsAppSolutionNameV2,
      version: "1.0.0",
      hostType: "Azure",
      capabilities: [],
      azureResources: [],
      activeResourcePlugins: [],
    };
    ctx.projectSetting.solutionSettings = solutionSettings;
    if (isAADEnabled(solutionSettings)) {
      //aad need this file
      await fs.writeJSON(`${inputs.projectPath}/permissions.json`, DEFAULT_PERMISSION_REQUEST, {
        spaces: 4,
      });
    }
  }
  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
    existingCapabilities: originalSettings.capabilities,
  };
  const canProceed = canAddCapability(solutionSettings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // 2. check answer
  const capabilitiesAnswer = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
  if (!capabilitiesAnswer || capabilitiesAnswer.length === 0) {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: [].join(";"),
    });
    return ok({});
  }

  // 3. check capability limit
  const alreadyHasTab = solutionSettings.capabilities.includes(TabOptionItem.id);
  const alreadyHasBot = solutionSettings.capabilities.includes(BotOptionItem.id);
  const alreadyHasME = solutionSettings.capabilities.includes(MessageExtensionItem.id);
  const toAddTab = capabilitiesAnswer.includes(TabOptionItem.id);
  const toAddBot = capabilitiesAnswer.includes(BotOptionItem.id);
  const toAddME = capabilitiesAnswer.includes(MessageExtensionItem.id);
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  const inputsWithProjectPath = inputs as v2.InputsWithProjectPath;
  const tabExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "staticTab"
  );
  if (tabExceedRes.isErr()) {
    return err(tabExceedRes.error);
  }
  const isTabAddable = !tabExceedRes.value;
  const botExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "Bot"
  );
  if (botExceedRes.isErr()) {
    return err(botExceedRes.error);
  }
  const isBotAddable = !botExceedRes.value;
  const meExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "MessageExtension"
  );
  if (meExceedRes.isErr()) {
    return err(meExceedRes.error);
  }
  const isMEAddable = !meExceedRes.value;
  if ((toAddTab && !isTabAddable) || (toAddBot && !isBotAddable) || (toAddME && !isMEAddable)) {
    const error = new UserError(
      SolutionError.FailedToAddCapability,
      getLocalizedString("core.addCapability.exceedMaxLimit"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddCapability,
        error,
        ctx.telemetryReporter
      )
    );
  }

  const capabilitiesToAddManifest: v3.ManifestCapability[] = [];
  const pluginNamesToScaffold: Set<string> = new Set<string>();
  const pluginNamesToArm: Set<string> = new Set<string>();
  const newCapabilitySet = new Set<string>();
  solutionSettings.capabilities.forEach((c) => newCapabilitySet.add(c));
  const vsProject = isVSProject(ctx.projectSetting);
  if (!originalSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.identity)) {
    pluginNamesToArm.add(ResourcePluginsV2.IdentityPlugin);
  }
  if (
    !isAadManifestEnabled() &&
    !originalSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.aad)
  ) {
    pluginNamesToArm.add(ResourcePluginsV2.AadPlugin);
  }

  // 4. check Tab
  if (capabilitiesAnswer.includes(TabOptionItem.id)) {
    if (vsProject) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (!alreadyHasTab) {
        pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
      }
    } else {
      if (!alreadyHasTab) {
        pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "staticTab" });
    newCapabilitySet.add(TabOptionItem.id);
  }
  // 5. check Bot
  if (capabilitiesAnswer.includes(BotOptionItem.id)) {
    if (vsProject) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (!alreadyHasBot && !alreadyHasME) {
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    } else {
      if (!alreadyHasBot && !alreadyHasME) {
        pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "Bot" });
    newCapabilitySet.add(BotOptionItem.id);
  }
  // 6. check MessageExtension
  if (capabilitiesAnswer.includes(MessageExtensionItem.id)) {
    if (vsProject) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (!alreadyHasBot && !alreadyHasME) {
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    } else {
      if (!alreadyHasBot && !alreadyHasME) {
        pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "MessageExtension" });
    newCapabilitySet.add(MessageExtensionItem.id);
  }

  // 7. update solution settings
  solutionSettings.capabilities = Array.from(newCapabilitySet);
  setActivatedResourcePluginsV2(ctx.projectSetting);

  if (
    !isAadManifestEnabled() &&
    !solutionSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.aad)
  ) {
    solutionSettings.activeResourcePlugins.push(BuiltInFeaturePluginNames.aad);
  }

  // 8. scaffold and update arm
  const pluginsToScaffold = Array.from(pluginNamesToScaffold).map((name) =>
    Container.get<v2.ResourcePlugin>(name)
  );
  const pluginsToArm = Array.from(pluginNamesToArm).map((name) =>
    Container.get<v2.ResourcePlugin>(name)
  );
  if (pluginsToScaffold.length > 0) {
    const scaffoldRes = await scaffoldCodeAndResourceTemplate(
      ctx,
      inputsNew,
      localSettings,
      pluginsToScaffold,
      pluginsToArm
    );
    if (scaffoldRes.isErr()) {
      ctx.projectSetting.solutionSettings = originalSettings;
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddCapability,
          scaffoldRes.error,
          ctx.telemetryReporter
        )
      );
    }
  }
  // 4. update manifest
  if (capabilitiesToAddManifest.length > 0 || pluginsToScaffold.length > 0) {
    await appStudioPlugin.addCapabilities(ctx, inputsWithProjectPath, capabilitiesToAddManifest);
  }
  if (capabilitiesAnswer.length > 0) {
    const addNames = capabilitiesAnswer.map((c) => `'${c}'`).join(" and ");
    const single = capabilitiesAnswer.length === 1;
    const template =
      inputs.platform === Platform.CLI
        ? single
          ? getLocalizedString("core.addCapability.addCapabilityNoticeForCli")
          : getLocalizedString("core.addCapability.addCapabilitiesNoticeForCli")
        : single
        ? getLocalizedString("core.addCapability.addCapabilityNotice")
        : getLocalizedString("core.addCapability.addCapabilitiesNotice");
    const msg = util.format(template, addNames);
    ctx.userInteraction.showMessage("info", msg, false);
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: capabilitiesAnswer.join(";"),
    });
  }
  return ok({
    solutionSettings: solutionSettings,
    solutionConfig: { provisionSucceeded: false },
  });
}

export function showUpdateArmTemplateNotice(ui?: UserInteraction) {
  const msg: string = getLocalizedString("core.updateArmTemplate.successNotice");
  ui?.showMessage("info", msg, false);
}

async function scaffoldCodeAndResourceTemplate(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  pluginsToScaffold: v2.ResourcePlugin[],
  pluginsToDoArm?: v2.ResourcePlugin[]
): Promise<Result<unknown, FxError>> {
  const result = await scaffoldByPlugins(ctx, inputs, localSettings, pluginsToScaffold);
  if (result.isErr()) {
    return result;
  }
  const scaffoldLocalDebugSettingsResult = await scaffoldLocalDebugSettings(
    ctx,
    inputs,
    localSettings
  );
  if (scaffoldLocalDebugSettingsResult.isErr()) {
    return scaffoldLocalDebugSettingsResult;
  }
  const pluginsToUpdateArm = pluginsToDoArm ? pluginsToDoArm : pluginsToScaffold;
  if (pluginsToUpdateArm.length > 0) {
    return generateResourceTemplateForPlugins(ctx, inputs, pluginsToUpdateArm);
  }
  return ok(undefined);
}

export async function addResource(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  func: Func,
  envInfo: v2.EnvInfoV2,
  tokenProvider: TokenProvider
): Promise<Result<unknown, FxError>> {
  ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResourceStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  // 1. checking addable
  const solutionSettings = getAzureSolutionSettings(ctx);
  if (!solutionSettings) {
    return err(new OperationNotPermittedError("addResource"));
  }
  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew: v2.InputsWithProjectPath & { existingResources: string[] } = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
  };
  const canProceed = canAddResource(ctx.projectSetting, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // 2. check answer
  const addResourcesAnswer = inputs[AzureSolutionQuestionNames.AddResources] as string[];
  if (!addResourcesAnswer || addResourcesAnswer.length === 0) {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Resources]: [].join(";"),
    });
    return ok({});
  }

  const alreadyHaveFunction = solutionSettings.azureResources.includes(AzureResourceFunction.id);
  const alreadyHaveApim = solutionSettings.azureResources.includes(AzureResourceApim.id);
  const alreadyHaveKeyVault = solutionSettings.azureResources.includes(AzureResourceKeyVault.id);
  const addSQL = addResourcesAnswer.includes(AzureResourceSQL.id);
  const addApim = addResourcesAnswer.includes(AzureResourceApim.id);
  const addKeyVault = addResourcesAnswer.includes(AzureResourceKeyVault.id);
  const addFunc =
    addResourcesAnswer.includes(AzureResourceFunction.id) || (addApim && !alreadyHaveFunction);

  // 3. check APIM and KeyVault addable
  if ((alreadyHaveApim && addApim) || (alreadyHaveKeyVault && addKeyVault)) {
    const e = new UserError(
      new Error("APIM/KeyVault is already added."),
      SolutionSource,
      SolutionError.AddResourceNotSupport
    );
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddResource,
        e,
        ctx.telemetryReporter
      )
    );
  }

  const newResourceSet = new Set<string>();
  solutionSettings.azureResources.forEach((r) => newResourceSet.add(r));
  const addedResources: string[] = [];
  const pluginsToScaffold: v2.ResourcePlugin[] = [];
  const pluginsToDoArm: v2.ResourcePlugin[] = [];
  let scaffoldApim = false;
  // 4. check Function
  if (addFunc) {
    // AAD plugin needs to be activated when adding function.
    // Since APIM also have dependency on Function, will only add depenedency here.
    if (!isAADEnabled(solutionSettings)) {
      if (isAadManifestEnabled()) {
        const aadPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AadPlugin);
        pluginsToScaffold.push(aadPlugin);
        pluginsToDoArm.push(aadPlugin);

        solutionSettings.capabilities.push(SsoItem.id);
      } else {
        solutionSettings.activeResourcePlugins?.push(PluginNames.AAD);
      }
    }
    const functionPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.FunctionPlugin);
    pluginsToScaffold.push(functionPlugin);
    if (!alreadyHaveFunction) {
      pluginsToDoArm.push(functionPlugin);
    }
    addedResources.push(AzureResourceFunction.id);
  }
  // 5. check SQL
  if (addSQL) {
    const sqlPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.SqlPlugin);
    const identityPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.IdentityPlugin);
    pluginsToDoArm.push(sqlPlugin);
    if (!solutionSettings.activeResourcePlugins.includes(identityPlugin.name)) {
      // add identity for first time
      pluginsToDoArm.push(identityPlugin);
    }
    addedResources.push(AzureResourceSQL.id);
  }
  // 6. check APIM
  const apimPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.ApimPlugin);
  if (addApim) {
    // We don't add apimPlugin into pluginsToScaffold because
    // apim plugin needs to modify config output during scaffolding,
    // which is not supported by the scaffoldSourceCode API.
    // The scaffolding will run later as a userTask as a work around.
    addedResources.push(AzureResourceApim.id);
    pluginsToDoArm.push(apimPlugin);
    scaffoldApim = true;
  }
  if (addKeyVault) {
    const keyVaultPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.KeyVaultPlugin);
    pluginsToDoArm.push(keyVaultPlugin);
    addedResources.push(AzureResourceKeyVault.id);
  }

  // 7. update solution settings
  addedResources.forEach((r) => newResourceSet.add(r));
  solutionSettings.azureResources = Array.from(newResourceSet);
  setActivatedResourcePluginsV2(ctx.projectSetting);

  // 8. scaffold and update arm
  if (pluginsToScaffold.length > 0 || pluginsToDoArm.length > 0) {
    let scaffoldRes = await scaffoldCodeAndResourceTemplate(
      ctx,
      inputsNew,
      localSettings,
      pluginsToScaffold,
      pluginsToDoArm
    );
    if (scaffoldApim) {
      if (apimPlugin && apimPlugin.executeUserTask) {
        const result = await apimPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          {},
          envInfo,
          tokenProvider
        );
        if (result.isErr()) {
          scaffoldRes = combine([scaffoldRes, result]);
        }
      }
    }
    if (scaffoldRes.isErr()) {
      ctx.projectSetting.solutionSettings = originalSettings;
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddResource,
          scaffoldRes.error,
          ctx.telemetryReporter
        )
      );
    }
    const addNames = addedResources.map((c) => `'${c}'`).join(" and ");
    const single = addedResources.length === 1;
    const template =
      inputs.platform === Platform.CLI
        ? single
          ? getLocalizedString("core.addResource.addResourceNoticeForCli")
          : getLocalizedString("core.addResource.addResourcesNoticeForCli")
        : single
        ? getLocalizedString("core.addResource.addResourceNotice")
        : getLocalizedString("core.addResource.addResourcesNotice");
    ctx.userInteraction.showMessage("info", util.format(template, addNames), false);
  }

  ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
    [SolutionTelemetryProperty.Resources]: addResourcesAnswer.join(";"),
  });
  return ok(
    pluginsToDoArm.length > 0
      ? { solutionSettings: solutionSettings, solutionConfig: { provisionSucceeded: false } }
      : Void
  );
}

export function extractParamForRegisterTeamsAppAndAad(
  answers?: Inputs
): Result<ParamForRegisterTeamsAppAndAad, FxError> {
  if (answers == undefined) {
    return err(
      returnSystemError(
        new Error("Input is undefined"),
        SolutionSource,
        SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
      )
    );
  }

  const param: ParamForRegisterTeamsAppAndAad = {
    "app-name": "",
    endpoint: "",
    environment: "local",
    "root-path": "",
  };
  for (const key of Object.keys(param)) {
    const value = answers[key];
    if (value == undefined) {
      return err(
        returnSystemError(
          new Error(`${key} not found`),
          SolutionSource,
          SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
        )
      );
    }
    (param as any)[key] = value;
  }

  return ok(param);
}

export type ParamForRegisterTeamsAppAndAad = {
  "app-name": string;
  environment: "local" | "remote";
  endpoint: string;
  "root-path": string;
};

// TODO: handle VS scenario
export function canAddSso(
  solutionSettings: AzureSolutionSettings,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  // Can not add sso if feature flag is not enabled
  if (!isAadManifestEnabled()) {
    const e = new SystemError(
      SolutionError.NeedEnableFeatureFlag,
      getLocalizedString("core.addSso.needEnableFeatureFlag"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddSso, e, telemetryReporter)
    );
  }

  if (!(solutionSettings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionError.AddSsoNotSupported,
      getLocalizedString("core.addSso.onlySupportAzure"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddSso, e, telemetryReporter)
    );
  }

  // Can only add sso when capability includes Tab, Bot, Messaging Extension, etc.
  if (
    !solutionSettings.capabilities.includes(TabOptionItem.id) &&
    !solutionSettings.capabilities.includes(BotOptionItem.id) &&
    !solutionSettings.capabilities.includes(MessageExtensionItem.id)
  ) {
    const e = new UserError(
      SolutionError.AddSsoNotSupported,
      getLocalizedString("core.addSso.needCapability"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddSso, e, telemetryReporter)
    );
  }

  // Check whether SSO is enabled
  const activeResourcePlugins = solutionSettings.activeResourcePlugins;
  const containSsoItem = solutionSettings.capabilities.includes(SsoItem.id);
  const containAadPlugin = activeResourcePlugins.includes(PluginNames.AAD);
  if (containSsoItem && containAadPlugin) {
    // Throw error if sso is already enabled
    const e = new UserError(
      SolutionError.SsoEnabled,
      getLocalizedString("core.addSso.ssoEnabled"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddSso, e, telemetryReporter)
    );
  } else if (containSsoItem || containAadPlugin) {
    // Throw error if the project is invalid
    const e = new UserError(
      SolutionError.InvalidSsoProject,
      getLocalizedString("core.addSso.invalidSsoProject"),
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddSso, e, telemetryReporter)
    );
  }

  return ok(Void);
}

export async function addSso(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json
): Promise<Result<unknown, FxError>> {
  ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddSsoStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  let solutionSettings = getAzureSolutionSettings(ctx);
  if (!solutionSettings) {
    // pure existing app
    solutionSettings = {
      name: TeamsAppSolutionNameV2,
      version: "1.0.0",
      hostType: "Azure",
      capabilities: [],
      azureResources: [],
      activeResourcePlugins: [],
    };
    ctx.projectSetting.solutionSettings = solutionSettings;
  }

  // Check whether can add sso
  const canProceed = canAddSso(solutionSettings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // Update project settings
  solutionSettings.activeResourcePlugins.push(PluginNames.AAD);
  solutionSettings.capabilities.push(SsoItem.id);

  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
    existingCapabilities: originalSettings.capabilities,
  };

  // TODO: Create folder with readme, auth page, etc.

  // Scaffold aad plugin and arm template
  const scaffoldRes = await scaffoldCodeAndResourceTemplate(
    ctx,
    inputsNew,
    localSettings,
    [Container.get<v2.ResourcePlugin>(PluginNames.AAD)],
    [Container.get<v2.ResourcePlugin>(PluginNames.AAD)]
  );
  if (scaffoldRes.isErr()) {
    ctx.projectSetting.solutionSettings = originalSettings;
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddSso,
        scaffoldRes.error,
        ctx.telemetryReporter
      )
    );
  }

  // Update manifest
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  await appStudioPlugin.addCapabilities(ctx, inputs as v2.InputsWithProjectPath, [
    { name: "WebApplicationInfo" },
  ]);

  return ok(undefined);
}
