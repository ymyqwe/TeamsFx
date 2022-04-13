import { SpfxPlugin } from "./../localdebug/constants";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  PluginContext,
  FxError,
  Result,
  ok,
  Platform,
  Colors,
  err,
  UserCancelError,
  v3,
  IStaticTab,
  IConfigurableTab,
} from "@microsoft/teamsfx-api";
import * as uuid from "uuid";
import lodash from "lodash";
import * as fs from "fs-extra";
import * as path from "path";
import { SPFXQuestionNames } from "./utils/questions";
import { Utils, sleep } from "./utils/utils";
import {
  Constants,
  DeployProgressMessage,
  ManifestTemplate,
  PlaceHolders,
  PreDeployProgressMessage,
  ScaffoldProgressMessage,
} from "./utils/constants";
import {
  BuildSPPackageError,
  CreateAppCatalogFailedError,
  GetGraphTokenFailedError,
  GetSPOTokenFailedError,
  NoSPPackageError,
  ScaffoldError,
  GetTenantFailedError,
  UploadAppPackageFailedError,
  InsufficientPermissionError,
  DependencyInstallError,
} from "./error";
import * as util from "util";
import { ProgressHelper } from "./utils/progress-helper";
import { getAppDirectory } from "../../../common/tools";
import { getTemplatesFolder } from "../../../folder";
import {
  MANIFEST_LOCAL,
  MANIFEST_TEMPLATE,
  MANIFEST_TEMPLATE_CONSOLIDATE,
} from "../appstudio/constants";
import axios from "axios";
import { SPOClient } from "./spoClient";
import { isConfigUnifyEnabled } from "../../../common";
import { DefaultManifestProvider } from "../../solution/fx-solution/v3/addFeature";
import { convert2Context } from "../utils4v2";
import { yeomanScaffoldEnabled } from "../../../core/globalVars";
import { getLocalizedString } from "../../../common/localizeUtils";
import { YoChecker } from "./depsChecker/yoChecker";
import { GeneratorChecker } from "./depsChecker/generatorChecker";
import { cpUtils } from "../../solution/fx-solution/utils/depsChecker/cpUtils";

export class SPFxPluginImpl {
  public async postScaffold(ctx: PluginContext): Promise<Result<any, FxError>> {
    try {
      const webpartName = ctx.answers![SPFXQuestionNames.webpart_name] as string;
      const componentName = Utils.normalizeComponentName(webpartName);
      const componentNameCamelCase = lodash.camelCase(componentName);
      let componentId: string;
      const replaceMap: Map<string, string> = new Map();
      if (yeomanScaffoldEnabled()) {
        const progressHandler = await ProgressHelper.startScaffoldProgressHandler(ctx.ui);
        await progressHandler?.next(ScaffoldProgressMessage.ScaffoldProject);

        const yoChecker = new YoChecker(ctx.logProvider!);
        const yoRes = await yoChecker.ensureDependency(ctx);
        if (yoRes.isErr()) {
          throw DependencyInstallError("yo");
        }
        const spGeneratorChecker = new GeneratorChecker(ctx.logProvider!);
        const spGeneratorRes = await spGeneratorChecker.ensureDependency(ctx);
        if (spGeneratorRes.isErr()) {
          throw DependencyInstallError("sharepoint generator");
        }

        const webpartDescription = ctx.answers![SPFXQuestionNames.webpart_desp] as string;
        const framework = ctx.answers![SPFXQuestionNames.framework_type] as string;
        const componentType = ctx.answers![SPFXQuestionNames.component_type] as string;
        const aceType = ctx.answers![SPFXQuestionNames.ace_type] as string;
        const solutionName = ctx.projectSettings?.appName as string;
        if (ctx.answers?.platform === Platform.VSCode) {
          (ctx.logProvider as any).outputChannel.show();
        }

        const yoEnv: NodeJS.ProcessEnv = process.env;
        yoEnv.PATH = `${yoChecker.getBinFolder()}${path.delimiter}${process.env.PATH ?? ""}`;
        await cpUtils.executeCommand(
          ctx.root,
          ctx.logProvider,
          {
            timeout: 2 * 60 * 1000,
            env: yoEnv,
          },
          "yo",
          spGeneratorChecker.getSpGeneratorPath(),
          "--skip-install",
          "true",
          "--component-type",
          componentType,
          "--component-description",
          webpartDescription,
          "--aceTemplateType",
          aceType,
          "--component-name",
          webpartName,
          "--solution-name",
          solutionName,
          "--skip-feature-deployment",
          "true",
          "--is-domain-isolated",
          "false"
        );

        const currentPath = path.join(ctx.root, solutionName);
        const newPath = path.join(ctx.root, "SPFx");
        await fs.rename(currentPath, newPath);

        await progressHandler?.next(ScaffoldProgressMessage.UpdateManifest);
        // const manifestPath = `${newPath}/src/webparts/${componentNameCamelCase}/${componentName}WebPart.manifest.json`;
        // const manifest = await fs.readFile(manifestPath, "utf8");
        // let manifestString = manifest.toString();
        // manifestString = manifestString.replace(
        //   `"supportedHosts": ["SharePointWebPart"]`,
        //   `"supportedHosts": ["SharePointWebPart", "TeamsPersonalApp", "TeamsTab"]`
        // );
        // await fs.writeFile(manifestPath, manifestString);

        // remove dataVersion() function, related issue: https://github.com/SharePoint/sp-dev-docs/issues/6469
        // const webpartFile = `${newPath}/src/webparts/${componentNameCamelCase}/${componentName}WebPart.ts`;
        // const codeFile = await fs.readFile(webpartFile, "utf8");
        // let codeString = codeFile.toString();
        // codeString = codeString.replace(
        //   `  protected get dataVersion(): Version {\r\n    return Version.parse('1.0');\r\n  }\r\n\r\n`,
        //   ``
        // );
        // codeString = codeString.replace(
        //   `import { Version } from '@microsoft/sp-core-library';\r\n`,
        //   ``
        // );
        // await fs.writeFile(webpartFile, codeString);

        // remove .vscode
        const debugPath = `${newPath}/.vscode`;
        await fs.remove(debugPath);

        const solutionPath = `${newPath}/config/package-solution.json`;
        const solution = await fs.readJson(solutionPath);
        componentId = solution.solution.id;

        replaceMap.set(PlaceHolders.componentId, componentId);
        replaceMap.set(PlaceHolders.componentNameUnescaped, webpartName);
      } else {
        componentId = uuid.v4();
        const componentClassName = `${componentName}WebPart`;
        const componentStrings = componentClassName + "Strings";
        const libraryName = lodash.kebabCase(ctx.projectSettings?.appName);
        let componentAlias = componentClassName;
        if (componentClassName.length > Constants.MAX_ALIAS_LENGTH) {
          componentAlias = componentClassName.substring(0, Constants.MAX_ALIAS_LENGTH);
        }
        let componentClassNameKebabCase = lodash.kebabCase(componentClassName);
        if (componentClassNameKebabCase.length > Constants.MAX_BUNDLE_NAME_LENGTH) {
          componentClassNameKebabCase = componentClassNameKebabCase.substring(
            0,
            Constants.MAX_BUNDLE_NAME_LENGTH
          );
          const lastCharacterIndex = componentClassNameKebabCase.length - 1;
          if (componentClassNameKebabCase[lastCharacterIndex] === "-") {
            componentClassNameKebabCase = componentClassNameKebabCase.substring(
              0,
              lastCharacterIndex
            );
          }
        }

        const outputFolderPath = `${ctx.root}/SPFx`;
        await fs.mkdir(outputFolderPath);

        // teams folder
        const teamsDir = `${outputFolderPath}/teams`;

        const templateFolder = path.join(getTemplatesFolder(), "plugins", "resource", "spfx");

        await fs.mkdir(teamsDir);
        await fs.copyFile(
          path.resolve(templateFolder, "./webpart/base/images/color.png"),
          `${teamsDir}/${componentId}_color.png`
        );
        await fs.copyFile(
          path.resolve(templateFolder, "./webpart/base/images/outline.png"),
          `${teamsDir}/${componentId}_outline.png`
        );

        // src folder
        const srcDir = `${outputFolderPath}/src`;
        await fs.mkdir(srcDir);
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/src/index.ts"),
          `${srcDir}/index.ts`
        );

        switch (ctx.answers![SPFXQuestionNames.framework_type] as string) {
          case Constants.FRAMEWORK_NONE:
            fs.mkdirSync(`${srcDir}/webparts/${componentNameCamelCase}`, {
              recursive: true,
            });
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/none/{componentClassName}.module.scss"),
              `${srcDir}/webparts/${componentNameCamelCase}/${componentClassName}.module.scss`
            );
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/none/{componentClassName}.ts"),
              `${srcDir}/webparts/${componentNameCamelCase}/${componentClassName}.ts`
            );
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/none/package.json"),
              `${outputFolderPath}/package.json`
            );
            break;
          case Constants.FRAMEWORK_REACT:
            const componentDir = `${srcDir}/webparts/${componentNameCamelCase}/components`;
            fs.mkdirSync(componentDir, { recursive: true });
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/react/{componentClassName}.ts"),
              `${srcDir}/webparts/${componentNameCamelCase}/${componentClassName}.ts`
            );
            await fs.copyFile(
              path.resolve(
                templateFolder,
                "./webpart/react/components/{componentName}.module.scss"
              ),
              `${componentDir}/${componentName}.module.scss`
            );
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/react/components/{componentName}.tsx"),
              `${componentDir}/${componentName}.tsx`
            );
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/react/components/I{componentName}Props.ts"),
              `${componentDir}/I${componentName}Props.ts`
            );
            await fs.copyFile(
              path.resolve(templateFolder, "./webpart/react/package.json"),
              `${outputFolderPath}/package.json`
            );
            break;
        }

        await fs.copy(
          path.resolve(templateFolder, "./webpart/base/loc"),
          `${srcDir}/webparts/${componentNameCamelCase}/loc`
        );
        await fs.copy(
          path.resolve(templateFolder, "./webpart/base/{componentClassName}.manifest.json"),
          `${srcDir}/webparts/${componentNameCamelCase}/${componentClassName}.manifest.json`
        );

        // config folder
        await fs.copy(
          path.resolve(templateFolder, "./solution/config"),
          `${outputFolderPath}/config`
        );

        // Other files
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/README.md"),
          `${outputFolderPath}/README.md`
        );
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/_gitignore"),
          `${outputFolderPath}/.gitignore`
        );
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/gulpfile.js"),
          `${outputFolderPath}/gulpfile.js`
        );
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/tsconfig.json"),
          `${outputFolderPath}/tsconfig.json`
        );
        await fs.copyFile(
          path.resolve(templateFolder, "./solution/tslint.json"),
          `${outputFolderPath}/tslint.json`
        );

        // Configure placeholders
        replaceMap.set(PlaceHolders.componentName, componentName);
        replaceMap.set(PlaceHolders.componentNameCamelCase, componentNameCamelCase);
        replaceMap.set(PlaceHolders.componentClassName, componentClassName);
        replaceMap.set(PlaceHolders.componentStrings, componentStrings);
        replaceMap.set(PlaceHolders.libraryName, libraryName);
        replaceMap.set(PlaceHolders.componentId, componentId);
        replaceMap.set(PlaceHolders.componentAlias, componentAlias);
        replaceMap.set(
          PlaceHolders.componentDescription,
          ctx.answers![SPFXQuestionNames.webpart_desp] as string
        );
        replaceMap.set(PlaceHolders.componentNameUnescaped, webpartName);
        replaceMap.set(PlaceHolders.componentClassNameKebabCase, componentClassNameKebabCase);

        await Utils.configure(outputFolderPath, replaceMap);
      }

      const appDirectory = await getAppDirectory(ctx.root);
      if (isConfigUnifyEnabled()) {
        await Utils.configure(path.join(appDirectory, MANIFEST_TEMPLATE_CONSOLIDATE), replaceMap);

        const appManifestProvider = new DefaultManifestProvider();
        const capabilitiesToAddManifest: v3.ManifestCapability[] = [];
        const remoteStaticSnippet: IStaticTab = {
          entityId: componentId,
          name: webpartName,
          contentUrl: util.format(ManifestTemplate.REMOTE_CONTENT_URL, componentId, componentId),
          websiteUrl: ManifestTemplate.WEBSITE_URL,
          scopes: ["personal"],
        };
        const remoteConfigurableSnippet: IConfigurableTab = {
          configurationUrl: util.format(
            ManifestTemplate.REMOTE_CONFIGURATION_URL,
            componentId,
            componentId
          ),
          canUpdateConfiguration: true,
          scopes: ["team"],
        };
        capabilitiesToAddManifest.push(
          {
            name: "staticTab",
            snippet: remoteStaticSnippet,
          },
          {
            name: "configurableTab",
            snippet: remoteConfigurableSnippet,
          }
        );

        const contextWithInputs = convert2Context(ctx, true);
        for (const capability of capabilitiesToAddManifest) {
          const addCapRes = await appManifestProvider.updateCapability(
            contextWithInputs.context,
            contextWithInputs.inputs,
            capability
          );
          if (addCapRes.isErr()) return err(addCapRes.error);
        }
      } else {
        await Utils.configure(path.join(appDirectory, MANIFEST_TEMPLATE), replaceMap);
        await Utils.configure(path.join(appDirectory, MANIFEST_LOCAL), replaceMap);
      }
      return ok(undefined);
    } catch (error) {
      return err(ScaffoldError(error));
    }
  }

  private async buildSPPackage(ctx: PluginContext): Promise<Result<any, FxError>> {
    const progressHandler = await ProgressHelper.startPreDeployProgressHandler(ctx.ui);
    if (ctx.answers?.platform === Platform.VSCode) {
      (ctx.logProvider as any).outputChannel.show();
    }
    try {
      const workspacePath = `${ctx.root}/SPFx`;
      await progressHandler?.next(PreDeployProgressMessage.NpmInstall);
      await Utils.execute(`npm install`, "SPFx", workspacePath, ctx.logProvider, true);
      const gulpCommand = await SPFxPluginImpl.findGulpCommand(workspacePath);
      await progressHandler?.next(PreDeployProgressMessage.GulpBundle);
      await Utils.execute(
        `${gulpCommand} bundle --ship --no-color`,
        "SPFx",
        workspacePath,
        ctx.logProvider,
        true
      );
      await progressHandler?.next(PreDeployProgressMessage.GulpPackage);
      await Utils.execute(
        `${gulpCommand} package-solution --ship --no-color`,
        "SPFx",
        workspacePath,
        ctx.logProvider,
        true
      );
      await ProgressHelper.endPreDeployProgress(true);

      const sharepointPackage = await this.getPackage(ctx.root);
      if (!(await fs.pathExists(sharepointPackage))) {
        throw NoSPPackageError(sharepointPackage);
      }

      const dir = path.normalize(path.parse(sharepointPackage).dir);

      if (ctx.answers?.platform === Platform.CLI) {
        const guidance = [
          {
            content: "Success: SharePoint package successfully built at ",
            color: Colors.BRIGHT_GREEN,
          },
          { content: dir, color: Colors.BRIGHT_MAGENTA },
        ];
        ctx.ui?.showMessage("info", guidance, false);
      } else {
        const guidance = getLocalizedString("plugins.spfx.buildNotice", dir);
        ctx.ui?.showMessage("info", guidance, false, "OK");
      }
      return ok(undefined);
    } catch (error) {
      await ProgressHelper.endPreDeployProgress(false);
      return err(BuildSPPackageError(error));
    }
  }

  public async preDeploy(ctx: PluginContext): Promise<Result<any, FxError>> {
    return this.buildSPPackage(ctx);
  }

  public async deploy(ctx: PluginContext): Promise<Result<any, FxError>> {
    const progressHandler = await ProgressHelper.startDeployProgressHandler(ctx.ui);
    let success = false;
    try {
      const tenant = await this.getTenant(ctx);
      if (tenant.isErr()) {
        return tenant;
      }
      SPOClient.setBaseUrl(tenant.value);

      const spoToken = await ctx.sharepointTokenProvider?.getAccessToken();
      if (!spoToken) {
        return err(GetSPOTokenFailedError());
      }

      let appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
      if (appCatalogSite) {
        SPOClient.setBaseUrl(appCatalogSite);
      } else {
        const res = await ctx.ui?.showMessage(
          "warn",
          getLocalizedString("plugins.spfx.createAppCatalogNotice", tenant.value),
          true,
          "OK",
          Constants.READ_MORE
        );
        const confirm = res?.isOk() ? res.value : undefined;
        switch (confirm) {
          case "OK":
            try {
              await progressHandler?.next(DeployProgressMessage.CreateSPAppCatalog);
              await SPOClient.createAppCatalog(spoToken);
            } catch (e: any) {
              return err(CreateAppCatalogFailedError(e));
            }
            let retry = 0;
            appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
            while (appCatalogSite == null && retry < Constants.APP_CATALOG_MAX_TIMES) {
              ctx.logProvider?.warning(`No tenant app catalog found, retry: ${retry}`);
              await sleep(Constants.APP_CATALOG_REFRESH_TIME);
              appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
              retry += 1;
            }
            if (appCatalogSite) {
              SPOClient.setBaseUrl(appCatalogSite);
              ctx.logProvider?.info(
                `Sharepoint tenant app catalog ${appCatalogSite} created, wait for a few minutes to be active.`
              );
              await sleep(Constants.APP_CATALOG_ACTIVE_TIME);
            } else {
              return err(
                CreateAppCatalogFailedError(
                  new Error(getLocalizedString("plugins.spfx,cannotGetAppcatalog"))
                )
              );
            }
            break;
          case Constants.READ_MORE:
            ctx.ui?.openUrl(Constants.CREATE_APP_CATALOG_GUIDE);
            return ok(UserCancelError);
          default:
            return ok(undefined);
        }
      }

      const appPackage = await this.getPackage(ctx.root);
      if (!(await fs.pathExists(appPackage))) {
        return err(NoSPPackageError(appPackage));
      }

      const fileName = path.parse(appPackage).base;
      const bytes = await fs.readFile(appPackage);
      try {
        await progressHandler?.next(DeployProgressMessage.UploadAndDeploy);
        await SPOClient.uploadAppPackage(spoToken, fileName, bytes);
      } catch (e: any) {
        if (e.response?.status === 403) {
          ctx.ui?.showMessage(
            "error",
            getLocalizedString("plugins.spfx.deployFailedNotice", appCatalogSite!),
            false,
            "OK"
          );
          return err(InsufficientPermissionError(appCatalogSite!));
        } else {
          return err(UploadAppPackageFailedError(e));
        }
      }

      const appID = await this.getAppID(ctx.root);
      await SPOClient.deployAppPackage(spoToken, appID);

      const guidance = getLocalizedString(
        "plugins.spfx.deployNotice",
        appPackage,
        appCatalogSite,
        appCatalogSite
      );
      if (ctx.answers?.platform === Platform.CLI) {
        ctx.ui?.showMessage("info", guidance, false);
      } else {
        ctx.ui?.showMessage("info", guidance, false, "OK");
      }
      success = true;
      return ok(undefined);
    } finally {
      await ProgressHelper.endDeployProgress(success);
    }
  }

  private async getTenant(ctx: PluginContext): Promise<Result<string, FxError>> {
    const graphToken = await ctx.graphTokenProvider?.getAccessToken();
    if (!graphToken) {
      return err(GetGraphTokenFailedError());
    }

    const tokenJson = await ctx.graphTokenProvider?.getJsonObject();
    const username = (tokenJson as any).unique_name;

    const instance = axios.create({
      baseURL: "https://graph.microsoft.com/v1.0",
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${graphToken}`;

    let tenant = "";
    try {
      const res = await instance.get("/sites/root?$select=webUrl");
      if (res && res.data && res.data.webUrl) {
        tenant = res.data.webUrl;
      } else {
        return err(GetTenantFailedError(username));
      }
    } catch (e) {
      return err(GetTenantFailedError(username, e));
    }
    return ok(tenant);
  }

  private async getPackage(root: string): Promise<string> {
    const solutionConfig = await fs.readJson(`${root}/SPFx/config/package-solution.json`);
    const sharepointPackage = `${root}/SPFx/sharepoint/${solutionConfig.paths.zippedPackage}`;
    return sharepointPackage;
  }

  private async getAppID(root: string): Promise<string> {
    const solutionConfig = await fs.readJson(`${root}/SPFx/config/package-solution.json`);
    const appID = solutionConfig["solution"]["id"];
    return appID;
  }

  private static async findGulpCommand(rootPath: string): Promise<string> {
    let gulpCommand: string;
    const platform = process.platform;
    if (
      platform === "win32" &&
      (await fs.pathExists(path.join(rootPath, "node_modules", ".bin", "gulp.cmd")))
    ) {
      gulpCommand = path.join(".", "node_modules", ".bin", "gulp.cmd");
    } else if (
      (platform === "linux" || platform === "darwin") &&
      (await fs.pathExists(path.join(rootPath, "node_modules", ".bin", "gulp")))
    ) {
      gulpCommand = path.join(".", "node_modules", ".bin", "gulp");
    } else {
      gulpCommand = "gulp";
    }
    return gulpCommand;
  }
}
