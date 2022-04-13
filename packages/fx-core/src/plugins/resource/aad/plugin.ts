// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, LogProvider, PluginContext, Result } from "@microsoft/teamsfx-api";
import { AadResult, ResultFactory } from "./results";
import {
  CheckGrantPermissionConfig,
  ConfigUtils,
  PostProvisionConfig,
  ProvisionConfig,
  SetApplicationInContextConfig,
  UpdatePermissionConfig,
  Utils,
} from "./utils/configs";
import { TelemetryUtils } from "./utils/telemetry";
import { TokenAudience, TokenProvider } from "./utils/tokenProvider";
import { AadAppClient } from "./aadAppClient";
import {
  AppIdUriInvalidError,
  ParsePermissionError,
  UnknownPermissionName,
  UnknownPermissionRole,
  UnknownPermissionScope,
  GetConfigError,
  ConfigErrorMessages,
} from "./errors";
import { Envs } from "./interfaces/models";
import { DialogUtils } from "./utils/dialog";
import {
  ConfigKeys,
  Constants,
  Messages,
  Plugins,
  ProgressDetail,
  ProgressTitle,
  Telemetry,
  TemplatePathInfo,
} from "./constants";
import { IPermission } from "./interfaces/IPermission";
import {
  IAADDefinition,
  RequiredResourceAccess,
  ResourceAccess,
} from "./interfaces/IAADDefinition";
import { validate as uuidValidate } from "uuid";
import * as path from "path";
import * as fs from "fs-extra";
import { ArmTemplateResult } from "../../../common/armInterface";
import { Bicep, ConstantString } from "../../../common/constants";
import { getTemplatesFolder } from "../../../folder";
import { AadOwner, ResourcePermission } from "../../../common/permissionInterface";
import { IUserList } from "../appstudio/interfaces/IAppDefinition";
import { isAadManifestEnabled, isConfigUnifyEnabled } from "../../../common/tools";
import { getPermissionMap } from "./permissions";

export class AadAppForTeamsImpl {
  public async provision(ctx: PluginContext, isLocalDebug = false): Promise<AadResult> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetryWithLocalDebug(
      ctx.logProvider,
      Messages.StartProvision,
      Messages.StartLocalDebug,
      isLocalDebug
    );

    const telemetryMessage = isLocalDebug
      ? Messages.EndLocalDebug.telemetry
      : Messages.EndProvision.telemetry;

    await TokenProvider.init({ graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken });

    // Move objectId etc. from input to output.
    const skip = Utils.skipAADProvision(
      ctx,
      isLocalDebug ? (isConfigUnifyEnabled() ? false : true) : false
    );
    DialogUtils.init(ctx.ui, ProgressTitle.Provision, ProgressTitle.ProvisionSteps);

    let config: ProvisionConfig = new ProvisionConfig(
      isLocalDebug ? (isConfigUnifyEnabled() ? false : true) : false
    );
    await config.restoreConfigFromContext(ctx);
    const permissions = AadAppForTeamsImpl.parsePermission(
      config.permissionRequest as string,
      ctx.logProvider
    );

    await DialogUtils.progress?.start(ProgressDetail.Starting);
    if (config.objectId) {
      if (!skip) {
        await DialogUtils.progress?.next(ProgressDetail.GetAadApp);
        config = await AadAppClient.getAadApp(
          telemetryMessage,
          config.objectId,
          config.password,
          ctx.graphTokenProvider,
          isLocalDebug
            ? isConfigUnifyEnabled()
              ? ctx.envInfo.envName
              : undefined
            : ctx.envInfo.envName
        );
        ctx.logProvider?.info(Messages.getLog(Messages.GetAadAppSuccess));
      }
    } else {
      await DialogUtils.progress?.next(ProgressDetail.ProvisionAadApp);
      await AadAppClient.createAadApp(telemetryMessage, config);
      config.password = undefined;
      ctx.logProvider?.info(Messages.getLog(Messages.CreateAadAppSuccess));
    }

    if (!config.password) {
      await DialogUtils.progress?.next(ProgressDetail.CreateAadAppSecret);
      await AadAppClient.createAadAppSecret(telemetryMessage, config);
      ctx.logProvider?.info(Messages.getLog(Messages.CreateAadAppPasswordSuccess));
    }

    await DialogUtils.progress?.next(ProgressDetail.UpdatePermission);
    await AadAppClient.updateAadAppPermission(
      telemetryMessage,
      config.objectId as string,
      permissions,
      skip
    );
    ctx.logProvider?.info(Messages.getLog(Messages.UpdatePermissionSuccess));

    await DialogUtils.progress?.end(true);
    config.saveConfigIntoContext(ctx, TokenProvider.tenantId as string);
    Utils.addLogAndTelemetryWithLocalDebug(
      ctx.logProvider,
      Messages.EndProvision,
      Messages.EndLocalDebug,
      isLocalDebug,
      skip ? { [Telemetry.skip]: Telemetry.yes } : {}
    );
    return ResultFactory.Success();
  }

  public setApplicationInContext(ctx: PluginContext, isLocalDebug = false): AadResult {
    const config: SetApplicationInContextConfig = new SetApplicationInContextConfig(isLocalDebug);
    config.restoreConfigFromContext(ctx);

    if (!config.frontendDomain && !config.botId) {
      throw ResultFactory.UserError(AppIdUriInvalidError.name, AppIdUriInvalidError.message());
    }

    let applicationIdUri = "api://";
    applicationIdUri += config.frontendDomain ? `${config.frontendDomain}/` : "";
    applicationIdUri += config.botId ? "botid-" + config.botId : config.clientId;
    config.applicationIdUri = applicationIdUri;

    ctx.logProvider?.info(Messages.getLog(Messages.SetAppIdUriSuccess));
    config.saveConfigIntoContext(ctx);
    return ResultFactory.Success();
  }

  public async postProvision(ctx: PluginContext, isLocalDebug = false): Promise<AadResult> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetryWithLocalDebug(
      ctx.logProvider,
      Messages.StartPostProvision,
      Messages.StartPostLocalDebug,
      isLocalDebug
    );

    const skip = Utils.skipAADProvision(
      ctx,
      isLocalDebug ? (isConfigUnifyEnabled() ? false : true) : false
    );
    DialogUtils.init(ctx.ui, ProgressTitle.PostProvision, ProgressTitle.PostProvisionSteps);

    await TokenProvider.init({ graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken });
    const config: PostProvisionConfig = new PostProvisionConfig(
      isLocalDebug ? (isConfigUnifyEnabled() ? false : true) : false
    );
    config.restoreConfigFromContext(ctx);

    await DialogUtils.progress?.start(ProgressDetail.Starting);
    await DialogUtils.progress?.next(ProgressDetail.UpdateRedirectUri);

    const redirectUris: IAADDefinition = AadAppForTeamsImpl.getRedirectUris(
      config.frontendEndpoint,
      config.botEndpoint,
      config.clientId!
    );
    await AadAppClient.updateAadAppRedirectUri(
      isLocalDebug ? Messages.EndPostLocalDebug.telemetry : Messages.EndPostProvision.telemetry,
      config.objectId as string,
      redirectUris,
      skip
    );
    ctx.logProvider?.info(Messages.getLog(Messages.UpdateRedirectUriSuccess));

    await DialogUtils.progress?.next(ProgressDetail.UpdateAppIdUri);
    await AadAppClient.updateAadAppIdUri(
      isLocalDebug ? Messages.EndPostLocalDebug.telemetry : Messages.EndPostProvision.telemetry,
      config.objectId as string,
      config.applicationIdUri as string,
      skip
    );
    ctx.logProvider?.info(Messages.getLog(Messages.UpdateAppIdUriSuccess));

    await DialogUtils.progress?.end(true);
    Utils.addLogAndTelemetryWithLocalDebug(
      ctx.logProvider,
      Messages.EndPostProvision,
      Messages.EndPostLocalDebug,
      isLocalDebug,
      skip ? { [Telemetry.skip]: Telemetry.yes } : {}
    );
    return ResultFactory.Success();
  }

  public async updatePermission(ctx: PluginContext): Promise<AadResult> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.StartUpdatePermission);
    const skip: boolean = ctx.config.get(ConfigKeys.skip) as boolean;
    if (skip) {
      ctx.logProvider?.info(Messages.SkipProvision);
      Utils.addLogAndTelemetry(ctx.logProvider, Messages.EndUpdatePermission);
      return ResultFactory.Success();
    }

    DialogUtils.init(ctx.ui, ProgressTitle.UpdatePermission, ProgressTitle.UpdatePermissionSteps);

    const configs = await AadAppForTeamsImpl.getUpdatePermissionConfigs(ctx);
    if (!configs) {
      return ResultFactory.Success();
    }

    await TokenProvider.init({ graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken });

    const permissions = AadAppForTeamsImpl.parsePermission(
      configs[0].permissionRequest as string,
      ctx.logProvider
    );

    await DialogUtils.progress?.start(ProgressDetail.Starting);
    await DialogUtils.progress?.next(ProgressDetail.UpdatePermission);
    for (const config of configs) {
      await AadAppClient.updateAadAppPermission(
        Messages.EndUpdatePermission.telemetry,
        config.objectId as string,
        permissions
      );
    }
    ctx.logProvider?.info(Messages.getLog(Messages.UpdatePermissionSuccess));

    await DialogUtils.progress?.end(true);
    DialogUtils.show(Messages.UpdatePermissionSuccessMessage);
    return ResultFactory.Success();
  }

  public async generateArmTemplates(ctx: PluginContext): Promise<AadResult> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.StartGenerateArmTemplates);

    const result: ArmTemplateResult = {
      Parameters: JSON.parse(
        await fs.readFile(
          path.join(
            getTemplatesFolder(),
            TemplatePathInfo.BicepTemplateRelativeDir,
            Bicep.ParameterFileName
          ),
          ConstantString.UTF8Encoding
        )
      ),
    };

    Utils.addLogAndTelemetry(ctx.logProvider, Messages.EndGenerateArmTemplates);
    return ResultFactory.Success(result);
  }

  public async checkPermission(
    ctx: PluginContext,
    userInfo: IUserList
  ): Promise<Result<ResourcePermission[], FxError>> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.StartCheckPermission);

    await TokenProvider.init(
      { graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken },
      TokenAudience.Graph
    );
    const config = new CheckGrantPermissionConfig();
    await config.restoreConfigFromContext(ctx);

    const userObjectId = userInfo.aadId;
    const isAadOwner = await AadAppClient.checkPermission(
      Messages.EndCheckPermission.telemetry,
      config.objectId!,
      userObjectId
    );

    const result = [
      {
        name: Constants.permissions.name,
        type: Constants.permissions.type,
        roles: isAadOwner ? [Constants.permissions.owner] : [Constants.permissions.noPermission],
        resourceId: config.objectId!,
      },
    ];
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.EndCheckPermission);
    return ResultFactory.Success(result);
  }

  public async listCollaborator(ctx: PluginContext): Promise<Result<AadOwner[], FxError>> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.StartListCollaborator);

    await TokenProvider.init(
      { graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken },
      TokenAudience.Graph
    );

    const objectId = ConfigUtils.getAadConfig(ctx, ConfigKeys.objectId, false);
    if (!objectId) {
      throw ResultFactory.SystemError(
        GetConfigError.name,
        GetConfigError.message(
          ConfigErrorMessages.GetConfigError(ConfigKeys.objectId, Plugins.pluginName)
        )
      );
    }

    const owners = await AadAppClient.listCollaborator(
      Messages.EndListCollaborator.telemetry,
      objectId
    );
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.EndListCollaborator);
    return ResultFactory.Success(owners);
  }

  public async grantPermission(
    ctx: PluginContext,
    userInfo: IUserList
  ): Promise<Result<ResourcePermission[], FxError>> {
    TelemetryUtils.init(ctx);
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.StartGrantPermission);

    await TokenProvider.init(
      { graph: ctx.graphTokenProvider, appStudio: ctx.appStudioToken },
      TokenAudience.Graph
    );
    const config = new CheckGrantPermissionConfig(true);
    await config.restoreConfigFromContext(ctx);

    const userObjectId = userInfo.aadId;
    await AadAppClient.grantPermission(ctx, config.objectId!, userObjectId);

    const result = [
      {
        name: Constants.permissions.name,
        type: Constants.permissions.type,
        roles: [Constants.permissions.owner],
        resourceId: config.objectId!,
      },
    ];
    Utils.addLogAndTelemetry(ctx.logProvider, Messages.EndGrantPermission);
    return ResultFactory.Success(result);
  }

  public static getRedirectUris(
    frontendEndpoint: string | undefined,
    botEndpoint: string | undefined,
    clientId: string
  ) {
    const redirectUris: IAADDefinition = {
      web: {
        redirectUris: [],
      },
      spa: {
        redirectUris: [],
      },
    };
    if (frontendEndpoint) {
      redirectUris.web?.redirectUris?.push(`${frontendEndpoint}/auth-end.html`);
      redirectUris.spa?.redirectUris?.push(`${frontendEndpoint}/blank-auth-end.html`);
      redirectUris.spa?.redirectUris?.push(
        `${frontendEndpoint}/auth-end.html?clientId=${clientId}`
      );
    }

    if (botEndpoint) {
      redirectUris.web?.redirectUris?.push(`${botEndpoint}/auth-end.html`);
    }

    return redirectUris;
  }

  private static async getUpdatePermissionConfigs(
    ctx: PluginContext
  ): Promise<UpdatePermissionConfig[] | undefined> {
    let azureAad = false;
    let localAad = false;
    if (ctx.config.get(ConfigKeys.objectId)) {
      azureAad = true;
    }
    if (ctx.config.get(Utils.addLocalDebugPrefix(true, ConfigKeys.objectId))) {
      localAad = true;
    }

    if (azureAad && localAad) {
      const ans = ctx.answers![Constants.AskForEnvName];
      if (!ans) {
        ctx.logProvider?.info(Messages.UserCancelled);
        return undefined;
      }
      if (ans === Envs.Azure) {
        localAad = false;
      } else if (ans === Envs.LocalDebug) {
        azureAad = false;
      }
    }

    if (!azureAad && !localAad) {
      await DialogUtils.show(Messages.NoSelection, "info");
      return undefined;
    }

    const configs: UpdatePermissionConfig[] = [];
    if (azureAad) {
      const config: UpdatePermissionConfig = new UpdatePermissionConfig();
      await config.restoreConfigFromContext(ctx);
      configs.push(config);
    }

    if (localAad) {
      const config: UpdatePermissionConfig = new UpdatePermissionConfig(true);
      await config.restoreConfigFromContext(ctx);
      configs.push(config);
    }

    return configs;
  }

  public static parsePermission(
    permissionRequest: string,
    logProvider?: LogProvider
  ): RequiredResourceAccess[] {
    let permissionRequestParsed: IPermission[];
    try {
      permissionRequestParsed = <IPermission[]>JSON.parse(permissionRequest as string);
    } catch (error) {
      throw ResultFactory.UserError(
        ParsePermissionError.name,
        ParsePermissionError.message(),
        error,
        undefined,
        ParsePermissionError.helpLink
      );
    }

    const permissions = AadAppForTeamsImpl.generateRequiredResourceAccess(permissionRequestParsed);
    if (!permissions) {
      throw ResultFactory.UserError(
        ParsePermissionError.name,
        ParsePermissionError.message(),
        undefined,
        undefined,
        ParsePermissionError.helpLink
      );
    }

    logProvider?.info(Messages.getLog(Messages.ParsePermissionSuccess));
    return permissions;
  }

  private static generateRequiredResourceAccess(
    permissions?: IPermission[]
  ): RequiredResourceAccess[] | undefined {
    if (!permissions) {
      return undefined;
    }

    const map = getPermissionMap();

    const requiredResourceAccessList: RequiredResourceAccess[] = [];

    permissions.forEach((permission) => {
      const requiredResourceAccess: RequiredResourceAccess = {};
      const resourceIdOrName = permission.resource;
      let resourceId = resourceIdOrName;
      if (!uuidValidate(resourceIdOrName)) {
        const res = map[resourceIdOrName];
        if (!res) {
          throw ResultFactory.UserError(
            UnknownPermissionName.name,
            UnknownPermissionName.message(resourceIdOrName),
            undefined,
            undefined,
            UnknownPermissionName.helpLink
          );
        }

        const id = res.id;
        if (!id) {
          throw ResultFactory.UserError(
            UnknownPermissionName.name,
            UnknownPermissionName.message(resourceIdOrName),
            undefined,
            undefined,
            UnknownPermissionName.helpLink
          );
        }
        resourceId = id;
      }

      requiredResourceAccess.resourceAppId = resourceId;
      requiredResourceAccess.resourceAccess = [];

      if (!permission.delegated) {
        permission.delegated = [];
      }

      if (!permission.application) {
        permission.application = [];
      }

      permission.delegated = permission.delegated?.concat(permission.scopes);
      permission.delegated = permission.delegated?.filter(
        (scopeName, i) => i === permission.delegated?.indexOf(scopeName)
      );

      permission.application = permission.application?.concat(permission.roles);
      permission.application = permission.application?.filter(
        (roleName, i) => i === permission.application?.indexOf(roleName)
      );

      permission.application?.forEach((roleName) => {
        if (!roleName) {
          return;
        }

        const resourceAccess: ResourceAccess = {
          id: roleName,
          type: "Role",
        };

        if (!uuidValidate(roleName)) {
          const roleId = map[resourceId].roles[roleName];
          if (!roleId) {
            throw ResultFactory.UserError(
              UnknownPermissionRole.name,
              UnknownPermissionRole.message(roleName, permission.resource),
              undefined,
              undefined,
              UnknownPermissionRole.helpLink
            );
          }
          resourceAccess.id = roleId;
        }

        requiredResourceAccess.resourceAccess!.push(resourceAccess);
      });

      permission.delegated?.forEach((scopeName) => {
        if (!scopeName) {
          return;
        }

        const resourceAccess: ResourceAccess = {
          id: scopeName,
          type: "Scope",
        };

        if (!uuidValidate(scopeName)) {
          const scopeId = map[resourceId].scopes[scopeName];
          if (!scopeId) {
            throw ResultFactory.UserError(
              UnknownPermissionScope.name,
              UnknownPermissionScope.message(scopeName, permission.resource),
              undefined,
              undefined,
              UnknownPermissionScope.helpLink
            );
          }
          resourceAccess.id = map[resourceId].scopes[scopeName];
        }

        requiredResourceAccess.resourceAccess!.push(resourceAccess);
      });

      requiredResourceAccessList.push(requiredResourceAccess);
    });

    return requiredResourceAccessList;
  }

  public async scaffold(ctx: PluginContext): Promise<AadResult> {
    if (isAadManifestEnabled()) {
      const templatesFolder = getTemplatesFolder();
      const appDir = `${ctx.root}/${Constants.appPackageFolder}`;
      const aadManifestTemplate = `${templatesFolder}/${Constants.aadManifestTemplateFolder}/${Constants.aadManifestTemplateName}`;
      await fs.ensureDir(appDir);
      await fs.copy(aadManifestTemplate, `${appDir}/${Constants.aadManifestTemplateName}`);
    }
    return ResultFactory.Success();
  }
}
