// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Plugin,
  PluginContext,
  SystemError,
  UserError,
  err,
  Func,
  ok,
  Result,
  FxError,
} from "@microsoft/teamsfx-api";
import { AadAppForTeamsImpl } from "./plugin";
import { AadResult, ResultFactory } from "./results";
import { UnhandledError } from "./errors";
import { TelemetryUtils } from "./utils/telemetry";
import { DialogUtils } from "./utils/dialog";
import { Messages, Plugins, Telemetry } from "./constants";
import { AzureSolutionSettings } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { ResourcePlugins } from "../../solution/fx-solution/ResourcePluginContainer";
import { Links } from "../bot/constants";
import { AadOwner, ResourcePermission } from "../../../common/permissionInterface";
import "./v2";
import "./v3";
import { IUserList } from "../appstudio/interfaces/IAppDefinition";
import { isAADEnabled } from "../../../common";
@Service(ResourcePlugins.AadPlugin)
export class AadAppForTeamsPlugin implements Plugin {
  name = "fx-resource-aad-app-for-teams";
  displayName = "AAD";
  activate(solutionSettings: AzureSolutionSettings): boolean {
    return isAADEnabled(solutionSettings);
  }

  public pluginImpl: AadAppForTeamsImpl = new AadAppForTeamsImpl();

  public async provision(ctx: PluginContext): Promise<AadResult> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.provision(ctx),
      ctx,
      Messages.EndProvision.telemetry
    );
  }

  public async localDebug(ctx: PluginContext): Promise<AadResult> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.provision(ctx, true),
      ctx,
      Messages.EndLocalDebug.telemetry
    );
  }

  public setApplicationInContext(ctx: PluginContext, isLocalDebug = false): AadResult {
    return this.runWithExceptionCatching(
      () => this.pluginImpl.setApplicationInContext(ctx, isLocalDebug),
      ctx
    );
  }

  public async postProvision(ctx: PluginContext): Promise<AadResult> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.postProvision(ctx),
      ctx,
      Messages.EndPostProvision.telemetry
    );
  }

  public async postLocalDebug(ctx: PluginContext): Promise<AadResult> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.postProvision(ctx, true),
      ctx,
      Messages.EndPostLocalDebug.telemetry
    );
  }

  public async generateArmTemplates(ctx: PluginContext): Promise<AadResult> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.generateArmTemplates(ctx),
      ctx,
      Messages.EndGenerateArmTemplates.telemetry
    );
  }

  public async executeUserTask(func: Func, ctx: PluginContext): Promise<AadResult> {
    if (func.method == "setApplicationInContext") {
      const isLocal: boolean =
        func.params && func.params.isLocal !== undefined ? (func.params.isLocal as boolean) : false;
      return Promise.resolve(this.setApplicationInContext(ctx, isLocal));
    }
    return err(
      new SystemError(
        "FunctionRouterError",
        `Failed to route function call:${JSON.stringify(func)}`,
        Plugins.pluginNameShort,
        undefined,
        Links.ISSUE_LINK
      )
    );
  }

  public async checkPermission(
    ctx: PluginContext,
    userInfo: Record<string, any>
  ): Promise<Result<ResourcePermission[], FxError>> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.checkPermission(ctx, userInfo as IUserList),
      ctx,
      Messages.EndCheckPermission.telemetry
    );
  }

  public async grantPermission(
    ctx: PluginContext,
    userInfo: Record<string, any>
  ): Promise<Result<ResourcePermission[], FxError>> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.grantPermission(ctx, userInfo as IUserList),
      ctx,
      Messages.EndCheckPermission.telemetry
    );
  }

  public async listCollaborator(ctx: PluginContext): Promise<Result<AadOwner[], FxError>> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.listCollaborator(ctx),
      ctx,
      Messages.EndListCollaborator.telemetry
    );
  }

  public async scaffold(ctx: PluginContext): Promise<Result<any, FxError>> {
    return await this.runWithExceptionCatchingAsync(
      () => this.pluginImpl.scaffold(ctx),
      ctx,
      Messages.Scaffold.telemetry
    );
  }

  private async runWithExceptionCatchingAsync(
    fn: () => Promise<AadResult>,
    ctx: PluginContext,
    stage: string
  ): Promise<AadResult> {
    try {
      return await fn();
    } catch (e) {
      return this.returnError(e, ctx, stage);
    }
  }

  private runWithExceptionCatching(fn: () => AadResult, ctx: PluginContext): AadResult {
    try {
      return fn();
    } catch (e) {
      return this.returnError(e, ctx, "");
    }
  }

  private returnError(e: any, ctx: PluginContext, stage: string): AadResult {
    if (e instanceof SystemError || e instanceof UserError) {
      let errorMessage = e.message;
      // For errors contains innerError, e.g. failures when calling Graph API
      if (e.innerError) {
        errorMessage += ` Detailed error: ${e.innerError.message}.`;
        if (e.innerError.response?.data?.errorMessage) {
          // For errors return from App Studio API
          errorMessage += ` Reason: ${e.innerError.response?.data?.errorMessage}`;
        } else if (e.innerError.response?.data?.error?.message) {
          // For errors return from Graph API
          errorMessage += ` Reason: ${e.innerError.response?.data?.error?.message}`;
        }
        e.message = errorMessage;
      }
      ctx.logProvider?.error(errorMessage);
      TelemetryUtils.init(ctx);
      TelemetryUtils.sendErrorEvent(
        stage,
        e.name,
        e instanceof UserError ? Telemetry.userError : Telemetry.systemError,
        errorMessage
      );
      DialogUtils.progress?.end(false);
      return err(e);
    } else {
      if (!(e instanceof Error)) {
        e = new Error(e.toString());
      }

      ctx.logProvider?.error(e.message);
      TelemetryUtils.init(ctx);
      TelemetryUtils.sendErrorEvent(
        stage,
        UnhandledError.name,
        Telemetry.systemError,
        UnhandledError.message() + " " + e.message
      );
      return err(
        ResultFactory.SystemError(
          UnhandledError.name,
          UnhandledError.message(),
          e,
          undefined,
          undefined
        )
      );
    }
  }
}

export default new AadAppForTeamsPlugin();
export { getPermissionMap } from "./permissions";
