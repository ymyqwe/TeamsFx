// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AzureSolutionSettings,
  err,
  FxError,
  ok,
  Plugin,
  PluginContext,
  QTreeNode,
  Result,
  Stage,
  UserCancelError,
} from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { HostTypeOptionSPFx } from "../../solution/fx-solution/question";
import { ResourcePlugins } from "../../solution/fx-solution/ResourcePluginContainer";
import { SPFxPluginImpl } from "./plugin";
import { TelemetryEvent } from "./utils/constants";
import { ProgressHelper } from "./utils/progress-helper";
import {
  frameworkQuestion,
  SPFXQuestionNames,
  webpartDescriptionQuestion,
  webpartNameQuestion,
  aceTypeQuestion,
  componentTypeQuestion,
} from "./utils/questions";
import { yeomanScaffoldEnabled } from "../../../core/globalVars";
import { telemetryHelper } from "./utils/telemetry-helper";
import "./v2";
import "./v3";
@Service(ResourcePlugins.SpfxPlugin)
export class SpfxPlugin implements Plugin {
  name = "fx-resource-spfx";
  displayName = "SharePoint Framework (SPFx)";
  activate(solutionSettings: AzureSolutionSettings): boolean {
    return solutionSettings.hostType === HostTypeOptionSPFx.id;
  }
  spfxPluginImpl: SPFxPluginImpl = new SPFxPluginImpl();

  async getQuestions(
    stage: Stage,
    ctx: PluginContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const spfx_frontend_host = new QTreeNode({
      type: "group",
    });

    if (stage === Stage.create) {
      if (yeomanScaffoldEnabled()) {
        const componet_type = new QTreeNode(componentTypeQuestion);
        spfx_frontend_host.addChild(componet_type);

        const ace_type = new QTreeNode(aceTypeQuestion);
        spfx_frontend_host.addChild(ace_type);
      }

      // const spfx_framework_type = new QTreeNode(frameworkQuestion);
      // spfx_frontend_host.addChild(spfx_framework_type);

      const spfx_webpart_name = new QTreeNode(webpartNameQuestion);
      spfx_frontend_host.addChild(spfx_webpart_name);

      const spfx_webpart_desp = new QTreeNode(webpartDescriptionQuestion);
      spfx_frontend_host.addChild(spfx_webpart_desp);
    }

    return ok(spfx_frontend_host);
  }

  public async postScaffold(ctx: PluginContext): Promise<Result<any, FxError>> {
    return await this.runWithErrorHandling(ctx, TelemetryEvent.Scaffold, () =>
      this.spfxPluginImpl.postScaffold(ctx)
    );
  }

  public async preDeploy(ctx: PluginContext): Promise<Result<any, FxError>> {
    return await this.runWithErrorHandling(ctx, TelemetryEvent.PreDeploy, () =>
      this.spfxPluginImpl.preDeploy(ctx)
    );
  }

  public async deploy(ctx: PluginContext): Promise<Result<any, FxError>> {
    const result = await this.spfxPluginImpl.deploy(ctx);
    if (result.isOk()) {
      telemetryHelper.sendSuccessEvent(ctx, "deploy");
    } else {
      telemetryHelper.sendErrorEvent(ctx, "deploy", result.error);
      if (result.error.name === "InsufficientPermission") {
        return err(UserCancelError);
      }
    }
    return result;
  }

  private async runWithErrorHandling(
    ctx: PluginContext,
    stage: string,
    fn: () => Promise<Result<any, FxError>>
  ): Promise<Result<any, FxError>> {
    try {
      telemetryHelper.sendSuccessEvent(ctx, stage + TelemetryEvent.StartSuffix);
      const result = await fn();
      if (result.isOk()) {
        telemetryHelper.sendSuccessEvent(ctx, stage);
      } else {
        telemetryHelper.sendErrorEvent(ctx, stage, result.error);
      }
      return result;
    } catch (error) {
      await ProgressHelper.endAllHandlers(false);
      telemetryHelper.sendErrorEvent(ctx, stage, error);
      return err(error);
    }
  }
}

export default new SpfxPlugin();
