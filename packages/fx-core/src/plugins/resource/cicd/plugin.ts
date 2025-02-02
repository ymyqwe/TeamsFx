// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Context } from "@microsoft/teamsfx-api/build/v2";
import { Inputs, v2, Platform } from "@microsoft/teamsfx-api";
import { FxCICDPluginResultFactory as ResultFactory, FxResult } from "./result";
import { CICDProviderFactory } from "./providers/factory";
import { ProviderKind } from "./providers/enums";
import { questionNames } from "./questions";
import { InternalError, NoProjectOpenedError } from "./errors";
import { Logger } from "./logger";
import { getLocalizedString } from "../../../common/localizeUtils";

export class CICDImpl {
  public commonProperties: { [key: string]: string } = {};
  public async addCICDWorkflows(
    context: Context,
    inputs: Inputs,
    envInfo: v2.EnvInfoV2
  ): Promise<FxResult> {
    // 1. Key inputs (envName, provider, template) x (hostingType, ).
    if (!inputs.projectPath) {
      throw new NoProjectOpenedError();
    }
    const projectPath = inputs.projectPath;
    // By default(VSC), get env name from plugin's own `target-env` question.
    let envName = inputs[questionNames.Environment];
    // TODO: add support for VS/.Net Projects.
    if (inputs.platform === Platform.CLI) {
      // In CLI, get env name from the default `env` question.
      envName = envInfo.envName;
    }
    const providerName = inputs[questionNames.Provider];
    const templateNames = inputs[questionNames.Template] as string[];
    if (!envName || !providerName || templateNames.length === 0) {
      throw new InternalError(getLocalizedString("error.cicd.PreconditionNotMet"));
    }

    this.commonProperties = {
      env: envName,
      provider: providerName,
      template: templateNames.join(","),
    };

    // 2. Call factory to get provider instance.
    const providerInstance = CICDProviderFactory.create(providerName as ProviderKind);

    // 3. Call instance.scaffold(template, replacements: any).
    //  3.1 Call the initial scaffold.
    const progressBar = context.userInteraction.createProgressBar(
      getLocalizedString("plugins.cicd.ProgressBar.scaffold.title"),
      templateNames.length
    );

    const created: string[] = [];
    const skipped: string[] = [];

    await progressBar.start(
      getLocalizedString("plugins.cicd.ProgressBar.scaffold.detail", templateNames[0])
    );
    let scaffolded = await providerInstance.scaffold(
      projectPath,
      templateNames[0],
      envName,
      context
    );
    if (scaffolded.isOk() && !scaffolded.value) {
      created.push(templateNames[0]);
    } else {
      skipped.push(templateNames[0]);
    }

    //  3.2 Call the next scaffold.
    for (const templateName of templateNames.slice(1)) {
      await progressBar.next(
        getLocalizedString("plugins.cicd.ProgressBar.scaffold.detail", templateName)
      );
      scaffolded = await providerInstance.scaffold(projectPath, templateName, envName, context);
      if (scaffolded.isOk() && !scaffolded.value) {
        created.push(templateName);
      } else {
        skipped.push(templateName);
      }
    }

    await progressBar.end(true);

    // 4. Send notification messages.
    let message = "";
    if (created.length > 0) {
      message += getLocalizedString(
        "plugins.cicd.result.scaffold.created",
        created.join(","),
        providerName,
        envName
      );
    }
    if (skipped.length > 0) {
      message += getLocalizedString(
        "plugins.cicd.result.scaffold.skipped",
        skipped.join(","),
        providerName,
        envName
      );
    }

    context.userInteraction.showMessage("info", message, false);
    Logger.info(message);

    return ResultFactory.Success();
  }
}
