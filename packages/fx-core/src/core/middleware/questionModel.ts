// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Middleware, NextFunction } from "@feathersjs/hooks";
import {
  CLIPlatforms,
  err,
  Func,
  FunctionRouter,
  FxError,
  Inputs,
  ok,
  Platform,
  QTreeNode,
  Result,
  Solution,
  SolutionContext,
  Stage,
  SystemError,
  traverse,
  v2,
  v3,
} from "@microsoft/teamsfx-api";
import { Container } from "typedi";
import { createV2Context, deepCopy } from "../../common/tools";
import { newProjectSettings } from "../../common/projectSettingsHelper";
import { SPFxPluginV3 } from "../../plugins/resource/spfx/v3";
import { TabSPFxItem } from "../../plugins/solution/fx-solution/question";
import {
  BuiltInFeaturePluginNames,
  BuiltInSolutionNames,
} from "../../plugins/solution/fx-solution/v3/constants";
import { getQuestionsForGrantPermission } from "../collaborator";
import { CoreSource, FunctionRouterError } from "../error";
import { TOOLS } from "../globalVars";
import {
  createAppNameQuestion,
  createCapabilityQuestion,
  getCreateNewOrFromSampleQuestion,
  M365AppTypeSelectQuestion,
  M365CapabilitiesFuncQuestion,
  M365CreateFromScratchSelectQuestion,
  ProgrammingLanguageQuestion,
  QuestionRootFolder,
  SampleSelect,
  ScratchOptionNo,
  ScratchOptionYes,
  ScratchOptionYesM365,
} from "../question";
import { getAllSolutionPluginsV2 } from "../SolutionPluginContainer";
import { CoreHookContext } from "../types";
/**
 * This middleware will help to collect input from question flow
 */
export const QuestionModelMW: Middleware = async (ctx: CoreHookContext, next: NextFunction) => {
  const inputs: Inputs = ctx.arguments[ctx.arguments.length - 1];
  const method = ctx.method;
  const core = ctx.self as any;

  let getQuestionRes: Result<QTreeNode | undefined, FxError> = ok(undefined);
  if (method === "createProjectV2") {
    getQuestionRes = await core._getQuestionsForCreateProjectV2(inputs);
  } else if (method === "createProjectV3") {
    getQuestionRes = await core._getQuestionsForCreateProjectV3(inputs);
  } else if (method === "init" || method === "_init") {
    getQuestionRes = await core._getQuestionsForInit(inputs);
  } else if (
    [
      "addFeature",
      "_addFeature",
      "provisionResourcesV3",
      "deployArtifactsV3",
      "publishApplicationV3",
      "executeUserTaskV3",
    ].includes(method || "")
  ) {
    const solutionV3 = ctx.solutionV3;
    const contextV2 = ctx.contextV2;
    if (solutionV3 && contextV2) {
      if (method === "addFeature" || method === "_addFeature") {
        getQuestionRes = await core._getQuestionsForAddFeature(
          inputs as v2.InputsWithProjectPath,
          solutionV3,
          contextV2
        );
      } else if (method === "provisionResourcesV3") {
        getQuestionRes = await core._getQuestionsForProvision(
          inputs as v2.InputsWithProjectPath,
          solutionV3,
          contextV2,
          ctx.envInfoV3 as v2.DeepReadonly<v3.EnvInfoV3>
        );
      } else if (method === "deployArtifactsV3") {
        getQuestionRes = await core._getQuestionsForDeploy(
          inputs as v2.InputsWithProjectPath,
          solutionV3,
          contextV2,
          ctx.envInfoV3 as v2.DeepReadonly<v3.EnvInfoV3>
        );
      } else if (method === "publishApplicationV3") {
        getQuestionRes = await core._getQuestionsForPublish(
          inputs as v2.InputsWithProjectPath,
          solutionV3,
          contextV2,
          ctx.envInfoV3 as v2.DeepReadonly<v3.EnvInfoV3>
        );
      } else if (method === "executeUserTaskV3") {
        const func = ctx.arguments[0] as Func;
        getQuestionRes = await core._getQuestionsForUserTaskV3(
          func,
          inputs,
          solutionV3,
          contextV2,
          ctx.envInfoV3 as v2.DeepReadonly<v3.EnvInfoV3>
        );
      }
    }
  } else if (method === "grantPermissionV3") {
    getQuestionRes = await getQuestionsForGrantPermission(inputs);
  } else {
    if (ctx.solutionV2 && ctx.contextV2) {
      const solution = ctx.solutionV2;
      const context = ctx.contextV2;
      if (solution && context) {
        if (method === "provisionResources" || method === "provisionResourcesV2") {
          getQuestionRes = await core._getQuestions(
            context,
            solution,
            Stage.provision,
            inputs,
            ctx.envInfoV2
          );
        } else if (method === "localDebug" || method === "localDebugV2") {
          getQuestionRes = await core._getQuestions(
            context,
            solution,
            Stage.debug,
            inputs,
            ctx.envInfoV2
          );
        } else if (method === "deployArtifacts" || method === "deployArtifactsV2") {
          getQuestionRes = await core._getQuestions(
            context,
            solution,
            Stage.deploy,
            inputs,
            ctx.envInfoV2
          );
        } else if (method === "publishApplication" || method === "publishApplicationV2") {
          getQuestionRes = await core._getQuestions(
            context,
            solution,
            Stage.publish,
            inputs,
            ctx.envInfoV2
          );
        } else if (method === "executeUserTaskV2") {
          const func = ctx.arguments[0] as Func;
          getQuestionRes = await core._getQuestionsForUserTask(
            context,
            solution,
            func,
            inputs,
            ctx.envInfoV2
          );
        } else if (method === "grantPermissionV2") {
          getQuestionRes = await core._getQuestions(
            context,
            solution,
            Stage.grantPermission,
            inputs,
            ctx.envInfoV2
          );
        }
      }
    }
  }

  if (getQuestionRes.isErr()) {
    TOOLS?.logProvider.error(
      `[core] failed to get questions for ${method}: ${getQuestionRes.error.message}`
    );
    ctx.result = err(getQuestionRes.error);
    return;
  }

  TOOLS?.logProvider.debug(`[core] success to get questions for ${method}`);

  const node = getQuestionRes.value;
  if (node) {
    const res = await traverse(node, inputs, TOOLS.ui, TOOLS.telemetryReporter);
    if (res.isErr()) {
      TOOLS?.logProvider.debug(`[core] failed to run question model for ${method}`);
      ctx.result = err(res.error);
      return;
    }
    const desensitized = desensitize(node, inputs);
    TOOLS?.logProvider.info(
      `[core] success to run question model for ${method}, answers:${JSON.stringify(desensitized)}`
    );
  }
  await next();
};

export function desensitize(node: QTreeNode, input: Inputs): Inputs {
  const copy = deepCopy(input);
  const names = new Set<string>();
  traverseToCollectPasswordNodes(node, names);
  for (const name of names) {
    copy[name] = "******";
  }
  return copy;
}

export function traverseToCollectPasswordNodes(node: QTreeNode, names: Set<string>): void {
  if (node.data.type === "text" && node.data.password === true) {
    names.add(node.data.name);
  }
  for (const child of node.children || []) {
    traverseToCollectPasswordNodes(child, names);
  }
}

//////V3 questions
export async function getQuestionsForAddFeature(
  inputs: v2.InputsWithProjectPath,
  solution: v3.ISolution,
  context: v2.Context
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (solution.getQuestionsForAddFeature) {
    const res = await solution.getQuestionsForAddFeature(context, inputs);
    return res;
  }
  return ok(undefined);
}

export async function getQuestionsForUserTaskV3(
  func: Func,
  inputs: Inputs,
  solution: v3.ISolution,
  context: v2.Context,
  envInfo: v2.DeepReadonly<v3.EnvInfoV3>
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (solution.getQuestionsForUserTask) {
    const res = await solution.getQuestionsForUserTask(
      context,
      inputs,
      func,
      envInfo,
      TOOLS.tokenProvider
    );
    return res;
  }
  return ok(undefined);
}

export async function getQuestionsForProvision(
  inputs: v2.InputsWithProjectPath,
  solution: v3.ISolution,
  context: v2.Context,
  envInfo: v2.DeepReadonly<v3.EnvInfoV3>
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (solution.getQuestionsForProvision) {
    const res = await solution.getQuestionsForProvision(
      context,
      inputs,
      envInfo,
      TOOLS.tokenProvider
    );
    return res;
  }
  return ok(undefined);
}

export async function getQuestionsForDeploy(
  inputs: v2.InputsWithProjectPath,
  solution: v3.ISolution,
  context: v2.Context,
  envInfo: v2.DeepReadonly<v3.EnvInfoV3>
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (solution.getQuestionsForDeploy) {
    const res = await solution.getQuestionsForDeploy(context, inputs, envInfo, TOOLS.tokenProvider);
    return res;
  }
  return ok(undefined);
}

export async function getQuestionsForPublish(
  inputs: v2.InputsWithProjectPath,
  solution: v3.ISolution,
  context: v2.Context,
  envInfo: v2.DeepReadonly<v3.EnvInfoV3>
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (solution.getQuestionsForPublish) {
    const res = await solution.getQuestionsForPublish(
      context,
      inputs,
      envInfo,
      TOOLS.tokenProvider.appStudioToken
    );
    return res;
  }
  return ok(undefined);
}

export async function getQuestionsForInit(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const node = new QTreeNode({ type: "group" });
  // no need to ask workspace folder for CLI.
  if (inputs.platform !== Platform.CLI) {
    node.addChild(new QTreeNode(QuestionRootFolder));
  }
  node.addChild(new QTreeNode(createAppNameQuestion(false)));
  const solution = Container.get<v3.ISolution>(BuiltInSolutionNames.azure);
  const context = createV2Context(newProjectSettings());
  if (solution.getQuestionsForInit) {
    const res = await solution.getQuestionsForInit(context, inputs);
    if (res.isErr()) return res;
    if (res.value) {
      const solutionNode = res.value as QTreeNode;
      if (solutionNode.data) node.addChild(solutionNode);
    }
  }
  return ok(node.trim());
}

export async function getQuestionsForCreateProjectV3(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const node = new QTreeNode(getCreateNewOrFromSampleQuestion(inputs.platform));
  // create new
  const createNew = new QTreeNode({ type: "group" });
  node.addChild(createNew);
  createNew.condition = { equals: ScratchOptionYes.id };

  // capabilities
  const capQuestion = createCapabilityQuestion();
  const capNode = new QTreeNode(capQuestion);
  createNew.addChild(capNode);
  const solution = Container.get<v3.ISolution>(BuiltInSolutionNames.azure);
  const context = createV2Context(newProjectSettings());
  if (solution.getQuestionsForInit) {
    const res = await solution.getQuestionsForInit(context, inputs);
    if (res.isErr()) return res;
    if (res.value) {
      const solutionNode = res.value as QTreeNode;
      if (solutionNode.data) capNode.addChild(solutionNode);
    }
  }
  const spfxPlugin = Container.get<SPFxPluginV3>(BuiltInFeaturePluginNames.spfx);
  const spfxRes = await spfxPlugin.getQuestionsForAddInstance(context, inputs);
  if (spfxRes.isOk()) {
    if (spfxRes.value?.data) {
      spfxRes.value.condition = { contains: TabSPFxItem.id };
      capNode.addChild(spfxRes.value);
    }
  }
  // Language
  const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
  programmingLanguage.condition = { minItems: 1 };
  createNew.addChild(programmingLanguage);

  // only CLI need folder input
  if (inputs.platform === Platform.CLI) {
    createNew.addChild(new QTreeNode(QuestionRootFolder));
  }
  createNew.addChild(new QTreeNode(createAppNameQuestion()));

  // create from sample
  const sampleNode = new QTreeNode(SampleSelect);
  node.addChild(sampleNode);
  sampleNode.condition = { equals: ScratchOptionNo.id };
  if (inputs.platform !== Platform.VSCode) {
    sampleNode.addChild(new QTreeNode(QuestionRootFolder));
  }
  return ok(node.trim());
}

async function getQuestionsForCreateM365ProjectV2(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const createNode = new QTreeNode(M365CreateFromScratchSelectQuestion);

  // app type
  createNode.addChild(new QTreeNode(M365AppTypeSelectQuestion));

  // capabilities
  createNode.addChild(new QTreeNode(M365CapabilitiesFuncQuestion));

  const globalSolutions: v2.SolutionPlugin[] = getAllSolutionPluginsV2();
  const context = createV2Context(newProjectSettings());
  for (const solutionPlugin of globalSolutions) {
    let res: Result<QTreeNode | QTreeNode[] | undefined, FxError> = ok(undefined);
    const v2plugin = solutionPlugin as v2.SolutionPlugin;
    res = v2plugin.getQuestionsForScaffolding
      ? await v2plugin.getQuestionsForScaffolding(context as v2.Context, inputs)
      : ok(undefined);
    if (res.isErr()) return err(new SystemError(res.error, CoreSource, "QuestionModelFail"));
    if (res.value) {
      const solutionNode = Array.isArray(res.value)
        ? (res.value as QTreeNode[])
        : [res.value as QTreeNode];
      for (const node of solutionNode) {
        if (node.data) createNode.addChild(node);
      }
    }
  }

  // programming language
  const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
  programmingLanguage.condition = { minItems: 1 };
  createNode.addChild(programmingLanguage);

  // only CLI need folder input
  if (CLIPlatforms.includes(inputs.platform)) {
    createNode.addChild(new QTreeNode(QuestionRootFolder));
  }

  // application name
  createNode.addChild(new QTreeNode(createAppNameQuestion()));

  return ok(createNode.trim());
}

//////V2 questions
export async function getQuestionsForCreateProjectV2(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (inputs.isM365) {
    return getQuestionsForCreateM365ProjectV2(inputs);
  }

  const node = new QTreeNode(getCreateNewOrFromSampleQuestion(inputs.platform));
  // create new
  const createNew = new QTreeNode({ type: "group" });
  node.addChild(createNew);
  createNew.condition = { equals: ScratchOptionYes.id };

  // capabilities
  const capQuestion = createCapabilityQuestion();
  const capNode = new QTreeNode(capQuestion);
  createNew.addChild(capNode);

  // create new M365
  const createNewM365 = new QTreeNode({ type: "group" });
  node.addChild(createNewM365);
  createNewM365.condition = { equals: ScratchOptionYesM365.id };

  // M365 app type
  createNewM365.addChild(new QTreeNode(M365AppTypeSelectQuestion));

  // M365 capabilities
  createNewM365.addChild(new QTreeNode(M365CapabilitiesFuncQuestion));

  const globalSolutions: v2.SolutionPlugin[] = await getAllSolutionPluginsV2();
  const context = createV2Context(newProjectSettings());
  for (const solutionPlugin of globalSolutions) {
    let res: Result<QTreeNode | QTreeNode[] | undefined, FxError> = ok(undefined);
    const v2plugin = solutionPlugin as v2.SolutionPlugin;
    res = v2plugin.getQuestionsForScaffolding
      ? await v2plugin.getQuestionsForScaffolding(context as v2.Context, inputs)
      : ok(undefined);
    if (res.isErr()) return err(new SystemError(res.error, CoreSource, "QuestionModelFail"));
    if (res.value) {
      const solutionNode = Array.isArray(res.value)
        ? (res.value as QTreeNode[])
        : [res.value as QTreeNode];
      for (const node of solutionNode) {
        if (node.data) {
          capNode.addChild(node);
          createNewM365.addChild(node);
        }
      }
    }
  }

  // Language
  const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
  programmingLanguage.condition = { minItems: 1 };
  createNew.addChild(programmingLanguage);
  createNewM365.addChild(programmingLanguage);

  // only CLI need folder input
  if (CLIPlatforms.includes(inputs.platform)) {
    createNew.addChild(new QTreeNode(QuestionRootFolder));
    createNewM365.addChild(new QTreeNode(QuestionRootFolder));
  }
  createNew.addChild(new QTreeNode(createAppNameQuestion()));
  createNewM365.addChild(new QTreeNode(createAppNameQuestion()));

  // create from sample
  const sampleNode = new QTreeNode(SampleSelect);
  node.addChild(sampleNode);
  sampleNode.condition = { equals: ScratchOptionNo.id };
  if (inputs.platform !== Platform.VSCode) {
    sampleNode.addChild(new QTreeNode(QuestionRootFolder));
  }
  return ok(node.trim());
}

export async function getQuestionsForUserTaskV2(
  ctx: SolutionContext | v2.Context,
  solution: Solution | v2.SolutionPlugin,
  func: FunctionRouter,
  inputs: Inputs,
  envInfo?: v2.EnvInfoV2
): Promise<Result<QTreeNode | undefined, FxError>> {
  const namespace = func.namespace;
  const array = namespace ? namespace.split("/") : [];
  if (namespace && "" !== namespace && array.length > 0) {
    let res: Result<QTreeNode | undefined, FxError> = ok(undefined);
    const solutionV2 = solution as v2.SolutionPlugin;
    if (solutionV2.getQuestionsForUserTask) {
      res = await solutionV2.getQuestionsForUserTask(
        ctx as v2.Context,
        inputs,
        func,
        envInfo!,
        TOOLS.tokenProvider
      );
    }
    if (res.isOk()) {
      if (res.value) {
        const node = res.value.trim();
        return ok(node);
      }
    }
    return res;
  }
  return err(FunctionRouterError(func));
}

export async function getQuestionsV2(
  ctx: SolutionContext | v2.Context,
  solution: Solution | v2.SolutionPlugin,
  stage: Stage,
  inputs: Inputs,
  envInfo?: v2.EnvInfoV2
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (stage !== Stage.create) {
    let res: Result<QTreeNode | undefined, FxError> = ok(undefined);
    const solutionV2 = solution as v2.SolutionPlugin;
    if (solutionV2.getQuestions) {
      inputs.stage = stage;
      res = await solutionV2.getQuestions(ctx as v2.Context, inputs, envInfo!, TOOLS.tokenProvider);
    }
    if (res.isErr()) return res;
    if (res.value) {
      const node = res.value as QTreeNode;
      if (node.data) {
        return ok(node.trim());
      }
    }
  }
  return ok(undefined);
}
