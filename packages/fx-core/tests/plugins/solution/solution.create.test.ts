// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { it } from "mocha";
import { TeamsAppSolution } from " ../../../src/plugins/solution";
import { Platform, SolutionContext, SolutionSettings } from "@microsoft/teamsfx-api";
import * as sinon from "sinon";
import fs, { PathLike } from "fs-extra";
import {
  GLOBAL_CONFIG,
  PROGRAMMING_LANGUAGE,
  SolutionError,
} from "../../../src/plugins/solution/fx-solution/constants";
import {
  AzureSolutionQuestionNames,
  BotOptionItem,
  TabOptionItem,
} from "../../../src/plugins/solution/fx-solution/question";
import * as uuid from "uuid";
import { newEnvInfo } from "../../../src";
import { LocalCrypto } from "../../../src/core/crypto";
import { SolutionRunningState } from "../../../src/plugins/solution/fx-solution/types";
const tool = require("../../../src/common/tools");

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("Solution running state on creation", () => {
  const solution = new TeamsAppSolution();
  it("should be idle", () => {
    expect(solution.runningState).equal(SolutionRunningState.Idle);
  });
});

describe("Solution create()", async () => {
  function mockSolutionContext(): SolutionContext {
    return {
      root: ".",
      envInfo: newEnvInfo(),
      answers: { platform: Platform.VSCode },
      projectSettings: undefined,
      cryptoProvider: new LocalCrypto(""),
    };
  }

  const mocker = sinon.createSandbox();
  const permissionsJsonPath = "./permissions.json";
  const fileContent: Map<string, any> = new Map();
  beforeEach(() => {
    mocker.stub(fs, "writeFile").callsFake((path: number | PathLike, data: any) => {
      fileContent.set(path.toString(), data);
    });
    // mocker.stub(fs, "writeFile").resolves();
    mocker.stub(fs, "writeJSON").callsFake((file: string, obj: any) => {
      fileContent.set(file, JSON.stringify(obj));
    });
    // Uses stub<any, any> to circumvent type check. Beacuse sinon fails to mock my target overload of readJson.
    mocker.stub<any, any>(fs, "readJson").withArgs(permissionsJsonPath).resolves({});
    mocker.stub<any, any>(fs, "pathExists").withArgs(permissionsJsonPath).resolves(true);
    mocker.stub<any, any>(fs, "copy").resolves();
  });

  it("should fail if projectSettings is undefined", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    // expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
    expect(mockedSolutionCtx.envInfo.state.get(GLOBAL_CONFIG)).to.be.not.undefined;
  });

  it("should fail if projectSettings.solutionSettings is undefined", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: undefined as unknown as SolutionSettings,
    };
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
  });

  it("should fail if capability is empty", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
  });

  it("should succeed if projectSettings, solution settings and capabilities are provided", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities] = [BotOptionItem.id];
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);
    expect(mockedSolutionCtx.envInfo.state.get(GLOBAL_CONFIG)).is.not.undefined;
  });

  it("should set programmingLanguage in config if programmingLanguage is in answers", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      programmingLanguage: "",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities as string] = [BotOptionItem.id];
    const programmingLanguage = "TypeScript";
    answers[AzureSolutionQuestionNames.ProgrammingLanguage as string] = programmingLanguage;
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);

    const lang = mockedSolutionCtx.projectSettings.programmingLanguage;
    expect(lang).equals(programmingLanguage);
  });

  it("shouldn't set programmingLanguage in config if programmingLanguage is not in answers", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities as string] = [BotOptionItem.id];
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);
    const lang = mockedSolutionCtx.envInfo.state
      .get(GLOBAL_CONFIG)
      ?.getString(PROGRAMMING_LANGUAGE);
    expect(lang).to.be.undefined;
  });

  afterEach(() => {
    mocker.restore();
  });
});

describe("Solution create() with AAD manifest enabled", async () => {
  function mockSolutionContext(): SolutionContext {
    return {
      root: ".",
      envInfo: newEnvInfo(),
      answers: { platform: Platform.VSCode },
      projectSettings: undefined,
      cryptoProvider: new LocalCrypto(""),
    };
  }

  const mocker = sinon.createSandbox();
  const fileContent: Map<string, any> = new Map();
  beforeEach(() => {
    mocker.stub(fs, "writeFile").callsFake((path: number | PathLike, data: any) => {
      fileContent.set(path.toString(), data);
    });
    // mocker.stub(fs, "writeFile").resolves();
    mocker.stub(fs, "writeJSON").callsFake((file: string, obj: any) => {
      fileContent.set(file, JSON.stringify(obj));
    });
    mocker.stub<any, any>(tool, "isAadManifestEnabled").returns(true);
  });

  it("should fail if projectSettings is undefined", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    // expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
    expect(mockedSolutionCtx.envInfo.state.get(GLOBAL_CONFIG)).to.be.not.undefined;
  });

  it("should fail if projectSettings.solutionSettings is undefined", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: undefined as unknown as SolutionSettings,
    };
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
  });

  it("should fail if capability is empty", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isErr()).equals(true);
    expect(result._unsafeUnwrapErr().name).equals(SolutionError.InternelError);
  });

  it("should succeed if projectSettings, solution settings and capabilities are provided", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities] = [BotOptionItem.id];
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);
    expect(mockedSolutionCtx.envInfo.state.get(GLOBAL_CONFIG)).is.not.undefined;
  });

  it("should set programmingLanguage in config if programmingLanguage is in answers", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      programmingLanguage: "",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities as string] = [BotOptionItem.id];
    const programmingLanguage = "TypeScript";
    answers[AzureSolutionQuestionNames.ProgrammingLanguage as string] = programmingLanguage;
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);

    const lang = mockedSolutionCtx.projectSettings.programmingLanguage;
    expect(lang).equals(programmingLanguage);
  });

  it("shouldn't set programmingLanguage in config if programmingLanguage is not in answers", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities as string] = [BotOptionItem.id];
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);
    const lang = mockedSolutionCtx.envInfo.state
      .get(GLOBAL_CONFIG)
      ?.getString(PROGRAMMING_LANGUAGE);
    expect(lang).to.be.undefined;
  });

  it("shouldn't create permissions.json", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedSolutionCtx = mockSolutionContext();
    mockedSolutionCtx.projectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: "azure",
        version: "1.0",
      },
    };
    const answers = mockedSolutionCtx.answers!;
    answers[AzureSolutionQuestionNames.Capabilities] = [TabOptionItem.id];
    const permissionsJsonPath = "./permissions.json";
    const result = await solution.create(mockedSolutionCtx);
    expect(result.isOk()).equals(true);
    expect(await fs.pathExists(permissionsJsonPath)).equals(false);
  });

  afterEach(() => {
    mocker.restore();
  });
});
