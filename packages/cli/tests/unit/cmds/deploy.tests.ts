// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import sinon from "sinon";
import yargs, { Options, PositionalOptions } from "yargs";

import { err, Func, FxError, Inputs, ok, QTreeNode, UserError } from "@microsoft/teamsfx-api";
import { FxCore } from "@microsoft/teamsfx-core";

import Deploy from "../../../src/cmds/deploy";
import CliTelemetry from "../../../src/telemetry/cliTelemetry";
import { TelemetryEvent } from "../../../src/telemetry/cliTelemetryEvents";
import HelpParamGenerator from "../../../src/helpParamGenerator";
import * as constants from "../../../src/constants";
import { expect } from "../utils";
import { NotSupportedProjectType } from "../../../src/error";
import UI from "../../../src/userInteraction";
import LogProvider from "../../../src/commonlib/log";
import { iteratee } from "lodash";

describe("Deploy Command Tests", function () {
  const sandbox = sinon.createSandbox();
  let telemetryEvents: string[] = [];
  let options: string[] = [];
  let positionals: string[] = [];
  let allArguments = new Map<string, any>();
  const params = {
    [constants.deployPluginNodeName]: {
      choices: ["a", "b", "c"],
      description: "deployPluginNodeName",
    },
    "open-api-document": {},
    "api-prefix": {},
    "api-version": {},
    "skip-manifest": {},
  };

  before(() => {
    sandbox.stub(HelpParamGenerator, "getYargsParamForHelp").callsFake(() => {
      return params;
    });
    sandbox.stub(HelpParamGenerator, "getQuestionRootNodeForHelp").callsFake(() => {
      return new QTreeNode({
        name: constants.deployPluginNodeName,
        type: "multiSelect",
        title: "deployPluginNodeName",
        staticOptions: ["a", "b", "c"],
      });
    });
    sandbox.stub(yargs, "options").callsFake((ops: { [key: string]: Options }) => {
      if (typeof ops === "string") {
        options.push(ops);
      } else {
        options = options.concat(...Object.keys(ops));
      }
      return yargs;
    });
    sandbox.stub(yargs, "positional").callsFake((name: string) => {
      positionals.push(name);
      return yargs;
    });
    sandbox.stub(yargs, "exit").callsFake((code: number, err: Error) => {
      throw err;
    });
    sandbox.stub(CliTelemetry, "sendTelemetryEvent").callsFake((eventName: string) => {
      telemetryEvents.push(eventName);
    });
    sandbox
      .stub(CliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName: string, error: FxError) => {
        telemetryEvents.push(eventName);
      });
    sandbox.stub(FxCore.prototype, "deployArtifacts").callsFake(async (inputs: Inputs) => {
      if (inputs.projectPath?.includes("real")) return ok("");
      else return err(NotSupportedProjectType());
    });
    sandbox.stub(UI, "updatePresetAnswer").callsFake((key: any, value: any) => {
      allArguments.set(key, value);
    });
    sandbox.stub(LogProvider, "necessaryLog").returns();
  });

  after(() => {
    sandbox.restore();
  });

  beforeEach(() => {
    telemetryEvents = [];
    options = [];
    positionals = [];
    allArguments = new Map<string, any>();
  });

  it("Builder Check", () => {
    const cmd = new Deploy();
    cmd.builder(yargs);
    expect(options).deep.equals(
      ["open-api-document", "api-prefix", "api-version", "skip-manifest"],
      JSON.stringify(options)
    );
    expect(positionals).deep.equals(["components"], JSON.stringify(positionals));
  });

  it("Deploy Command Running -- no components", async () => {
    const cmd = new Deploy();
    cmd["params"] = params;
    const args = {
      [constants.RootFolderNode.data.name as string]: "real",
    };
    await cmd.handler(args);
    expect(allArguments.get("open-api-document")).equals(undefined);
    expect(allArguments.get("api-prefix")).equals(undefined);
    expect(allArguments.get("api-version")).equals(undefined);
    expect(allArguments.get("skip-manifest")).equals(undefined);
    expect(telemetryEvents).deep.equals([TelemetryEvent.DeployStart, TelemetryEvent.Deploy]);
  });

  it("Deploy Command Running -- 1 component", async () => {
    const cmd = new Deploy();
    cmd["params"] = params;
    const args = {
      [constants.RootFolderNode.data.name as string]: "real",
      components: ["a"],
    };
    await cmd.handler(args);
    expect(telemetryEvents).deep.equals([TelemetryEvent.DeployStart, TelemetryEvent.Deploy]);
  });

  it("Deploy Command Running -- deployArtifacts error", async () => {
    const cmd = new Deploy();
    cmd["params"] = params;
    const args = {
      [constants.RootFolderNode.data.name as string]: "fake",
    };
    try {
      await cmd.handler(args);
    } catch (e) {
      expect(telemetryEvents).deep.equals([TelemetryEvent.DeployStart, TelemetryEvent.Deploy]);
      expect(e).instanceOf(UserError);
      expect(e.name).equals("NotSupportedProjectType");
    }
  });
});

describe("Deploy manifest", function () {
  const sandbox = sinon.createSandbox();
  let telemetryEvents: string[] = [];
  let options: string[] = [];
  let positionals: { name: string; opt: PositionalOptions }[] = [];
  let allArguments = new Map<string, any>();
  const params = {
    [constants.deployPluginNodeName]: {
      choices: ["a", "b", "c", "appstudio"],
      description: "deployPluginNodeName",
    },
    "open-api-document": {},
    "api-prefix": {},
    "api-version": {},
    "skip-manifest": {},
  };

  before(() => {
    sandbox.stub(HelpParamGenerator, "getYargsParamForHelp").callsFake(() => {
      return params;
    });
    sandbox.stub(HelpParamGenerator, "getQuestionRootNodeForHelp").callsFake(() => {
      return new QTreeNode({
        name: constants.deployPluginNodeName,
        type: "multiSelect",
        title: "deployPluginNodeName",
        staticOptions: ["a", "b", "c"],
      });
    });
    sandbox.stub(yargs, "options").callsFake((ops: { [key: string]: Options }) => {
      if (typeof ops === "string") {
        options.push(ops);
      } else {
        options = options.concat(...Object.keys(ops));
      }
      return yargs;
    });
    sandbox.stub(yargs, "positional").callsFake((name: string, opt: PositionalOptions) => {
      positionals.push({ name, opt });
      return yargs;
    });
    sandbox.stub(yargs, "exit").callsFake((code: number, err: Error) => {
      throw err;
    });
    sandbox.stub(CliTelemetry, "sendTelemetryEvent").callsFake((eventName: string) => {
      telemetryEvents.push(eventName);
    });
    sandbox
      .stub(CliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName: string, error: FxError) => {
        telemetryEvents.push(eventName);
      });
    sandbox.stub(FxCore.prototype, "deployArtifacts").callsFake(async (inputs: Inputs) => {
      if (inputs.projectPath?.includes("real")) return ok("");
      else return err(NotSupportedProjectType());
    });
    sandbox.stub(UI, "updatePresetAnswer").callsFake((key: any, value: any) => {
      allArguments.set(key, value);
    });
    sandbox.stub(LogProvider, "necessaryLog").returns();
  });

  after(() => {
    sandbox.restore();
  });

  beforeEach(() => {
    telemetryEvents = [];
    options = [];
    positionals = [];
    allArguments = new Map<string, any>();
  });

  it("should rename appstudio to manifest", () => {
    const cmd = new Deploy();
    cmd.builder(yargs);
    expect(options).deep.equals(
      ["open-api-document", "api-prefix", "api-version", "skip-manifest"],
      JSON.stringify(options)
    );
    expect(positionals.length).equals(1);
    expect(positionals[0].name).equals("components");
    expect(positionals[0].opt.choices).deep.equals(["a", "b", "c", "manifest"]);
  });
});
