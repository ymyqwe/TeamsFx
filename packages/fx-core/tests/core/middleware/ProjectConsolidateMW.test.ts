// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ConfigFolderName,
  FxError,
  InputConfigsFolderName,
  Inputs,
  ok,
  Platform,
  ProjectSettings,
  ProjectSettingsFileName,
  Result,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { setTools } from "../../../src";
import { ProjectConsolidateMW } from "../../../src/core/middleware/consolidateLocalRemote";
import { CoreHookContext } from "../../../src/core/types";
import { MockProjectSettings, MockTools, MockUserInteraction, randomAppName } from "../utils";
import * as projectSettingsLoader from "../../../src/core/middleware/projectSettingsLoader";
import * as projectMigrator from "../../../src/core/middleware/projectMigrator";
import { environmentManager } from "../../../src";

describe("Middleware - ProjectSettingsWriterMW", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(function () {
    sandbox.stub<any, any>(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith(".fx")) {
        return true;
      }
      return false;
    });
    setTools(new MockTools());
  });

  afterEach(function () {
    sandbox.restore();
  });

  it("not consolidate local remote", async () => {
    class MyClass {
      async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
        return ok("");
      }
    }

    hooks(MyClass, {
      myMethod: [ProjectConsolidateMW],
    });
    const my = new MyClass();
    const inputs1: Inputs = { platform: Platform.VSCode };
    await my.myMethod(inputs1);
    const inputs2: Inputs = {
      platform: Platform.CLI_HELP,
      projectPath: path.join(os.tmpdir(), randomAppName()),
    };
    await my.myMethod(inputs2);
  });

  it("consolidate happy path", async () => {
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("Upgrade"));
    const appName = randomAppName();
    const projectSettings: ProjectSettings = MockProjectSettings(appName);
    sandbox.stub(projectSettingsLoader, "loadProjectSettings").resolves(ok(projectSettings));
    sandbox.stub(environmentManager, "writeEnvConfig").resolves(ok(""));
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(projectMigrator, "addPathToGitignore").resolves();

    class MyClass {
      async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
        return ok("");
      }
    }

    hooks(MyClass, {
      myMethod: [ProjectConsolidateMW],
    });
    const my = new MyClass();
    const inputs1: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    await my.myMethod(inputs1);
    const inputs2: Inputs = {
      platform: Platform.CLI_HELP,
      projectPath: path.join(os.tmpdir(), appName),
    };
    await my.myMethod(inputs2);
  });
});
