// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Aocheng Wang <aochengwang@microsoft.com>
 */

import { AppPackageFolderName, BuildFolderName } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import fs from "fs-extra";
import { describe } from "mocha";
import path from "path";
import AppStudioLogin from "../../../src/commonlib/appStudioLogin";
import {
  AadValidator,
  AppStudioValidator,
  BotValidator,
  FrontendValidator,
  FunctionValidator,
  SqlValidator,
} from "../../commonlib";
import { CliHelper } from "../../commonlib/cliHelper";
import {
  cleanUp,
  execAsync,
  execAsyncWithRetry,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  loadContext,
  mockTeamsfxMultiEnvFeatureFlag,
  setBotSkuNameToB1Bicep,
  setSimpleAuthSkuNameToB1Bicep,
} from "../commonUtils";

import { it } from "../../commonlib/it";

describe("Multi Env Happy Path for Azure", function () {
  const env = "e2e";
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const subscription = getSubscriptionId();
  const projectPath = path.resolve(testFolder, appName);
  const processEnv = mockTeamsfxMultiEnvFeatureFlag();

  it(
    `Can create/provision/deploy/build/validate/launch remote a azure tab/function/sql/bot project`,
    { testPlanCaseId: 10308408 },
    async function () {
      try {
        let result;
        result = await execAsync(
          `teamsfx new --interactive false --app-name ${appName} --capabilities tab bot --azure-resources function sql --programming-language javascript`,
          {
            cwd: testFolder,
            env: processEnv,
            timeout: 0,
          }
        );
        console.log(
          `[Successfully] scaffold to ${projectPath}, stdout: '${result.stdout}', stderr: '${result.stderr}''`
        );

        await CliHelper.setSubscription(subscription, projectPath, processEnv);

        // add env
        await CliHelper.addEnv(env, projectPath, processEnv);

        // update SKU from free to B1 to prevent free SKU limit error
        await setSimpleAuthSkuNameToB1Bicep(projectPath, env);
        await setBotSkuNameToB1Bicep(projectPath, env);
        console.log(`[Successfully] update simple auth sku to B1`);

        // list env
        result = await execAsync(`teamsfx env list`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });
        const envs = result.stdout.trim().split(/\r?\n/).sort();
        chai.expect(envs).to.deep.equal(["dev", "e2e"]);
        chai.expect(result.stderr).to.be.empty;
        console.log(
          `[Successfully] env list, stdout: '${result.stdout}', stderr: '${result.stderr}'`
        );

        // provision
        result = await execAsyncWithRetry(
          `teamsfx provision --sql-admin-name e2e --sql-password 'Abc123456%' --env ${env}`,
          {
            cwd: projectPath,
            env: processEnv,
            timeout: 0,
          }
        );
        console.log(
          `[Successfully] provision, stdout: '${result.stdout}', stderr: '${result.stderr}'`
        );

        let functionValidator: FunctionValidator;
        {
          // Validate provision
          // Get context
          const contextResult = await loadContext(projectPath, env);
          if (contextResult.isErr()) {
            throw contextResult.error;
          }
          const context = contextResult.value;

          // Validate Aad App
          const aad = AadValidator.init(context, false, AppStudioLogin);
          await AadValidator.validate(aad);

          // Validate Tab Frontend
          const frontend = FrontendValidator.init(context, true);
          await FrontendValidator.validateProvision(frontend);

          // Validate Function App
          functionValidator = new FunctionValidator(context, projectPath, env);
          await functionValidator.validateProvision();

          // Validate SQL
          await SqlValidator.init(context);
          await SqlValidator.validateSql();

          // Validate Bot Provision
          const bot = new BotValidator(context, projectPath, env);
          await bot.validateProvision();
        }

        // deploy
        await execAsyncWithRetry(`teamsfx deploy --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate deployment
          // Get context
          const contextResult = await loadContext(projectPath, env);
          if (contextResult.isErr()) {
            throw contextResult.error;
          }
          const context = contextResult.value;

          // Validate Tab Frontend
          const frontend = FrontendValidator.init(context, true);
          await FrontendValidator.validateDeploy(frontend);

          // Validate Function App
          await functionValidator.validateDeploy();

          // Validate Bot Deploy
          const bot = new BotValidator(context, projectPath, env);
          await bot.validateProvision();
        }

        // validate manifest
        result = await execAsyncWithRetry(`teamsfx validate --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate validate manifest
          chai.expect(result.stderr).to.be.empty;
        }

        // update manifest
        result = await execAsyncWithRetry(`teamsfx deploy manifest --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate update manifest
          chai.expect(result.stderr).to.be.empty;
        }

        // package
        await execAsyncWithRetry(`teamsfx package --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate package
          const file = `${projectPath}/${BuildFolderName}/${AppPackageFolderName}/appPackage.${env}.zip`;
          chai.expect(await fs.pathExists(file)).to.be.true;
        }

        // Temporarily disable publish
        // publish
        /*
        await execAsyncWithRetry(`teamsfx publish --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate publish result
          const contextResult = await loadContext(projectPath, env);
          if (contextResult.isErr()) {
            throw contextResult.error;
          }
          const context = contextResult.value;
          const aad = AadValidator.init(context, false, AppStudioLogin);
          const appId = aad.clientId;

          AppStudioValidator.init(context);
          await AppStudioValidator.validatePublish(appId);
        }*/
      } catch (e) {
        console.log("Unexpected exception is thrown when running test: " + e);
        console.log(e.stack);
        throw e;
      }
    }
  );

  after(async () => {
    // clean up
    await cleanUp(appName, projectPath, true, true, false, env);
  });
});
