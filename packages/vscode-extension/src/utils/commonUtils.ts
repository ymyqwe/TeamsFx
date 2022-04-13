// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import * as os from "os";
import * as extensionPackage from "./../../package.json";
import * as fs from "fs-extra";
import { ext } from "../extensionVariables";
import * as path from "path";
import {
  ConfigFolderName,
  InputConfigsFolderName,
  ProjectSettingsFileName,
  EnvStateFileNameTemplate,
  StatesFolderName,
  EnvNamePlaceholder,
  Json,
  SubscriptionInfo,
} from "@microsoft/teamsfx-api";
import { environmentManager, isValidProject, PluginNames } from "@microsoft/teamsfx-core";
import { workspace, WorkspaceConfiguration } from "vscode";
import * as commonUtils from "../debug/commonUtils";
import { ConfigurationKey, CONFIGURATION_PREFIX, UserState } from "../constants";
import { execSync } from "child_process";
import * as versionUtil from "./versionUtil";
import { TelemetryTiggerFrom } from "../telemetry/extTelemetryEvents";

export function getPackageVersion(versionStr: string): string {
  if (versionStr.includes("alpha")) {
    return "alpha";
  }

  if (versionStr.includes("beta")) {
    return "beta";
  }

  if (versionStr.includes("rc")) {
    return "rc";
  }

  return "formal";
}

export function isFeatureFlag(): boolean {
  return extensionPackage.featureFlag === "true";
}

export async function sleep(ms: number) {
  await new Promise((resolve) => setTimeout(resolve, ms));

  await new Promise((resolve) => setTimeout(resolve, 0));
}

export function isWindows() {
  return os.type() === "Windows_NT";
}

export function isMacOS() {
  return os.type() === "Darwin";
}

export function isLinux() {
  return os.type() === "Linux";
}

export interface TeamsAppTelemetryInfo {
  appId: string;
  tenantId: string;
}

// Only used for telemetry when multi-env is enabled
export function getTeamsAppTelemetryInfoByEnv(env: string): TeamsAppTelemetryInfo | undefined {
  try {
    const ws = ext.workspaceUri.fsPath;

    if (isValidProject(ws)) {
      const result = environmentManager.getEnvStateFilesPath(env, ws);
      const envJson = JSON.parse(fs.readFileSync(result.envState, "utf8"));
      const appstudioState = envJson[PluginNames.APPST];
      return {
        appId: appstudioState.teamsAppId,
        tenantId: appstudioState.tenantId,
      };
    }
  } catch (e) {
    return undefined;
  }
}

export function getProjectId(): string | undefined {
  try {
    const ws = ext.workspaceUri.fsPath;
    const settingsJsonPathNew = path.join(
      ws,
      `.${ConfigFolderName}`,
      InputConfigsFolderName,
      ProjectSettingsFileName
    );
    const settingsJsonPathOld = path.join(ws, `.${ConfigFolderName}/settings.json`);

    // Do not check validity of project in multi-env.
    // Before migration, `isValidProject()` is false, but we still need to send `project-id` telemetry property.
    try {
      const settingsJson = JSON.parse(fs.readFileSync(settingsJsonPathNew, "utf8"));
      return settingsJson.projectId;
    } catch (e) {}

    // Also try reading from the old project location to support `ProjectMigratorMW` telemetry.
    // While doing migration, sending telemetry will call this `getProjectId()` function.
    // But before migration done, the settings file is still in the old location.
    const settingsJson = JSON.parse(fs.readFileSync(settingsJsonPathOld, "utf8"));
    return settingsJson.projectId;
  } catch (e) {
    return undefined;
  }
}

export async function isSPFxProject(workspacePath: string): Promise<boolean> {
  if (await fs.pathExists(`${workspacePath}/SPFx`)) {
    return true;
  }

  return false;
}

export function anonymizeFilePaths(stack?: string): string {
  if (stack === undefined || stack === null) {
    return "";
  }

  const cleanupPatterns: RegExp[] = [];

  let updatedStack = stack;

  const cleanUpIndexes: [number, number][] = [];

  for (const regexp of cleanupPatterns) {
    while (true) {
      const result = regexp.exec(stack);

      if (!result) {
        break;
      }

      cleanUpIndexes.push([result.index, regexp.lastIndex]);
    }
  }

  const nodeModulesRegex = /^[\\\/]?(node_modules|node_modules\.asar)[\\\/]/;

  const fileRegex =
    /(file:\/\/)?([a-zA-Z]:(\\\\|\\|\/)|(\\\\|\\|\/))?([\w-\._]+(\\\\|\\|\/))+[\w-\._]*/g;

  let lastIndex = 0;

  updatedStack = "";

  while (true) {
    const result = fileRegex.exec(stack);

    if (!result) {
      break;
    }

    // Anoynimize user file paths that do not need to be retained or cleaned up.

    if (
      !nodeModulesRegex.test(result[0]) &&
      cleanUpIndexes.every(([x, y]) => result.index < x || result.index >= y)
    ) {
      updatedStack += stack.substring(lastIndex, result.index) + "<REDACTED: user-file-path>";

      lastIndex = fileRegex.lastIndex;
    }
  }

  if (lastIndex < stack.length) {
    updatedStack += stack.substr(lastIndex);
  }

  // sanitize with configured cleanup patterns

  for (const regexp of cleanupPatterns) {
    updatedStack = updatedStack.replace(regexp, "");
  }

  return updatedStack;
}

export async function isTeamsfx(): Promise<boolean> {
  if (workspace.workspaceFolders && workspace.workspaceFolders.length > 0) {
    const workspaceFolder = workspace.workspaceFolders[0];

    return await commonUtils.isFxProject(workspaceFolder.uri.fsPath);
  }

  return false;
}

export function getConfiguration(key: string): boolean {
  const configuration: WorkspaceConfiguration = workspace.getConfiguration(CONFIGURATION_PREFIX);

  return configuration.get<boolean>(key, false);
}

export function syncFeatureFlags() {
  process.env["TEAMSFX_BICEP_ENV_CHECKER_ENABLE"] = getConfiguration(
    ConfigurationKey.BicepEnvCheckerEnable
  ).toString();

  process.env["TEAMSFX_ROOT_DIRECTORY"] = getConfiguration(
    ConfigurationKey.RootDirectory
  ).toString();

  process.env["TEAMSFX_CONFIG_UNIFY"] = getConfiguration(ConfigurationKey.UnifyConfigs).toString();

  process.env["TEAMSFX_INIT_APP"] = getConfiguration(ConfigurationKey.EnableInitApp).toString();
}

export class FeatureFlags {
  static readonly InsiderPreview = "__TEAMSFX_INSIDER_PREVIEW";
  static readonly TelemetryTest = "TEAMSFX_TELEMETRY_TEST";
}

// Determine whether feature flag is enabled based on environment variable setting

export function isFeatureFlagEnabled(featureFlagName: string, defaultValue = false): boolean {
  const flag = process.env[featureFlagName];

  if (flag === undefined) {
    return defaultValue; // allows consumer to set a default value when environment variable not set
  } else {
    return flag === "1" || flag.toLowerCase() === "true"; // can enable feature flag by set environment variable value to "1" or "true"
  }
}

export function getAllFeatureFlags(): string[] | undefined {
  const result = Object.values(FeatureFlags)

    .filter((featureFlag) => {
      return isFeatureFlagEnabled(featureFlag);
    })

    .map((featureFlag) => {
      return featureFlag;
    });

  return result;
}

export function getIsExistingUser(): string | undefined {
  return ext.context.globalState.get<string>(UserState.IsExisting);
}

export async function getSubscriptionInfoFromEnv(
  env: string
): Promise<SubscriptionInfo | undefined> {
  let provisionResult: Json | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.

    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  if (provisionResult.solution && provisionResult.solution.subscriptionId) {
    return {
      subscriptionName: provisionResult.solution.subscriptionName,

      subscriptionId: provisionResult.solution.subscriptionId,

      tenantId: provisionResult.solution.tenantId,
    };
  } else {
    return undefined;
  }
}

export async function getM365TenantFromEnv(env: string): Promise<string | undefined> {
  let provisionResult: Json | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.
    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  return provisionResult?.[PluginNames.SOLUTION]?.teamsAppTenantId;
}

export async function getResourceGroupNameFromEnv(env: string): Promise<string | undefined> {
  let provisionResult: Json | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.

    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  return provisionResult.solution.resourceGroupName;
}

export async function getProvisionSucceedFromEnv(env: string): Promise<boolean | undefined> {
  let provisionResult: Json | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.

    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  return provisionResult.solution?.provisionSucceeded;
}

async function getProvisionResultJson(env: string): Promise<Json | undefined> {
  if (vscode.workspace.workspaceFolders) {
    const workspaceFolder: vscode.WorkspaceFolder = vscode.workspace.workspaceFolders[0];

    const workspacePath: string = workspaceFolder.uri.fsPath;

    if (!(await commonUtils.isFxProject(workspacePath))) {
      return undefined;
    }

    const configRoot = await commonUtils.getProjectRoot(
      workspaceFolder.uri.fsPath,

      `.${ConfigFolderName}`
    );

    const provisionOutputFile = path.join(
      configRoot!,
      path.join(
        StatesFolderName,

        EnvStateFileNameTemplate.replace(EnvNamePlaceholder, env)
      )
    );

    if (!fs.existsSync(provisionOutputFile)) {
      return undefined;
    }

    const provisionResult = await fs.readJSON(provisionOutputFile);

    return provisionResult;
  }
}

export async function canUpgradeToArmAndMultiEnv(workspacePath?: string): Promise<boolean> {
  if (!workspacePath) return false;
  try {
    const fx = path.join(workspacePath, ".fx");
    if (!(await fs.pathExists(fx))) {
      return false;
    }
    const envFileExist = await fs.pathExists(path.join(fx, "env.default.json"));
    const configDirExist = await fs.pathExists(path.join(fx, "configs"));
    const armParameterExist = await fs.pathExists(
      path.join(fx, "configs", "azure.parameters.dev.json")
    );
    return envFileExist && (!armParameterExist || !configDirExist);
  } catch (err) {
    return false;
  }
}

export function isValidNode(): boolean {
  try {
    const supportedVersions = ["10", "12", "14"];
    const output = execSync("node --version");
    const regex = /v(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;

    const match = regex.exec(output.toString());
    if (!match) {
      return false;
    }

    const majorVersion = match.groups?.major_version;
    if (!majorVersion) {
      return false;
    }

    return supportedVersions.includes(majorVersion);
  } catch (e) {}
  return false;
}

export function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function isSupportAutoOpenAPI(): boolean {
  return versionUtil.compare(vscode.version, "1.64.2") > 0;
}

export function isTriggerFromWalkThrough(args?: any[]): boolean {
  if (!args || (args && args.length === 0)) {
    return false;
  } else if (
    args[0].toString() === TelemetryTiggerFrom.WalkThrough ||
    args[0].toString() === TelemetryTiggerFrom.Notification
  ) {
    return true;
  }

  return false;
}
