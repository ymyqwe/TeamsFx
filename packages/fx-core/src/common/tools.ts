// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AppPackageFolderName,
  AzureAccountProvider,
  ConfigFolderName,
  ConfigMap,
  err,
  FxError,
  Json,
  ok,
  OptionItem,
  Result,
  returnSystemError,
  returnUserError,
  SubscriptionInfo,
  SystemError,
  UserInteraction,
  ProjectSettings,
  AzureSolutionSettings,
  v2,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import { exec, ExecOptions } from "child_process";
import * as fs from "fs-extra";
import * as Handlebars from "handlebars";
import * as path from "path";
import { promisify } from "util";
import * as uuid from "uuid";
import { getResourceFolder } from "../folder";
import {
  ConstantString,
  FeatureFlagName,
  TeamsClientId,
  OfficeClientId,
  OutlookClientId,
  ResourcePlugins,
} from "./constants";
import * as crypto from "crypto";
import * as os from "os";
import { FailedToParseResourceIdError } from "../core/error";
import { SolutionError } from "../plugins/solution/fx-solution/constants";
import Mustache from "mustache";
import {
  Component,
  sendTelemetryErrorEvent,
  sendTelemetryEvent,
  TelemetryEvent,
  TelemetryProperty,
} from "./telemetry";
import { HostTypeOptionAzure, SsoItem } from "../plugins/solution/fx-solution/question";
import { TOOLS } from "../core/globalVars";
import { LocalCrypto } from "../core/crypto";
import { getLocalizedString } from "./localizeUtils";

Handlebars.registerHelper("contains", (value, array) => {
  array = array instanceof Array ? array : [array];
  return array.indexOf(value) > -1 ? this : "";
});
Handlebars.registerHelper("notContains", (value, array) => {
  array = array instanceof Array ? array : [array];
  return array.indexOf(value) == -1 ? this : "";
});

export const Executor = {
  async execCommandAsync(command: string, options?: ExecOptions) {
    const execAsync = promisify(exec);
    return await execAsync(command, options);
  },
};

export async function npmInstall(path: string) {
  await Executor.execCommandAsync("npm install", {
    cwd: path,
  });
}

export async function ensureUniqueFolder(folderPath: string): Promise<string> {
  let folderId = 1;
  let testFolder = folderPath;

  let pathExists = await fs.pathExists(testFolder);
  while (pathExists) {
    testFolder = `${folderPath}${folderId}`;
    folderId++;

    pathExists = await fs.pathExists(testFolder);
  }

  return testFolder;
}

/**
 * Convert a `Map` to a Json recursively.
 * @param {Map} map to convert.
 * @returns {Json} converted Json.
 */
export function mapToJson(map?: Map<any, any>): Json {
  if (!map) return {};
  const out: Json = {};
  for (const entry of map.entries()) {
    if (entry[1] instanceof Map) {
      out[entry[0]] = mapToJson(entry[1]);
    } else {
      out[entry[0]] = entry[1];
    }
  }
  return out;
}

/**
 * Convert an `Object` to a Map recursively
 * @param {Json} Json to convert.
 * @returns {Map} converted Json.
 */
export function objectToMap(o: Json): Map<any, any> {
  const m = new Map();
  for (const entry of Object.entries(o)) {
    if (entry[1] instanceof Array) {
      m.set(entry[0], entry[1]);
    } else if (entry[1] instanceof Object) {
      m.set(entry[0], objectToConfigMap(entry[1] as Json));
    } else {
      m.set(entry[0], entry[1]);
    }
  }
  return m;
}

/**
 * @param {Json} Json to convert.
 * @returns {Map} converted Json.
 */
export function objectToConfigMap(o?: Json): ConfigMap {
  const m = new ConfigMap();
  if (o) {
    for (const entry of Object.entries(o)) {
      {
        m.set(entry[0], entry[1]);
      }
    }
  }
  return m;
}

const SecretDataMatchers = [
  "fx-resource-aad-app-for-teams.clientSecret",
  "fx-resource-simple-auth.filePath",
  "fx-resource-simple-auth.environmentVariableParams",
  "fx-resource-local-debug.*",
  "fx-resource-bot.botPassword",
  "fx-resource-apim.apimClientAADClientSecret",
  "fx-resource-azure-sql.adminPassword",
];

export const CryptoDataMatchers = new Set([
  "fx-resource-aad-app-for-teams.clientSecret",
  "fx-resource-aad-app-for-teams.local_clientSecret",
  "fx-resource-simple-auth.environmentVariableParams",
  "fx-resource-bot.botPassword",
  "fx-resource-bot.localBotPassword",
  "fx-resource-apim.apimClientAADClientSecret",
  "fx-resource-azure-sql.adminPassword",
]);

export const AzurePortalUrl = "https://portal.azure.com";

/**
 * Only data related to secrets need encryption.
 * @param key - the key name of data in user data file
 * @returns whether it needs encryption
 */
export function dataNeedEncryption(key: string): boolean {
  return CryptoDataMatchers.has(key);
}

export function separateSecretData(configJson: Json): Record<string, string> {
  const res: Record<string, string> = {};
  for (const matcher of SecretDataMatchers) {
    const splits = matcher.split(".");
    const resourceId = splits[0];
    const item = splits[1];
    const resourceConfig: any = configJson[resourceId];
    if (!resourceConfig) continue;
    if ("*" !== item) {
      const configValue = resourceConfig[item];
      if (configValue) {
        const keyName = `${resourceId}.${item}`;
        res[keyName] = configValue;
        resourceConfig[item] = `{{${keyName}}}`;
      }
    } else {
      for (const itemName of Object.keys(resourceConfig)) {
        const configValue = resourceConfig[itemName];
        if (configValue !== undefined) {
          const keyName = `${resourceId}.${itemName}`;
          res[keyName] = configValue;
          resourceConfig[itemName] = `{{${keyName}}}`;
        }
      }
    }
  }
  return res;
}

export function convertDotenvToEmbeddedJson(dict: Record<string, string>): Json {
  const result: Json = {};
  for (const key of Object.keys(dict)) {
    const array = key.split(".");
    let obj = result;
    for (let i = 0; i < array.length - 1; ++i) {
      const subKey = array[i];
      let subObj = obj[subKey];
      if (!subObj) {
        subObj = {};
        obj[subKey] = subObj;
      }
      obj = subObj;
    }
    obj[array[array.length - 1]] = dict[key];
  }
  return result;
}

export function replaceTemplateWithUserData(
  template: string,
  userData: Record<string, string>
): string {
  const view = convertDotenvToEmbeddedJson(userData);
  Mustache.escape = (t: string) => {
    if (!t) {
      return t;
    }
    const str = JSON.stringify(t);
    return str.substr(1, str.length - 2);
    // return t;
  };
  const result = Mustache.render(template, view);
  return result;
}

export function serializeDict(dict: Record<string, string>): string {
  const array: string[] = [];
  for (const key of Object.keys(dict)) {
    const value = dict[key];
    array.push(`${key}=${value}`);
  }
  return array.join("\n");
}

export const deepCopy = <T>(target: T): T => {
  if (target === null) {
    return target;
  }
  if (target instanceof Date) {
    return new Date(target.getTime()) as any;
  }
  if (target instanceof Array) {
    const cp = [] as any[];
    (target as any[]).forEach((v) => {
      cp.push(v);
    });
    return cp.map((n: any) => deepCopy<any>(n)) as any;
  }
  if (typeof target === "object" && target !== {}) {
    const cp = { ...(target as { [key: string]: any }) } as {
      [key: string]: any;
    };
    Object.keys(cp).forEach((k) => {
      cp[k] = deepCopy<any>(cp[k]);
    });
    return cp as T;
  }
  return target;
};

export function isUserCancelError(error: Error): boolean {
  const errorName = "name" in error ? (error as any)["name"] : "";
  return (
    errorName === "User Cancel" || errorName === "CancelProvision" || errorName === "UserCancel"
  );
}

export function isCheckAccountError(error: Error): boolean {
  const errorName = "name" in error ? (error as any)["name"] : "";
  return (
    errorName === SolutionError.TeamsAppTenantIdNotRight ||
    errorName === SolutionError.SubscriptionNotFound
  );
}

export async function askSubscription(
  azureAccountProvider: AzureAccountProvider,
  ui: UserInteraction,
  activeSubscriptionId?: string
): Promise<Result<SubscriptionInfo, FxError>> {
  const subscriptions: SubscriptionInfo[] = await azureAccountProvider.listSubscriptions();

  if (subscriptions.length === 0) {
    return err(
      returnUserError(new Error("Failed to find a subscription."), "Core", "NoSubscriptionFound")
    );
  }
  let resultSub = subscriptions.find((sub) => sub.subscriptionId === activeSubscriptionId);
  if (activeSubscriptionId === undefined || resultSub === undefined) {
    let selectedSub: SubscriptionInfo | undefined = undefined;
    if (subscriptions.length === 1) {
      selectedSub = subscriptions[0];
    } else {
      const options: OptionItem[] = subscriptions.map((sub) => {
        return {
          id: sub.subscriptionId,
          label: sub.subscriptionName,
          data: sub.tenantId,
        } as OptionItem;
      });
      const askRes = await ui.selectOption({
        name: "subscription",
        title: "Select a subscription",
        options: options,
        returnObject: true,
      });
      if (askRes.isErr()) return err(askRes.error);
      const subItem = askRes.value.result as OptionItem;
      selectedSub = {
        subscriptionId: subItem.id,
        subscriptionName: subItem.label,
        tenantId: subItem.data as string,
      };
    }
    if (selectedSub === undefined) {
      return err(
        returnSystemError(new Error("Subscription not found"), "Core", "NoSubscriptionFound")
      );
    }
    resultSub = selectedSub;
  }
  return ok(resultSub);
}

export function getResourceGroupInPortal(
  subscriptionId?: string,
  tenantId?: string,
  resourceGroupName?: string
): string | undefined {
  if (subscriptionId && tenantId && resourceGroupName) {
    return `${AzurePortalUrl}/#@${tenantId}/resource/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}`;
  } else {
    return undefined;
  }
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

/**
 * @deprecated Please DO NOT use this method any more, it will be removed in near future.
 */
export function isMultiEnvEnabled(): boolean {
  return true;
}

export function isBicepEnvCheckerEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.BicepEnvCheckerEnable, true);
}

export function isConfigUnifyEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.ConfigUnify, false);
}

export function isInitAppEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.EnableInitApp, false);
}

export function isAadManifestEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.AadManifest, false);
}

export function isM365AppEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.M365App, false);
}

// This method is for deciding whether AAD should be activated.
// Currently AAD plugin will always be activated when scaffold.
// This part will be updated when we support adding aad separately.
export function isAADEnabled(solutionSettings: AzureSolutionSettings): boolean {
  if (!solutionSettings) {
    return false;
  }

  if (isAadManifestEnabled()) {
    return (
      solutionSettings.hostType === HostTypeOptionAzure.id &&
      solutionSettings.capabilities.includes(SsoItem.id)
    );
  } else {
    return (
      solutionSettings.hostType === HostTypeOptionAzure.id &&
      // For scaffold, activeResourecPlugins is undefined
      (!solutionSettings.activeResourcePlugins ||
        solutionSettings.activeResourcePlugins?.includes(ResourcePlugins.Aad))
    );
  }
}

export function isBotNotificationEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.BotNotification, false);
}

export function getRootDirectory(): string {
  const root = process.env[FeatureFlagName.rootDirectory];
  if (root === undefined || root === "") {
    return path.join(os.homedir(), ConstantString.rootFolder);
  } else {
    return path.resolve(root.replace("${homeDir}", os.homedir()));
  }
}

export async function generateBicepFromFile(
  templateFilePath: string,
  context: any
): Promise<string> {
  try {
    const templateString = await fs.readFile(templateFilePath, ConstantString.UTF8Encoding);
    const updatedBicepFile = compileHandlebarsTemplateString(templateString, context);
    return updatedBicepFile;
  } catch (error) {
    throw returnSystemError(
      new Error(`Failed to generate bicep file ${templateFilePath}. Reason: ${error.message}`),
      "Core",
      "BicepGenerationError"
    );
  }
}

export function compileHandlebarsTemplateString(templateString: string, context: any): string {
  const template = Handlebars.compile(templateString);
  return template(context);
}

export async function getAppDirectory(projectRoot: string): Promise<string> {
  const REMOTE_MANIFEST = "manifest.source.json";
  const appDirNewLocForMultiEnv = `${projectRoot}/templates/${AppPackageFolderName}`;
  const appDirNewLoc = `${projectRoot}/${AppPackageFolderName}`;
  const appDirOldLoc = `${projectRoot}/.${ConfigFolderName}`;

  if (await fs.pathExists(`${appDirNewLocForMultiEnv}`)) {
    return appDirNewLocForMultiEnv;
  } else if (await fs.pathExists(`${appDirNewLoc}/${REMOTE_MANIFEST}`)) {
    return appDirNewLoc;
  } else {
    return appDirOldLoc;
  }
}

/**
 * Get app studio endpoint for prod/int environment, mainly for ux e2e test
 */
export function getAppStudioEndpoint(): string {
  if (process.env.APP_STUDIO_ENV && process.env.APP_STUDIO_ENV === "int") {
    return "https://dev-int.teams.microsoft.com";
  } else {
    return "https://dev.teams.microsoft.com";
  }
}

export function getStorageAccountNameFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(
    /providers\/Microsoft.Storage\/storageAccounts\/([^\/]*)/i,
    resourceId
  );
  if (!result) {
    throw FailedToParseResourceIdError("storage accounts name", resourceId);
  }
  return result;
}

export function getSiteNameFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(/providers\/Microsoft.Web\/sites\/([^\/]*)/i, resourceId);
  if (!result) {
    throw FailedToParseResourceIdError("site name", resourceId);
  }
  return result;
}

export function getResourceGroupNameFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(/\/resourceGroups\/([^\/]*)\//i, resourceId);
  if (!result) {
    throw FailedToParseResourceIdError("resource group name", resourceId);
  }
  return result;
}

export function getSubscriptionIdFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(/\/subscriptions\/([^\/]*)\//i, resourceId);
  if (!result) {
    throw FailedToParseResourceIdError("subscription id", resourceId);
  }
  return result;
}

export function parseFromResourceId(pattern: RegExp, resourceId: string): string {
  const result = resourceId.match(pattern);
  return result ? result[1].trim() : "";
}

export async function waitSeconds(second: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, second * 1000));
}

export function getUuid(): string {
  return uuid.v4();
}

export function isSPFxProject(projectSettings?: ProjectSettings): boolean {
  const solutionSettings = projectSettings?.solutionSettings as AzureSolutionSettings;
  if (solutionSettings) {
    const selectedPlugins = solutionSettings.activeResourcePlugins;
    return selectedPlugins && selectedPlugins.indexOf("fx-resource-spfx") !== -1;
  }
  return false;
}

export function getHashedEnv(envName: string): string {
  return crypto.createHash("sha256").update(envName).digest("hex");
}

export function IsSimpleAuthEnabled(projectSettings: ProjectSettings | undefined): boolean {
  const solutionSettings = projectSettings?.solutionSettings as AzureSolutionSettings;
  return solutionSettings?.activeResourcePlugins?.includes(ResourcePlugins.SimpleAuth);
}

interface BasicJsonSchema {
  type: string;
  properties?: {
    [k: string]: unknown;
  };
}
function isBasicJsonSchema(jsonSchema: unknown): jsonSchema is BasicJsonSchema {
  if (!jsonSchema || typeof jsonSchema !== "object") {
    return false;
  }
  return typeof (jsonSchema as { type: unknown })["type"] === "string";
}

function _redactObject(
  obj: unknown,
  jsonSchema: unknown,
  maxRecursionDepth = 8,
  depth = 0
): unknown {
  if (depth >= maxRecursionDepth) {
    // prevent stack overflow if anything bad happens
    return null;
  }
  if (!obj || !isBasicJsonSchema(jsonSchema)) {
    return null;
  }

  if (
    !(
      jsonSchema.type === "object" &&
      jsonSchema.properties &&
      typeof jsonSchema.properties === "object"
    )
  ) {
    // non-object types including unsupported types
    return null;
  }

  const newObj: { [key: string]: any } = {};
  const objAny = obj as any;
  for (const key in jsonSchema.properties) {
    if (key in objAny && objAny[key] !== undefined) {
      const filteredObj = _redactObject(
        objAny[key],
        jsonSchema.properties[key],
        maxRecursionDepth,
        depth + 1
      );
      newObj[key] = filteredObj;
    }
  }
  return newObj;
}

/** Redact user content in "obj";
 *
 * DFS "obj" and "jsonSchema" together to redact the following things:
 * - properties that is not defined in jsonSchema
 * - the value of properties that is defined in jsonSchema, but the keys will remain
 *
 * Example:
 * Input:
 * ```
 *  obj = {
 *    "name": "some name",
 *    "user defined property": {
 *      "key1": "value1"
 *    }
 *  }
 *  jsonSchema = {
 *    "type": "object",
 *    "properties": {
 *      "name": { "type": "string" }
 *    }
 *  }
 * ```
 * Output:
 * ```
 *  {"name": null}
 * ```
 **/
export function redactObject(obj: unknown, jsonSchema: unknown, maxRecursionDepth = 8): unknown {
  return _redactObject(obj, jsonSchema, maxRecursionDepth, 0);
}

export function getAllowedAppIds(): string[] {
  return [
    TeamsClientId.MobileDesktop,
    TeamsClientId.Web,
    OfficeClientId.Desktop,
    OfficeClientId.Web1,
    OfficeClientId.Web2,
    OutlookClientId.Desktop,
    OutlookClientId.Web1,
    OutlookClientId.Web2,
  ];
}

export function getAllowedAppMaps(): Record<string, string> {
  return {
    [TeamsClientId.MobileDesktop]: getLocalizedString("core.common.TeamsMobileDesktopClientName"),
    [TeamsClientId.Web]: getLocalizedString("core.common.TeamsWebClientName"),
    [OfficeClientId.Desktop]: getLocalizedString("core.common.OfficeDesktopClientName"),
    [OfficeClientId.Web1]: getLocalizedString("core.common.OfficeWebClientName1"),
    [OfficeClientId.Web2]: getLocalizedString("core.common.OfficeWebClientName2"),
    [OutlookClientId.Desktop]: getLocalizedString("core.common.OutlookDesktopClientName"),
    [OutlookClientId.Web1]: getLocalizedString("core.common.OutlookWebClientName1"),
    [OutlookClientId.Web2]: getLocalizedString("core.common.OutlookWebClientName2"),
  };
}

export async function getSideloadingStatus(token: string): Promise<boolean | undefined> {
  const instance = axios.create({
    baseURL: getAppStudioEndpoint(),
    timeout: 30000,
  });
  instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;

  let retry = 0;
  const retryIntervalSeconds = 2;
  do {
    try {
      const response = await instance.get("/api/usersettings/mtUserAppPolicy");
      let result: boolean | undefined;
      if (response.status >= 400) {
        result = undefined;
      } else {
        result = response.data?.value?.isSideloadingAllowed as boolean;
      }

      if (result !== undefined) {
        sendTelemetryEvent(Component.core, TelemetryEvent.CheckSideloading, {
          [TelemetryProperty.IsSideloadingAllowed]: result + "",
        });
      } else {
        sendTelemetryErrorEvent(
          Component.core,
          TelemetryEvent.CheckSideloading,
          new SystemError(
            "UnknownValue",
            `AppStudio response code: ${response.status}, body: ${response.data}`,
            "M365Account"
          )
        );
      }

      return result;
    } catch (error) {
      sendTelemetryErrorEvent(
        Component.core,
        TelemetryEvent.CheckSideloading,
        new SystemError(error as Error, "M365Account")
      );
      await waitSeconds((retry + 1) * retryIntervalSeconds);
    }
  } while (++retry < 3);

  return undefined;
}

export function createV2Context(projectSettings: ProjectSettings): v2.Context {
  const context: v2.Context = {
    userInteraction: TOOLS.ui,
    logProvider: TOOLS.logProvider,
    telemetryReporter: TOOLS.telemetryReporter!,
    cryptoProvider: new LocalCrypto(projectSettings.projectId),
    permissionRequestProvider: TOOLS.permissionRequest,
    projectSetting: projectSettings,
  };
  return context;
}

export function undefinedName(objs: any[], names: string[]) {
  for (let i = 0; i < objs.length; ++i) {
    if (objs[i] === undefined) {
      return names[i];
    }
  }
  return undefined;
}
