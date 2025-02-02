// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { VS_CODE_UI } from "../extension";
import graphLoginInstance from "../commonlib/graphLogin";
import * as commonUtils from "./commonUtils";
import { ExtensionErrors, ExtensionSource } from "../error";
import { localize } from "../utils/localizeUtils";
import { generateAccountHint } from "./teamsfxDebugProvider";

import axios from "axios";
import { returnSystemError, returnUserError } from "@microsoft/teamsfx-api";
import { environmentManager, isConfigUnifyEnabled } from "@microsoft/teamsfx-core";

export async function showInstallAppInTeamsMessage(detected: boolean): Promise<boolean> {
  const message = `${
    detected
      ? localize("teamstoolkit.localDebug.installApp.detection")
      : localize("teamstoolkit.localDebug.installApp.description")
  }${localize("teamstoolkit.localDebug.installApp.guide")}`;
  const result = await VS_CODE_UI.showMessage(
    "info",
    message,
    true,
    localize("teamstoolkit.localDebug.installApp.installInTeams"),
    localize("teamstoolkit.localDebug.installApp.continue")
  );
  if (result.isOk()) {
    if (result.value === localize("teamstoolkit.localDebug.installApp.cancel")) {
      return false;
    }

    const debugConfig = isConfigUnifyEnabled()
      ? await commonUtils.getDebugConfig(false, environmentManager.getLocalEnvName())
      : await commonUtils.getDebugConfig(true);
    if (debugConfig?.appId === undefined) {
      throw returnUserError(
        new Error("Debug config not found"),
        ExtensionSource,
        ExtensionErrors.GetTeamsAppInstallationFailed
      );
    }

    if (result.value === localize("teamstoolkit.localDebug.installApp.installInTeams")) {
      const url = `https://teams.microsoft.com/l/app/${
        debugConfig.appId
      }?installAppPackage=true&webjoin=true&${generateAccountHint()}`;
      await VS_CODE_UI.openUrl(url);
      return await showInstallAppInTeamsMessage(false);
    } else if (result.value === localize("teamstoolkit.localDebug.installApp.continue")) {
      const internalId = await getTeamsAppInternalId(debugConfig.appId);
      return internalId === undefined ? await showInstallAppInTeamsMessage(true) : true;
    }
  }
  return false;
}

export async function getTeamsAppInternalId(appId: string): Promise<string | undefined> {
  const loginStatus = await graphLoginInstance.getStatus();
  if (loginStatus.accountInfo?.oid === undefined || loginStatus.token === undefined) {
    throw returnUserError(
      new Error("M365 account info not found"),
      ExtensionSource,
      ExtensionErrors.GetTeamsAppInstallationFailed
    );
  }
  const url = `https://graph.microsoft.com/v1.0/users/${loginStatus.accountInfo.oid}/teamwork/installedApps?$expand=teamsApp,teamsAppDefinition&$filter=teamsApp/externalId eq '${appId}'`;
  try {
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${loginStatus.token}` },
    });
    for (const teamsAppInstallation of response.data.value) {
      if (teamsAppInstallation.teamsApp.distributionMethod === "sideloaded") {
        return teamsAppInstallation.teamsApp.id;
      }
    }
    return undefined;
  } catch (error: any) {
    throw returnSystemError(error, ExtensionSource, ExtensionErrors.GetTeamsAppInstallationFailed);
  }
}
