/* eslint-disable @typescript-eslint/ban-ts-comment */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { UserError } from "@microsoft/teamsfx-api";
import { GraphTokenProvider } from "@microsoft/teamsfx-api";
import { LogLevel } from "@azure/msal-node";
import { ExtensionErrors } from "../error";
import { CodeFlowLogin } from "./codeFlowLogin";
import VsCodeLogInstance from "./log";
import * as vscode from "vscode";
import { signedIn, signedOut } from "./common/constant";
import { login, LoginStatus } from "./common/login";
import { CryptoCachePlugin } from "./cacheAccess";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  AccountType,
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";

const accountName = "appStudio";
const scopes = ["Application.ReadWrite.All", "TeamsAppInstallation.ReadForUser"];

const cachePlugin = new CryptoCachePlugin(accountName);

const config = {
  auth: {
    clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
    authority: "https://login.microsoftonline.com/common",
  },
  system: {
    loggerOptions: {
      // @ts-ignore
      loggerCallback(loglevel, message, containsPii) {
        if (loglevel <= LogLevel.Error) {
          VsCodeLogInstance.error(message);
        }
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Error,
    },
  },
  cache: {
    cachePlugin,
  },
};

const SERVER_PORT = 0;

/**
 * use msal to implement graph login
 */
export class GraphLogin extends login implements GraphTokenProvider {
  private static instance: GraphLogin;

  private static codeFlowInstance: CodeFlowLogin;

  private static statusChange?: (
    status: string,
    token?: string,
    accountInfo?: Record<string, unknown>
  ) => Promise<void>;

  private constructor() {
    super();
    GraphLogin.codeFlowInstance = new CodeFlowLogin(scopes, config, SERVER_PORT, accountName);
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): GraphLogin {
    if (!GraphLogin.instance) {
      GraphLogin.instance = new GraphLogin();
    }

    return GraphLogin.instance;
  }

  async getAccessToken(showDialog = true): Promise<string | undefined> {
    await GraphLogin.codeFlowInstance.reloadCache();
    if (!GraphLogin.codeFlowInstance.account) {
      if (showDialog) {
        const userConfirmation: boolean = await this.doesUserConfirmLogin();
        if (!userConfirmation) {
          // throw user cancel error
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
            [TelemetryProperty.AccountType]: AccountType.M365,
            [TelemetryProperty.Success]: TelemetrySuccess.No,
            [TelemetryProperty.UserId]: "",
            [TelemetryProperty.Internal]: "",
            [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
            [TelemetryProperty.ErrorCode]: `${localize(
              "teamstoolkit.codeFlowLogin.loginComponent"
            )}.${ExtensionErrors.UserCancel}`,
            [TelemetryProperty.ErrorMessage]: `${localize("teamstoolkit.common.userCancel")}`,
          });
          throw new UserError(
            ExtensionErrors.UserCancel,
            localize("teamstoolkit.common.userCancel"),
            "Login"
          );
        }
      }
      const loginToken = await GraphLogin.codeFlowInstance.getToken();
      if (loginToken && GraphLogin.statusChange !== undefined) {
        const tokenJson = await this.getJsonObject();
        await GraphLogin.statusChange("SignedIn", loginToken, tokenJson);
      }
      await this.notifyStatus();
      return loginToken;
    }
    const accessToken = GraphLogin.codeFlowInstance.getToken();
    return accessToken;
  }

  async getJsonObject(showDialog = true): Promise<Record<string, unknown> | undefined> {
    const token = await this.getAccessToken();
    if (token) {
      const array = token.split(".");
      const buff = Buffer.from(array[1], "base64");
      return new Promise((resolve) => {
        resolve(JSON.parse(buff.toString("utf-8")));
      });
    } else {
      return new Promise((resolve) => {
        resolve(undefined);
      });
    }
  }

  async signout(): Promise<boolean> {
    GraphLogin.codeFlowInstance.account = undefined;
    if (GraphLogin.statusChange !== undefined) {
      await GraphLogin.statusChange("SignedOut", undefined, undefined);
    }
    await this.notifyStatus();
    return new Promise((resolve) => {
      resolve(true);
    });
  }

  private async doesUserConfirmLogin(): Promise<boolean> {
    const warningMsg = localize("teamstoolkit.graphLogin.warningMsg");
    const confirm = localize("teamstoolkit.common.confirm");
    const userSelected: string | undefined = await vscode.window.showWarningMessage(
      warningMsg,
      { modal: true },
      confirm
    );
    return Promise.resolve(userSelected === confirm);
  }

  async getStatus(): Promise<LoginStatus> {
    await GraphLogin.codeFlowInstance.reloadCache();
    if (GraphLogin.codeFlowInstance.account) {
      const loginToken = await GraphLogin.codeFlowInstance.getToken(false);
      const tokenJson = await this.getJsonObject();
      return Promise.resolve({ status: signedIn, token: loginToken, accountInfo: tokenJson });
    } else {
      return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
    }
  }
}

export default GraphLogin.getInstance();
