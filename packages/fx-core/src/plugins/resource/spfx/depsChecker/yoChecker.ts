// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as os from "os";
import {
  ConfigFolderName,
  err,
  FxError,
  LogProvider,
  ok,
  PluginContext,
  Result,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { DependencyChecker } from "./dependencyChecker";
import { telemetryHelper } from "../utils/telemetry-helper";
import { TelemetryEvents, TelemetryProperty } from "../utils/telemetryEvents";
import { DependencyValidateError, NpmInstallError, NpmNotFoundError } from "../error";
import { cpUtils } from "../../../../common/deps-checker/util/cpUtils";

const name = "yo";
const supportedVersion = "4.3.0";
const displayName = `${name}@${supportedVersion}`;
const timeout = 5 * 60 * 1000;

export class YoChecker implements DependencyChecker {
  private readonly _logger: LogProvider;

  constructor(logger: LogProvider) {
    this._logger = logger;
  }

  public async ensureDependency(ctx: PluginContext): Promise<Result<boolean, FxError>> {
    telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureYoStart);
    try {
      if (!(await this.isInstalled())) {
        this._logger.info(`Installing ${displayName}...`);
        await this.install();
        this._logger.info(`Successfully installed ${displayName}`);
      }
      telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureYo);
    } catch (error) {
      telemetryHelper.sendErrorEvent(
        ctx,
        TelemetryEvents.EnsureYo,
        error as UserError | SystemError,
        { [TelemetryProperty.EnsureYoReason]: (error as UserError | SystemError).name }
      );
      await this._logger.error(`Failed to install 'yo', error = '${error}'`);
      return err(error as UserError | SystemError);
    }

    return ok(true);
  }

  public async isInstalled(): Promise<boolean> {
    let isVersionSupported = false,
      hasSentinel = false;
    try {
      const yoVersion = await this.queryVersion();
      isVersionSupported = yoVersion !== undefined && supportedVersion === yoVersion;
      hasSentinel = await fs.pathExists(this.getSentinelPath());
    } catch (error) {
      return false;
    }
    return isVersionSupported && hasSentinel;
  }

  public async install(): Promise<void> {
    if (!(await this.hasNPM())) {
      this._logger.error("Failed to install yo since npm is not found");
      throw NpmNotFoundError();
    }

    await this.cleanup();
    await this.installYo();

    if (!(await this.validate())) {
      this._logger.debug("Failed to validate yo, cleaning up...");
      await this.cleanup();
      throw DependencyValidateError(name);
    }
  }

  public getBinFolder(): string {
    return this.getDefaultInstallPath();
  }

  private async validate(): Promise<boolean> {
    return await this.isInstalled();
  }

  private getDefaultInstallPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "yo");
  }

  private getSentinelPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "yo-sentinel");
  }

  private async queryVersion(): Promise<string | undefined> {
    const packagePath = path.join(
      this.getDefaultInstallPath(),
      "node_modules",
      "yo",
      "package.json"
    );
    if (await fs.pathExists(packagePath)) {
      const packageJson = await fs.readJson(packagePath);
      return packageJson.version ?? undefined;
    }
    return undefined;
  }

  private async hasNPM(): Promise<boolean> {
    try {
      await cpUtils.executeCommand(undefined, this._logger, { shell: true }, "npm", "--version");
      return true;
    } catch (error) {
      return false;
    }
  }

  private async cleanup(): Promise<void> {
    try {
      await fs.emptyDir(this.getDefaultInstallPath());
      await fs.remove(this.getSentinelPath());
    } catch (err) {
      await this._logger.error(
        `Failed to clean up path: ${this.getDefaultInstallPath()}, error: ${err}`
      );
    }
  }

  private async installYo(): Promise<void> {
    try {
      await cpUtils.executeCommand(
        undefined,
        this._logger,
        { timeout: timeout, shell: false },
        this.getExecCommand("npm"),
        "install",
        `${name}@${supportedVersion}`,
        "--prefix",
        `${this.getDefaultInstallPath()}`,
        "--no-audit"
      );

      await fs.ensureFile(this.getSentinelPath());
    } catch (error) {
      this._logger.error("Failed to npm install yo");
      throw NpmInstallError(error as Error);
    }
  }

  private getExecCommand(command: string): string {
    return this.isWindows() ? `${command}.cmd` : command;
  }

  private isWindows(): boolean {
    return os.type() === "Windows_NT";
  }
}
