import { execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

import type { BrowserChannel } from "../config.js";
import type { Logger } from "../logger.js";

const execFileAsync = promisify(execFile);

const PROBE_INTERVAL_MS = 250;
const PROBE_TIMEOUT_MS = 1_000;
const START_TIMEOUT_MS = 30_000;
const LOCK_TIMEOUT_MS = 60_000;
const LOCK_STALE_MS = 90_000;

export interface SharedEdgeSettings {
  readonly profileDir: string;
  readonly cdpPort: number;
  readonly headless: boolean;
  readonly browserChannel: BrowserChannel;
  readonly edgeExecutablePath: string | undefined;
}

export interface SharedEdgeState {
  readonly pid: number;
  readonly port: number;
  readonly profileDir: string;
  readonly startedAt: string;
}

export interface SharedEdgeDeps {
  probe: (port: number) => Promise<boolean>;
  spawnEdge: (executablePath: string, args: readonly string[]) => number | undefined;
  resolveExecutable: (settings: SharedEdgeSettings) => string | undefined;
  killProcess: (pid: number) => Promise<void>;
  readProcessCommandLine: (pid: number) => Promise<string | undefined>;
  sleep: (ms: number) => Promise<void>;
}

/**
 * Owns the single Edge process that every MCP server process shares.
 *
 * Edge (Chromium) locks its profile directory to one browser process, so
 * letting each MCP server launch its own persistent context makes every
 * session after the first fail. Instead, the first server that needs the
 * browser starts one detached Edge with a loopback debugging port, and all
 * servers (including the starter) connect to it over CDP. The Edge process
 * outlives individual MCP servers, so starting or stopping an agent session
 * never closes the browser and never requires signing in again.
 */
export class SharedEdgeManager {
  private readonly stateFilePath: string;
  private readonly lockDirPath: string;

  constructor(
    private readonly settings: SharedEdgeSettings,
    private readonly logger: Logger,
    private readonly deps: SharedEdgeDeps = defaultDeps,
  ) {
    const stateDir = path.dirname(settings.profileDir);
    this.stateFilePath = path.join(stateDir, "shared-edge.json");
    this.lockDirPath = path.join(stateDir, "shared-edge-launch.lock");
  }

  get endpoint(): string {
    return `http://127.0.0.1:${this.settings.cdpPort}`;
  }

  async ensureRunning(): Promise<string> {
    if (await this.deps.probe(this.settings.cdpPort)) {
      return this.endpoint;
    }

    const deadline = Date.now() + LOCK_TIMEOUT_MS;
    while (Date.now() < deadline) {
      const acquired = await this.tryAcquireLaunchLock();
      if (!acquired) {
        if (await this.deps.probe(this.settings.cdpPort)) {
          return this.endpoint;
        }
        await this.deps.sleep(PROBE_INTERVAL_MS);
        continue;
      }

      try {
        if (await this.deps.probe(this.settings.cdpPort)) {
          return this.endpoint;
        }
        await this.startEdge();
        await this.waitUntilReachable();
        return this.endpoint;
      } finally {
        await this.releaseLaunchLock();
      }
    }

    throw new Error(
      "Timed out waiting for the shared Edge process to start. " +
        "Close any Edge window using the dedicated profile and try again.",
    );
  }

  /**
   * Stops the shared Edge recorded in the state file. Only a process whose
   * command line references this profile directory is terminated, so a
   * recycled PID belonging to another program is never killed.
   */
  async stop(): Promise<boolean> {
    const state = await this.readState();
    if (!state) {
      return false;
    }

    const commandLine = await this.deps.readProcessCommandLine(state.pid);
    if (commandLine === undefined) {
      await this.removeState();
      return false;
    }
    if (!commandLineReferencesProfile(commandLine, this.settings.profileDir)) {
      this.logger.warn("shared_edge_stop_skipped_pid_mismatch", {
        pid: state.pid,
      });
      await this.removeState();
      return false;
    }

    await this.deps.killProcess(state.pid);
    await this.removeState();
    this.logger.info("shared_edge_stopped", { pid: state.pid });
    return true;
  }

  async readState(): Promise<SharedEdgeState | undefined> {
    try {
      const raw = await readFile(this.stateFilePath, "utf8");
      const parsed = JSON.parse(raw) as Partial<SharedEdgeState>;
      if (
        typeof parsed.pid !== "number" ||
        typeof parsed.port !== "number" ||
        typeof parsed.profileDir !== "string" ||
        typeof parsed.startedAt !== "string"
      ) {
        return undefined;
      }
      return parsed as SharedEdgeState;
    } catch {
      return undefined;
    }
  }

  private async startEdge(): Promise<void> {
    const executablePath = this.deps.resolveExecutable(this.settings);
    if (!executablePath) {
      throw new Error(
        "Microsoft Edge executable was not found. " +
          "Set SHAREPOINT_EDGE_EXECUTABLE_PATH to the full path of msedge.exe, " +
          "or set SHAREPOINT_EDGE_MODE=exclusive to use the previous launch method.",
      );
    }

    await mkdir(this.settings.profileDir, { recursive: true });
    const args = buildSharedEdgeArgs(this.settings);
    this.logger.info("shared_edge_starting", {
      port: this.settings.cdpPort,
      headless: this.settings.headless,
    });
    const pid = this.deps.spawnEdge(executablePath, args);
    if (pid !== undefined) {
      await writeFile(
        this.stateFilePath,
        `${JSON.stringify(
          {
            pid,
            port: this.settings.cdpPort,
            profileDir: this.settings.profileDir,
            startedAt: new Date().toISOString(),
          } satisfies SharedEdgeState,
          null,
          2,
        )}\n`,
        "utf8",
      ).catch((error: unknown) => {
        this.logger.warn("shared_edge_state_write_failed", {
          errorType: getErrorType(error),
        });
      });
    }
  }

  private async waitUntilReachable(): Promise<void> {
    const deadline = Date.now() + START_TIMEOUT_MS;
    while (Date.now() < deadline) {
      if (await this.deps.probe(this.settings.cdpPort)) {
        return;
      }
      await this.deps.sleep(PROBE_INTERVAL_MS);
    }
    throw new Error(
      "The shared Edge process did not open its debugging port in time.",
    );
  }

  private async tryAcquireLaunchLock(): Promise<boolean> {
    try {
      await mkdir(this.lockDirPath);
      return true;
    } catch (error) {
      if (!isAlreadyExistsError(error)) {
        throw error;
      }
    }

    try {
      const lockStat = await stat(this.lockDirPath);
      if (Date.now() - lockStat.mtimeMs > LOCK_STALE_MS) {
        await rm(this.lockDirPath, { recursive: true, force: true });
        return this.tryAcquireLaunchLock();
      }
    } catch {
      // The lock was released between mkdir and stat; retry on the next pass.
    }
    return false;
  }

  private async releaseLaunchLock(): Promise<void> {
    await rm(this.lockDirPath, { recursive: true, force: true }).catch(
      () => undefined,
    );
  }

  private async removeState(): Promise<void> {
    await rm(this.stateFilePath, { force: true }).catch(() => undefined);
  }
}

export function buildSharedEdgeArgs(
  settings: SharedEdgeSettings,
): readonly string[] {
  return [
    `--user-data-dir=${settings.profileDir}`,
    "--remote-debugging-address=127.0.0.1",
    `--remote-debugging-port=${settings.cdpPort}`,
    "--no-first-run",
    "--no-default-browser-check",
    ...(settings.headless ? ["--headless=new"] : []),
    "about:blank",
  ];
}

export function commandLineReferencesProfile(
  commandLine: string,
  profileDir: string,
): boolean {
  // Profile dirs reach this function already resolved to absolute paths by
  // config loading. Normalize textually (not via path.resolve) so Windows
  // paths compare correctly regardless of the host running the comparison.
  const normalize = (value: string): string =>
    value.replaceAll("\\", "/").toLowerCase().replace(/\/+$/u, "");
  return normalize(commandLine).includes(normalize(profileDir));
}

async function probeCdpEndpoint(port: number): Promise<boolean> {
  try {
    const response = await fetch(`http://127.0.0.1:${port}/json/version`, {
      signal: AbortSignal.timeout(PROBE_TIMEOUT_MS),
    });
    if (!response.ok) {
      return false;
    }
    const payload = (await response.json()) as { webSocketDebuggerUrl?: unknown };
    return typeof payload.webSocketDebuggerUrl === "string";
  } catch {
    return false;
  }
}

function spawnDetachedEdge(
  executablePath: string,
  args: readonly string[],
): number | undefined {
  const child = spawn(executablePath, [...args], {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
  });
  child.unref();
  return child.pid;
}

const EDGE_APPLICATION_DIRS: Readonly<Record<BrowserChannel, string>> = {
  msedge: "Edge",
  "msedge-beta": "Edge Beta",
  "msedge-dev": "Edge Dev",
  "msedge-canary": "Edge SxS",
};

function resolveEdgeExecutable(
  settings: SharedEdgeSettings,
): string | undefined {
  if (settings.edgeExecutablePath) {
    return settings.edgeExecutablePath;
  }
  if (process.platform !== "win32") {
    return undefined;
  }

  const appDir = EDGE_APPLICATION_DIRS[settings.browserChannel];
  const roots = [
    process.env.ProgramFiles,
    process.env["ProgramFiles(x86)"],
    process.env.LOCALAPPDATA,
  ];
  for (const root of roots) {
    if (!root) {
      continue;
    }
    const candidate = path.join(
      root,
      "Microsoft",
      appDir,
      "Application",
      "msedge.exe",
    );
    if (existsSync(candidate)) {
      return candidate;
    }
  }
  return undefined;
}

async function killProcessTree(pid: number): Promise<void> {
  if (process.platform === "win32") {
    await execFileAsync("taskkill.exe", ["/PID", String(pid), "/T", "/F"]);
    return;
  }
  process.kill(pid, "SIGTERM");
}

async function readProcessCommandLine(
  pid: number,
): Promise<string | undefined> {
  if (process.platform === "win32") {
    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        [
          "-NoProfile",
          "-NonInteractive",
          "-Command",
          `(Get-CimInstance Win32_Process -Filter "ProcessId = ${pid}").CommandLine`,
        ],
        { timeout: 10_000 },
      );
      const commandLine = stdout.trim();
      return commandLine === "" ? undefined : commandLine;
    } catch {
      return undefined;
    }
  }

  try {
    const raw = await readFile(`/proc/${pid}/cmdline`, "utf8");
    const commandLine = raw.replaceAll("\0", " ").trim();
    return commandLine === "" ? undefined : commandLine;
  } catch {
    return undefined;
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

const defaultDeps: SharedEdgeDeps = {
  probe: probeCdpEndpoint,
  spawnEdge: spawnDetachedEdge,
  resolveExecutable: resolveEdgeExecutable,
  killProcess: killProcessTree,
  readProcessCommandLine,
  sleep,
};

function isAlreadyExistsError(error: unknown): boolean {
  return (
    error instanceof Error &&
    "code" in error &&
    (error as NodeJS.ErrnoException).code === "EEXIST"
  );
}

function getErrorType(error: unknown): string {
  return error instanceof Error ? error.name : typeof error;
}
