import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdtemp, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import {
  buildSharedEdgeArgs,
  commandLineReferencesProfile,
  SharedEdgeManager,
  type SharedEdgeDeps,
  type SharedEdgeSettings,
} from "../src/browser/shared-edge.js";
import type { Logger } from "../src/logger.js";

const LOGGER: Logger = {
  info: () => undefined,
  warn: () => undefined,
  error: () => undefined,
};

interface FakeState {
  running: boolean;
  spawnCalls: number;
  spawnedArgs: readonly string[];
  killedPids: number[];
  commandLine: string | undefined;
}

function createSettings(profileDir: string): SharedEdgeSettings {
  return {
    profileDir,
    cdpPort: 9229,
    headless: true,
    browserChannel: "msedge",
    edgeExecutablePath: "C:\\fake\\msedge.exe",
  };
}

function createFakeDeps(state: FakeState): SharedEdgeDeps {
  return {
    probe: async () => state.running,
    spawnEdge: (_executablePath, args) => {
      state.spawnCalls += 1;
      state.spawnedArgs = args;
      state.running = true;
      return 4321;
    },
    resolveExecutable: (settings) => settings.edgeExecutablePath,
    killProcess: async (pid) => {
      state.killedPids.push(pid);
      state.running = false;
    },
    readProcessCommandLine: async () => state.commandLine,
    sleep: async () => undefined,
  };
}

async function createTempProfileDir(): Promise<string> {
  const base = await mkdtemp(path.join(os.tmpdir(), "shared-edge-test-"));
  return path.join(base, "edge-profile");
}

test("reuses a running shared Edge without spawning", async () => {
  const state: FakeState = {
    running: true,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: undefined,
  };
  const manager = new SharedEdgeManager(
    createSettings(await createTempProfileDir()),
    LOGGER,
    createFakeDeps(state),
  );

  const endpoint = await manager.ensureRunning();

  assert.equal(endpoint, "http://127.0.0.1:9229");
  assert.equal(state.spawnCalls, 0);
});

test("starts Edge once with the profile and debugging port when not running", async () => {
  const profileDir = await createTempProfileDir();
  const state: FakeState = {
    running: false,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: undefined,
  };
  const manager = new SharedEdgeManager(
    createSettings(profileDir),
    LOGGER,
    createFakeDeps(state),
  );

  const endpoint = await manager.ensureRunning();

  assert.equal(endpoint, "http://127.0.0.1:9229");
  assert.equal(state.spawnCalls, 1);
  assert.ok(
    state.spawnedArgs.includes(`--user-data-dir=${profileDir}`),
    "spawn args must pin the dedicated profile directory",
  );
  assert.ok(state.spawnedArgs.includes("--remote-debugging-port=9229"));
  assert.ok(state.spawnedArgs.includes("--headless=new"));

  const recorded = await manager.readState();
  assert.equal(recorded?.pid, 4321);
  assert.equal(recorded?.port, 9229);
  assert.equal(recorded?.profileDir, profileDir);
});

test("concurrent starters spawn Edge only once", async () => {
  const profileDir = await createTempProfileDir();
  const settings = createSettings(profileDir);
  const state: FakeState = {
    running: false,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: undefined,
  };
  const deps = createFakeDeps(state);
  const first = new SharedEdgeManager(settings, LOGGER, deps);
  const second = new SharedEdgeManager(settings, LOGGER, deps);

  const [firstEndpoint, secondEndpoint] = await Promise.all([
    first.ensureRunning(),
    second.ensureRunning(),
  ]);

  assert.equal(firstEndpoint, "http://127.0.0.1:9229");
  assert.equal(secondEndpoint, "http://127.0.0.1:9229");
  assert.equal(state.spawnCalls, 1);
});

test("fails clearly when the Edge executable cannot be resolved", async () => {
  const state: FakeState = {
    running: false,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: undefined,
  };
  const deps = createFakeDeps(state);
  deps.resolveExecutable = () => undefined;
  const manager = new SharedEdgeManager(
    createSettings(await createTempProfileDir()),
    LOGGER,
    deps,
  );

  await assert.rejects(
    () => manager.ensureRunning(),
    /SHAREPOINT_EDGE_EXECUTABLE_PATH/u,
  );
  assert.equal(state.spawnCalls, 0);
});

test("stops only the recorded Edge that uses this profile", async () => {
  const profileDir = await createTempProfileDir();
  const stateFile = path.join(path.dirname(profileDir), "shared-edge.json");
  await writeFile(
    stateFile,
    JSON.stringify({
      pid: 4321,
      port: 9229,
      profileDir,
      startedAt: "2026-11-11T00:00:00.000Z",
    }),
    "utf8",
  );
  const state: FakeState = {
    running: true,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: `"C:\\fake\\msedge.exe" --user-data-dir=${profileDir}`,
  };
  const manager = new SharedEdgeManager(
    createSettings(profileDir),
    LOGGER,
    createFakeDeps(state),
  );

  const stopped = await manager.stop();

  assert.equal(stopped, true);
  assert.deepEqual(state.killedPids, [4321]);
  assert.equal(existsSync(stateFile), false);
});

test("refuses to stop a recycled PID that belongs to another process", async () => {
  const profileDir = await createTempProfileDir();
  const stateFile = path.join(path.dirname(profileDir), "shared-edge.json");
  await writeFile(
    stateFile,
    JSON.stringify({
      pid: 4321,
      port: 9229,
      profileDir,
      startedAt: "2026-11-11T00:00:00.000Z",
    }),
    "utf8",
  );
  const state: FakeState = {
    running: true,
    spawnCalls: 0,
    spawnedArgs: [],
    killedPids: [],
    commandLine: `"C:\\other\\app.exe" --user-data-dir=C:\\other\\profile`,
  };
  const manager = new SharedEdgeManager(
    createSettings(profileDir),
    LOGGER,
    createFakeDeps(state),
  );

  const stopped = await manager.stop();

  assert.equal(stopped, false);
  assert.deepEqual(state.killedPids, []);
});

test("builds headed args without the headless flag", () => {
  const args = buildSharedEdgeArgs({
    ...createSettings("/tmp/edge-profile"),
    headless: false,
  });

  assert.equal(args.includes("--headless=new"), false);
  assert.ok(args.includes("--remote-debugging-address=127.0.0.1"));
});

test("matches profile paths regardless of case and slash direction", () => {
  assert.equal(
    commandLineReferencesProfile(
      "msedge.exe --user-data-dir=C:/Apps/SharePointBrowserMcp/edge-profile",
      "C:\\Apps\\SharePointBrowserMcp\\Edge-Profile",
    ),
    true,
  );
  assert.equal(
    commandLineReferencesProfile(
      "msedge.exe --user-data-dir=C:\\Apps\\other\\edge-profile",
      "C:\\Apps\\SharePointBrowserMcp\\edge-profile",
    ),
    false,
  );
});
