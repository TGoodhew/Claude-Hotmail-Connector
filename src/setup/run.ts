/**
 * `setup` command: point the user's MCP host apps at this connector with no
 * hand-editing. v1 targets Windows Claude Desktop (standalone + Microsoft Store
 * builds); other clients/platforms are documented for manual setup.
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { createLogger, type Logger } from "../util/log.js";
import {
  applyToConfigFile,
  buildServerEntry,
  discoverDesktopConfigsWindows,
  removeFromConfigFile,
  SERVER_KEY,
  type FsLike,
} from "./clients.js";

const realFs: FsLike = {
  existsSync,
  readFileSync: (p) => readFileSync(p, "utf8"),
  writeFileSync: (p, d) => writeFileSync(p, d),
  readdirSync: (p) => readdirSync(p),
  mkdirSync: (p, o) => {
    mkdirSync(p, o);
  },
};

export interface SetupDeps {
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  fs?: FsLike;
  log?: Logger;
  /** Launcher executable (defaults to the current process's executable). */
  command?: string;
  /** Args used to launch the server (defaults derived from how we started). */
  args?: string[];
  /** Client id to embed; defaults to the env override, else the bundled default. */
  clientId?: string;
}

export interface SetupOutcome {
  configured: { path: string; outcome: "created" | "updated" }[];
  message: string;
}

/** Detect and configure host apps. Pure-ish: all effects go through injected deps. */
export function runSetup(deps: SetupDeps = {}): SetupOutcome {
  const env = deps.env ?? process.env;
  const platform = deps.platform ?? process.platform;
  const fs = deps.fs ?? realFs;
  const log = deps.log ?? createLogger({ name: "setup" });

  const command = deps.command ?? process.execPath;
  const args = deps.args ?? defaultArgs();
  const clientId = deps.clientId ?? envClientId(env);
  const entry = buildServerEntry({ command, args, clientId });

  const targets = platform === "win32" ? discoverDesktopConfigsWindows(env, fs) : [];

  const configured: SetupOutcome["configured"] = [];
  for (const t of targets) {
    try {
      const res = applyToConfigFile(t.path, SERVER_KEY, entry, fs);
      configured.push({ path: res.path, outcome: res.outcome });
      log.info(`Configured Claude Desktop (${res.outcome})`, {
        path: res.path,
        backup: res.backupPath,
      });
    } catch (e) {
      log.error("Failed to update a host config", { path: t.path, error: (e as Error).message });
    }
  }

  return { configured, message: buildMessage(platform, configured.length) };
}

/** Reverse of {@link runSetup}: remove our entry from detected host configs (used on uninstall). */
export function runUnsetup(deps: Pick<SetupDeps, "env" | "platform" | "fs" | "log"> = {}): SetupOutcome {
  const env = deps.env ?? process.env;
  const platform = deps.platform ?? process.platform;
  const fs = deps.fs ?? realFs;
  const log = deps.log ?? createLogger({ name: "setup" });

  const targets = platform === "win32" ? discoverDesktopConfigsWindows(env, fs) : [];
  const configured: SetupOutcome["configured"] = [];
  for (const t of targets) {
    try {
      const res = removeFromConfigFile(t.path, SERVER_KEY, fs);
      if (res.removed) {
        configured.push({ path: res.path, outcome: "updated" });
        log.info("Removed connector from Claude Desktop config", { path: res.path });
      }
    } catch (e) {
      log.error("Failed to update a host config", { path: t.path, error: (e as Error).message });
    }
  }
  return {
    configured,
    message: configured.length
      ? `Removed the connector from ${configured.length} Claude Desktop config file(s).`
      : "Nothing to remove.",
  };
}

function envClientId(env: NodeJS.ProcessEnv): string | undefined {
  const v = env.MICROSOFT_CLIENT_ID?.trim();
  return v && v.length > 0 ? v : undefined;
}

function defaultArgs(): string[] {
  const script = process.argv[1];
  // `node dist/index.js`: relaunch node with the script (server is the default
  // command). Packaged single-exe: no script arg — the exe defaults to server.
  if (script && /\.[cm]?js$/i.test(script)) return [script];
  return [];
}

function buildMessage(platform: NodeJS.Platform, count: number): string {
  if (count > 0) {
    return (
      `Configured ${count} Claude Desktop config file(s).\n` +
      "Next: FULLY quit Claude Desktop from the system tray (closing the window only minimises it), " +
      "reopen it, then run `login` once if you haven't signed in yet."
    );
  }
  if (platform === "win32") {
    return (
      "No Claude Desktop config was found. Open Claude Desktop once so it creates its config, " +
      "then run `setup` again. (Claude Code is configured via a project .mcp.json — see the README.)"
    );
  }
  return "Automatic setup currently supports Windows Claude Desktop only; see the README for manual setup.";
}
