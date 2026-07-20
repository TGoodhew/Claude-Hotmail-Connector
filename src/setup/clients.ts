/**
 * Auto-configure MCP host apps (Claude Desktop, Claude Code) to launch this
 * connector — so a non-technical user never hand-edits a JSON config.
 *
 * This encodes the platform quirks learned the hard way while wiring the
 * connector up by hand:
 *
 *  - **Claude Desktop on Windows has two possible config locations.** The
 *    standalone build uses `%APPDATA%\Claude\claude_desktop_config.json`; the
 *    Microsoft Store (MSIX) build virtualises `%APPDATA%`, so its config lives
 *    under `%LOCALAPPDATA%\Packages\<pkgFamily>\LocalCache\Roaming\Claude\`.
 *  - **Merge, never overwrite.** The Desktop config also holds unrelated user
 *    preferences; we only add/replace our own `mcpServers` entry.
 *  - **Write via JSON.stringify** so Windows backslashes are escaped correctly
 *    (a hand-written single-backslash path is invalid JSON and is silently
 *    skipped by the host).
 *  - **Launch with an absolute executable path** — `node` is frequently not on
 *    the host process's PATH.
 *
 * The pure functions (`buildServerEntry`, `mergeMcpServer`) and the fs-driven
 * ones (via the injectable {@link FsLike}) are unit-tested without touching the
 * real filesystem.
 */

import { join } from "node:path";

export const SERVER_KEY = "hotmail-connector";

/** A stdio MCP server entry as written into a host config's `mcpServers` map. */
export interface ServerEntry {
  type: "stdio";
  command: string;
  args: string[];
  env?: Record<string, string>;
}

/** Minimal filesystem surface, injectable for testing. */
export interface FsLike {
  existsSync(path: string): boolean;
  readFileSync(path: string, encoding: "utf8"): string;
  writeFileSync(path: string, data: string): void;
  readdirSync(path: string): string[];
  mkdirSync(path: string, options: { recursive: true }): void;
}

export interface BuildEntryOptions {
  /** Absolute path to the launcher executable (node.exe, or a bundled exe). */
  command: string;
  /** Args to pass (e.g. the built `dist/index.js` for a node launcher; []). */
  args?: string[];
  /** Client id to pass via env; omit to rely on the bundled default. */
  clientId?: string;
}

/** Build the stdio server entry the host will use to launch the connector. */
export function buildServerEntry(opts: BuildEntryOptions): ServerEntry {
  const entry: ServerEntry = {
    type: "stdio",
    command: opts.command,
    args: opts.args ?? [],
  };
  if (opts.clientId) entry.env = { MICROSOFT_CLIENT_ID: opts.clientId };
  return entry;
}

/**
 * Merge a server entry into a host config object without disturbing anything
 * else. Returns a NEW object (does not mutate the input). Preserves every
 * existing top-level key and every other server under `mcpServers`.
 */
export function mergeMcpServer(
  config: unknown,
  name: string,
  entry: ServerEntry,
): Record<string, unknown> {
  const base: Record<string, unknown> =
    config && typeof config === "object" && !Array.isArray(config)
      ? { ...(config as Record<string, unknown>) }
      : {};
  const existingServers = base.mcpServers;
  const servers: Record<string, unknown> =
    existingServers && typeof existingServers === "object" && !Array.isArray(existingServers)
      ? { ...(existingServers as Record<string, unknown>) }
      : {};
  servers[name] = entry;
  base.mcpServers = servers;
  return base;
}

/**
 * Remove a named server from a host config, leaving everything else intact.
 * Returns a NEW object and whether anything was actually removed.
 */
export function removeMcpServer(
  config: unknown,
  name: string,
): { config: Record<string, unknown>; removed: boolean } {
  const base: Record<string, unknown> =
    config && typeof config === "object" && !Array.isArray(config)
      ? { ...(config as Record<string, unknown>) }
      : {};
  const existing = base.mcpServers;
  if (!existing || typeof existing !== "object" || Array.isArray(existing)) {
    return { config: base, removed: false };
  }
  const servers = { ...(existing as Record<string, unknown>) };
  const removed = name in servers;
  delete servers[name];
  base.mcpServers = servers;
  return { config: base, removed };
}

/** Where a host config lives and which app it belongs to. */
export interface ConfigTarget {
  client: "claude-desktop" | "claude-code";
  /** Absolute path to the config file (may not exist yet). */
  path: string;
  /** Whether the file exists right now. */
  exists: boolean;
}

/**
 * Discover Claude Desktop config files on Windows: the standalone location and
 * any Microsoft Store (MSIX) package location. Only returns paths whose file
 * currently exists (we don't create a Desktop config for an app that isn't
 * installed).
 */
export function discoverDesktopConfigsWindows(env: NodeJS.ProcessEnv, fs: FsLike): ConfigTarget[] {
  const targets: ConfigTarget[] = [];
  const seen = new Set<string>();
  const add = (path: string): void => {
    if (seen.has(path)) return;
    seen.add(path);
    if (fs.existsSync(path)) targets.push({ client: "claude-desktop", path, exists: true });
  };

  // Standalone build.
  if (env.APPDATA) add(join(env.APPDATA, "Claude", "claude_desktop_config.json"));

  // MSIX / Microsoft Store build(s): %LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\
  if (env.LOCALAPPDATA) {
    const packagesDir = join(env.LOCALAPPDATA, "Packages");
    if (fs.existsSync(packagesDir)) {
      for (const pkg of safeReaddir(fs, packagesDir)) {
        if (!pkg.startsWith("Claude")) continue;
        add(join(packagesDir, pkg, "LocalCache", "Roaming", "Claude", "claude_desktop_config.json"));
      }
    }
  }
  return targets;
}

function safeReaddir(fs: FsLike, dir: string): string[] {
  try {
    return fs.readdirSync(dir);
  } catch {
    return [];
  }
}

export interface ApplyResult {
  path: string;
  outcome: "created" | "updated";
  backupPath?: string;
}

/**
 * Apply a server entry to a single config file: read (or start empty), merge,
 * back up the pristine original once, and write valid JSON.
 */
export function applyToConfigFile(
  filePath: string,
  name: string,
  entry: ServerEntry,
  fs: FsLike,
): ApplyResult {
  const existed = fs.existsSync(filePath);
  let current: unknown = {};
  if (existed) {
    const raw = fs.readFileSync(filePath, "utf8");
    if (raw.trim().length > 0) {
      current = JSON.parse(raw) as unknown; // throws on invalid JSON — surfaced to caller
    }
  }

  let backupPath: string | undefined;
  if (existed) {
    // Keep the very first backup pristine; never clobber an earlier one.
    backupPath = `${filePath}.backup`;
    if (!fs.existsSync(backupPath)) {
      fs.writeFileSync(backupPath, fs.readFileSync(filePath, "utf8"));
    }
  } else {
    ensureDir(filePath, fs);
  }

  const merged = mergeMcpServer(current, name, entry);
  fs.writeFileSync(filePath, `${JSON.stringify(merged, null, 2)}\n`);
  return { path: filePath, outcome: existed ? "updated" : "created", backupPath };
}

function ensureDir(filePath: string, fs: FsLike): void {
  const dir = filePath.replace(/[\\/][^\\/]*$/, "");
  if (dir && dir !== filePath && !fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

/** Remove our server entry from a single config file (no-op if absent/missing). */
export function removeFromConfigFile(
  filePath: string,
  name: string,
  fs: FsLike,
): { path: string; removed: boolean } {
  if (!fs.existsSync(filePath)) return { path: filePath, removed: false };
  const raw = fs.readFileSync(filePath, "utf8");
  if (raw.trim().length === 0) return { path: filePath, removed: false };
  let current: unknown;
  try {
    current = JSON.parse(raw) as unknown;
  } catch {
    // Don't touch a config we can't parse.
    return { path: filePath, removed: false };
  }
  const { config, removed } = removeMcpServer(current, name);
  if (removed) fs.writeFileSync(filePath, `${JSON.stringify(config, null, 2)}\n`);
  return { path: filePath, removed };
}
