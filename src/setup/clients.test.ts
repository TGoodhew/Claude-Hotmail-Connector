import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  applyToConfigFile,
  buildServerEntry,
  discoverDesktopConfigsWindows,
  mergeMcpServer,
  removeFromConfigFile,
  removeMcpServer,
  SERVER_KEY,
  type FsLike,
} from "./clients.js";

function memFs(files: Record<string, string> = {}, dirs: Record<string, string[]> = {}) {
  const store: Record<string, string> = { ...files };
  const dirStore: Record<string, string[]> = { ...dirs };
  const fs: FsLike = {
    existsSync: (p) => p in store || p in dirStore,
    readFileSync: (p) => {
      if (!(p in store)) throw new Error(`ENOENT: ${p}`);
      return store[p]!;
    },
    writeFileSync: (p, data) => {
      store[p] = data;
    },
    readdirSync: (p) => {
      if (!(p in dirStore)) throw new Error(`ENOENT: ${p}`);
      return dirStore[p]!;
    },
    mkdirSync: () => {},
  };
  return { fs, store, dirStore };
}

describe("buildServerEntry", () => {
  it("builds a stdio entry with the client id in env", () => {
    const entry = buildServerEntry({
      command: "C:\\node.exe",
      args: ["C:\\index.js"],
      clientId: "cid",
    });
    expect(entry).toEqual({
      type: "stdio",
      command: "C:\\node.exe",
      args: ["C:\\index.js"],
      env: { MICROSOFT_CLIENT_ID: "cid" },
    });
  });
  it("omits env when no client id is given (uses the bundled default)", () => {
    const entry = buildServerEntry({ command: "app.exe" });
    expect(entry).toEqual({ type: "stdio", command: "app.exe", args: [] });
  });
});

describe("mergeMcpServer", () => {
  const entry = buildServerEntry({ command: "n", args: ["i"], clientId: "c" });

  it("preserves unrelated top-level keys and other servers", () => {
    const existing = {
      coworkUserFilesPath: "C:\\Users\\Tony\\Claude",
      preferences: { sidebarMode: "chat" },
      mcpServers: { other: { command: "keep-me" } },
    };
    const merged = mergeMcpServer(existing, SERVER_KEY, entry);
    expect(merged.coworkUserFilesPath).toBe("C:\\Users\\Tony\\Claude");
    expect(merged.preferences).toEqual({ sidebarMode: "chat" });
    expect((merged.mcpServers as Record<string, unknown>).other).toEqual({ command: "keep-me" });
    expect((merged.mcpServers as Record<string, unknown>)[SERVER_KEY]).toEqual(entry);
  });

  it("does not mutate the input", () => {
    const existing = { mcpServers: { other: { command: "x" } } };
    const snapshot = JSON.stringify(existing);
    mergeMcpServer(existing, SERVER_KEY, entry);
    expect(JSON.stringify(existing)).toBe(snapshot);
  });

  it("starts fresh for a null / non-object / array config", () => {
    for (const bad of [null, undefined, 42, "str", [1, 2]]) {
      const merged = mergeMcpServer(bad, SERVER_KEY, entry);
      expect((merged.mcpServers as Record<string, unknown>)[SERVER_KEY]).toEqual(entry);
    }
  });
});

describe("discoverDesktopConfigsWindows", () => {
  const APPDATA = "C:\\A\\Roaming";
  const LOCALAPPDATA = "C:\\A\\Local";
  const standalone = join(APPDATA, "Claude", "claude_desktop_config.json");
  const packagesDir = join(LOCALAPPDATA, "Packages");
  const msix = join(
    packagesDir,
    "Claude_pzs8sxrjxfjjc",
    "LocalCache",
    "Roaming",
    "Claude",
    "claude_desktop_config.json",
  );

  it("finds both the standalone and MSIX configs that exist, skips the rest", () => {
    const { fs } = memFs(
      { [standalone]: "{}", [msix]: "{}" },
      { [packagesDir]: ["Claude_pzs8sxrjxfjjc", "SomethingElse_x", "Claude_uninstalled"] },
    );
    const found = discoverDesktopConfigsWindows({ APPDATA, LOCALAPPDATA } as NodeJS.ProcessEnv, fs);
    const paths = found.map((t) => t.path);
    expect(paths).toContain(standalone);
    expect(paths).toContain(msix);
    // "Claude_uninstalled" has no config file, "SomethingElse_x" isn't Claude.
    expect(paths).toHaveLength(2);
    expect(found.every((t) => t.client === "claude-desktop" && t.exists)).toBe(true);
  });

  it("returns nothing when no configs exist", () => {
    const { fs } = memFs({}, { [packagesDir]: ["Claude_pzs8sxrjxfjjc"] });
    expect(
      discoverDesktopConfigsWindows({ APPDATA, LOCALAPPDATA } as NodeJS.ProcessEnv, fs),
    ).toEqual([]);
  });
});

describe("applyToConfigFile", () => {
  const path = "C:\\cfg\\claude_desktop_config.json";
  const entry = buildServerEntry({
    command: "C:\\node.exe",
    args: ["C:\\dist\\index.js"],
    clientId: "cid",
  });

  it("merges into an existing config, backs it up, and writes valid JSON", () => {
    const original = JSON.stringify({ preferences: { a: 1 } });
    const { fs, store } = memFs({ [path]: original });
    const res = applyToConfigFile(path, SERVER_KEY, entry, fs);

    expect(res.outcome).toBe("updated");
    expect(res.backupPath).toBe(`${path}.backup`);
    expect(store[`${path}.backup`]).toBe(original); // pristine backup
    const written = JSON.parse(store[path]!) as Record<string, unknown>;
    expect(written.preferences).toEqual({ a: 1 });
    expect((written.mcpServers as Record<string, unknown>)[SERVER_KEY]).toEqual(entry);
  });

  it("creates a new config when none exists (no backup)", () => {
    const { fs, store } = memFs({});
    const res = applyToConfigFile(path, SERVER_KEY, entry, fs);
    expect(res.outcome).toBe("created");
    expect(res.backupPath).toBeUndefined();
    expect(
      (JSON.parse(store[path]!) as { mcpServers: Record<string, unknown> }).mcpServers[SERVER_KEY],
    ).toEqual(entry);
  });

  it("keeps the first backup pristine across repeated applies", () => {
    const original = JSON.stringify({ preferences: { a: 1 } });
    const { fs, store } = memFs({ [path]: original });
    applyToConfigFile(path, SERVER_KEY, entry, fs);
    // Second run with a different entry must not overwrite the pristine backup.
    const entry2 = buildServerEntry({ command: "other.exe" });
    applyToConfigFile(path, SERVER_KEY, entry2, fs);
    expect(store[`${path}.backup`]).toBe(original);
    expect(
      (JSON.parse(store[path]!) as { mcpServers: Record<string, unknown> }).mcpServers[SERVER_KEY],
    ).toEqual(entry2);
  });

  it("throws on an existing invalid-JSON config (rather than silently clobbering)", () => {
    const { fs } = memFs({ [path]: "{ not valid json" });
    expect(() => applyToConfigFile(path, SERVER_KEY, entry, fs)).toThrow();
  });
});

describe("removeMcpServer / removeFromConfigFile", () => {
  const path = "C:\\cfg\\claude_desktop_config.json";
  const entry = buildServerEntry({ command: "n", args: ["i"], clientId: "c" });

  it("removes only our entry and preserves the rest", () => {
    const cfg = {
      preferences: { a: 1 },
      mcpServers: { [SERVER_KEY]: entry, other: { command: "keep" } },
    };
    const { config, removed } = removeMcpServer(cfg, SERVER_KEY);
    expect(removed).toBe(true);
    expect((config.mcpServers as Record<string, unknown>)[SERVER_KEY]).toBeUndefined();
    expect((config.mcpServers as Record<string, unknown>).other).toEqual({ command: "keep" });
    expect(config.preferences).toEqual({ a: 1 });
  });

  it("reports removed=false when the entry is absent", () => {
    expect(removeMcpServer({ mcpServers: {} }, SERVER_KEY).removed).toBe(false);
    expect(removeMcpServer({}, SERVER_KEY).removed).toBe(false);
  });

  it("rewrites the file only when something was removed, preserving other servers", () => {
    const original = JSON.stringify({
      mcpServers: { [SERVER_KEY]: entry, other: { command: "keep" } },
    });
    const { fs, store } = memFs({ [path]: original });
    const res = removeFromConfigFile(path, SERVER_KEY, fs);
    expect(res.removed).toBe(true);
    const written = JSON.parse(store[path]!) as { mcpServers: Record<string, unknown> };
    expect(written.mcpServers[SERVER_KEY]).toBeUndefined();
    expect(written.mcpServers.other).toEqual({ command: "keep" });
  });

  it("is a no-op for a missing file or unparseable config", () => {
    const { fs: fs1 } = memFs({});
    expect(removeFromConfigFile(path, SERVER_KEY, fs1).removed).toBe(false);
    const { fs: fs2, store } = memFs({ [path]: "{ broken" });
    expect(removeFromConfigFile(path, SERVER_KEY, fs2).removed).toBe(false);
    expect(store[path]).toBe("{ broken"); // untouched
  });
});
