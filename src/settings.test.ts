import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { isValidClientId, readSettings, settingsFilePath, writeSettings } from "./settings.js";
import { ConfigError } from "./util/errors.js";

const ID = "f0e1d2c3-b4a5-4697-8879-6a5b4c3d2e1f";

let dir: string;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "cht-settings-"));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("isValidClientId", () => {
  it("accepts a GUID (any case, surrounding whitespace)", () => {
    expect(isValidClientId(ID)).toBe(true);
    expect(isValidClientId(` ${ID.toUpperCase()} `)).toBe(true);
  });

  it("rejects non-GUIDs", () => {
    for (const bad of ["", "abc", ID.slice(1), `${ID}0`, ID.replace(/-/g, "")]) {
      expect(isValidClientId(bad)).toBe(false);
    }
  });
});

describe("readSettings / writeSettings", () => {
  it("returns {} when no settings file exists", () => {
    expect(readSettings(dir)).toEqual({});
  });

  it("round-trips the client id, creating the directory if needed", () => {
    const nested = join(dir, "a", "b");
    const path = writeSettings(nested, { clientId: ID });
    expect(path).toBe(settingsFilePath(nested));
    expect(readSettings(nested)).toEqual({ clientId: ID });
    expect(JSON.parse(readFileSync(path, "utf8"))).toEqual({ clientId: ID });
  });

  it("merges into existing settings rather than replacing them", () => {
    writeFileSync(settingsFilePath(dir), JSON.stringify({ clientId: "old", extra: 1 }));
    writeSettings(dir, { clientId: ID });
    expect(JSON.parse(readFileSync(settingsFilePath(dir), "utf8"))).toEqual({
      clientId: ID,
      extra: 1,
    });
  });

  it("tolerates a UTF-8 BOM (e.g. after editing in Notepad)", () => {
    writeFileSync(settingsFilePath(dir), "\uFEFF" + JSON.stringify({ clientId: ID }));
    expect(readSettings(dir)).toEqual({ clientId: ID });
  });

  it("throws a ConfigError for a corrupt or invalid file", () => {
    writeFileSync(settingsFilePath(dir), "{ not json");
    expect(() => readSettings(dir)).toThrow(ConfigError);
    writeFileSync(settingsFilePath(dir), JSON.stringify({ clientId: 42 }));
    expect(() => readSettings(dir)).toThrow(ConfigError);
  });
});
