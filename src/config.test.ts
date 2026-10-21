import { join, resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { DEFAULT_SCOPES, dotenvPath, loadConfig, MISSING_CLIENT_ID_MESSAGE } from "./config.js";
import { ConfigError } from "./util/errors.js";

const base = { MICROSOFT_CLIENT_ID: "11111111-2222-3333-4444-555555555555" };

describe("dotenvPath (#41)", () => {
  it("is the connector checkout root, derived from the entry script, not the CWD", () => {
    const repo = resolve("/work/Claude-Hotmail-Connector");
    expect(dotenvPath(join(repo, "dist", "index.js"))).toBe(join(repo, ".env"));
    // An extension bundle: <ext>/server/index.cjs -> <ext>/.env (none shipped).
    const ext = resolve("/ext/hotmail");
    expect(dotenvPath(join(ext, "server", "index.cjs"))).toBe(join(ext, ".env"));
    expect(dotenvPath(join(repo, "dist", "index.js"))).not.toBe(join(process.cwd(), ".env"));
  });

  it("is undefined without an entry script", () => {
    expect(dotenvPath(undefined)).toBeUndefined();
  });
});

describe("loadConfig", () => {
  it("requires a client id: there is no bundled shared default", () => {
    expect(() => loadConfig({})).toThrow(ConfigError);
    expect(() => loadConfig({ MICROSOFT_CLIENT_ID: "   " })).toThrow(MISSING_CLIENT_ID_MESSAGE);
    expect(MISSING_CLIENT_ID_MESSAGE).toContain("Settings > Extensions");
  });

  it("uses MICROSOFT_CLIENT_ID (trimmed)", () => {
    expect(
      loadConfig({ MICROSOFT_CLIENT_ID: ` ${base.MICROSOFT_CLIENT_ID} ` }).microsoftClientId,
    ).toBe(base.MICROSOFT_CLIENT_ID);
  });

  it("applies the documented defaults", () => {
    const cfg = loadConfig(base);
    expect(cfg.microsoftClientId).toBe(base.MICROSOFT_CLIENT_ID);
    expect(cfg.microsoftTenant).toBe("consumers");
    expect(cfg.authority).toBe("https://login.microsoftonline.com/consumers");
    expect(cfg.defaultTimeZone).toBe("Australia/Brisbane");
    expect(cfg.graphBaseUrl).toBe("https://graph.microsoft.com/v1.0");
    expect(cfg.loopbackHost).toBe("127.0.0.1");
    expect(cfg.scopes).toEqual([...DEFAULT_SCOPES]);
    expect(cfg.logLevel).toBe("info");
    expect(cfg.tokenCacheFile).toMatch(/token-cache\.enc$/);
  });

  it("parses space- or comma-separated scope overrides, always adding the sign-in scopes", () => {
    expect(loadConfig({ ...base, MICROSOFT_SCOPES: "openid Mail.Read" }).scopes).toEqual([
      "openid",
      "offline_access",
      "Mail.Read",
    ]);
    expect(loadConfig({ ...base, MICROSOFT_SCOPES: "Mail.Read , User.Read" }).scopes).toEqual([
      "openid",
      "offline_access",
      "User.Read",
      "Mail.Read",
    ]);
  });

  it("allows narrowing Calendars.ReadWrite to Calendars.Read, case-insensitively", () => {
    expect(loadConfig({ ...base, MICROSOFT_SCOPES: "mail.read calendars.read" }).scopes).toEqual([
      "openid",
      "offline_access",
      "Mail.Read",
      "Calendars.Read",
    ]);
  });

  it("rejects scope overrides that would broaden access", () => {
    for (const broad of ["Mail.ReadWrite", "Mail.Send", "Files.Read", "Calendars.Read.Shared"]) {
      expect(() => loadConfig({ ...base, MICROSOFT_SCOPES: `openid ${broad}` })).toThrow(
        new RegExp(`not allowed: ${broad.replace(/\./g, "\\.")}`),
      );
    }
  });

  it("treats an empty override as the default set", () => {
    expect(loadConfig({ ...base, MICROSOFT_SCOPES: " , " }).scopes).toEqual([...DEFAULT_SCOPES]);
  });

  it("honours a custom tenant and timezone", () => {
    const cfg = loadConfig({
      ...base,
      MICROSOFT_TENANT: "common",
      DEFAULT_TIMEZONE: "Europe/London",
    });
    expect(cfg.microsoftTenant).toBe("common");
    expect(cfg.authority).toBe("https://login.microsoftonline.com/common");
    expect(cfg.defaultTimeZone).toBe("Europe/London");
  });

  it("honours an explicit cache dir override", () => {
    const cfg = loadConfig({ ...base, CLAUDE_HOTMAIL_CACHE_DIR: "/tmp/cache-x" });
    expect(cfg.cacheDir).toBe("/tmp/cache-x");
  });

  it("rejects an invalid LOG_LEVEL", () => {
    expect(() => loadConfig({ ...base, LOG_LEVEL: "verbose" })).toThrow(ConfigError);
  });
});
