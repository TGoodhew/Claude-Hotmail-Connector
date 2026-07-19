import { describe, expect, it } from "vitest";
import { DEFAULT_SCOPES, loadConfig } from "./config.js";
import { ConfigError } from "./util/errors.js";

const base = { MICROSOFT_CLIENT_ID: "11111111-2222-3333-4444-555555555555" };

describe("loadConfig", () => {
  it("throws a ConfigError when MICROSOFT_CLIENT_ID is missing", () => {
    expect(() => loadConfig({})).toThrow(ConfigError);
    expect(() => loadConfig({})).toThrow(/MICROSOFT_CLIENT_ID/);
  });

  it("throws when MICROSOFT_CLIENT_ID is blank", () => {
    expect(() => loadConfig({ MICROSOFT_CLIENT_ID: "   " })).toThrow(ConfigError);
  });

  it("applies the documented defaults", () => {
    const cfg = loadConfig(base);
    expect(cfg.microsoftClientId).toBe(base.MICROSOFT_CLIENT_ID);
    expect(cfg.microsoftTenant).toBe("common");
    expect(cfg.authority).toBe("https://login.microsoftonline.com/common");
    expect(cfg.defaultTimeZone).toBe("Australia/Brisbane");
    expect(cfg.graphBaseUrl).toBe("https://graph.microsoft.com/v1.0");
    expect(cfg.loopbackHost).toBe("127.0.0.1");
    expect(cfg.scopes).toEqual([...DEFAULT_SCOPES]);
    expect(cfg.logLevel).toBe("info");
    expect(cfg.tokenCacheFile).toMatch(/token-cache\.enc$/);
  });

  it("parses space- or comma-separated scope overrides", () => {
    expect(loadConfig({ ...base, MICROSOFT_SCOPES: "openid Mail.Read" }).scopes).toEqual([
      "openid",
      "Mail.Read",
    ]);
    expect(
      loadConfig({ ...base, MICROSOFT_SCOPES: "openid, Mail.Read , User.Read" }).scopes,
    ).toEqual(["openid", "Mail.Read", "User.Read"]);
  });

  it("honours a custom tenant and timezone", () => {
    const cfg = loadConfig({
      ...base,
      MICROSOFT_TENANT: "consumers",
      DEFAULT_TIMEZONE: "Europe/London",
    });
    expect(cfg.microsoftTenant).toBe("consumers");
    expect(cfg.authority).toBe("https://login.microsoftonline.com/consumers");
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
