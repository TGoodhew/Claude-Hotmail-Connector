import { describe, expect, it } from "vitest";
import { DEFAULT_SCOPES, loadConfig, MISSING_CLIENT_ID_MESSAGE } from "./config.js";
import { ConfigError } from "./util/errors.js";

const base = { MICROSOFT_CLIENT_ID: "11111111-2222-3333-4444-555555555555" };
const savedId = "99999999-8888-7777-6666-555555555555";

describe("loadConfig", () => {
  it("requires a client id: there is no bundled shared default", () => {
    expect(() => loadConfig({})).toThrow(ConfigError);
    expect(() => loadConfig({ MICROSOFT_CLIENT_ID: "   " })).toThrow(MISSING_CLIENT_ID_MESSAGE);
  });

  it("uses the saved settings client id when the env var is absent", () => {
    expect(loadConfig({}, { clientId: savedId }).microsoftClientId).toBe(savedId);
  });

  it("prefers MICROSOFT_CLIENT_ID over the saved settings", () => {
    expect(loadConfig(base, { clientId: savedId }).microsoftClientId).toBe(
      base.MICROSOFT_CLIENT_ID,
    );
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
