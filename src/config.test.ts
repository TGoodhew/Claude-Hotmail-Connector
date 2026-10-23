import { mkdirSync, mkdtempSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  DEFAULT_SCOPES,
  dotenvPath,
  FALLBACK_TIMEZONE,
  loadConfig,
  MISSING_CLIENT_ID_MESSAGE,
  SEND_SCOPE,
  systemTimeZone,
} from "./config.js";
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

  it("follows a symlinked entry script back to the real checkout", (ctx) => {
    const root = realpathSync(mkdtempSync(join(tmpdir(), "cht-dotenv-")));
    try {
      const checkout = join(root, "checkout");
      const project = join(root, "project");
      mkdirSync(join(checkout, "dist"), { recursive: true });
      mkdirSync(join(project, "bin"), { recursive: true });
      writeFileSync(join(checkout, "dist", "index.js"), "");
      const link = join(project, "bin", "connector.js");
      try {
        symlinkSync(join(checkout, "dist", "index.js"), link, "file");
      } catch {
        ctx.skip(); // e.g. Windows without symlink privilege
      }
      expect(dotenvPath(link)).toBe(join(checkout, ".env"));
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
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
    expect(cfg.defaultTimeZone).toBe(systemTimeZone());
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

  describe("default time zone (#34)", () => {
    afterEach(() => vi.restoreAllMocks());
    const systemZone = (timeZone: string) =>
      vi
        .spyOn(Intl.DateTimeFormat.prototype, "resolvedOptions")
        .mockReturnValue({ timeZone } as Intl.ResolvedDateTimeFormatOptions);

    it("defaults to the computer's own zone, not a hard-coded one", () => {
      systemZone("America/New_York");
      expect(loadConfig(base).defaultTimeZone).toBe("America/New_York");
      // Blank, as an empty extension setting arrives, means unset.
      expect(loadConfig({ ...base, DEFAULT_TIMEZONE: "  " }).defaultTimeZone).toBe(
        "America/New_York",
      );
    });

    it("falls back to UTC when the system zone is missing or unusable", () => {
      systemZone("");
      expect(loadConfig(base).defaultTimeZone).toBe(FALLBACK_TIMEZONE);
      vi.restoreAllMocks();
      systemZone("Not/AZone");
      expect(loadConfig(base).defaultTimeZone).toBe("UTC");
    });

    it("rejects an invalid DEFAULT_TIMEZONE at startup", () => {
      const load = () => loadConfig({ ...base, DEFAULT_TIMEZONE: "Mars/Olympus" });
      expect(load).toThrow(ConfigError);
      expect(load).toThrow(/not an IANA time zone/);
    });
  });

  it("honours an explicit cache dir override", () => {
    const cfg = loadConfig({ ...base, CLAUDE_HOTMAIL_CACHE_DIR: "/tmp/cache-x" });
    expect(cfg.cacheDir).toBe("/tmp/cache-x");
  });

  it("rejects an invalid LOG_LEVEL", () => {
    expect(() => loadConfig({ ...base, LOG_LEVEL: "verbose" })).toThrow(ConfigError);
  });
});

describe("ALLOW_SEND_MAIL (#57)", () => {
  it("is off unless turned on, and only then requests Mail.Send", () => {
    expect(loadConfig(base)).toMatchObject({ allowSendMail: false, scopes: [...DEFAULT_SCOPES] });
    for (const off of ["", "false", "0", "no", "OFF", "${user_config.allow_send}"]) {
      expect(loadConfig({ ...base, ALLOW_SEND_MAIL: off }).allowSendMail, off).toBe(false);
    }
    for (const on of ["true", "1", "yes", "On"]) {
      const cfg = loadConfig({ ...base, ALLOW_SEND_MAIL: on });
      expect(cfg.allowSendMail, on).toBe(true);
      expect(cfg.scopes).toEqual([...DEFAULT_SCOPES, SEND_SCOPE]);
    }
    expect(() => loadConfig({ ...base, ALLOW_SEND_MAIL: "maybe" })).toThrow(ConfigError);
  });

  it("can't be granted through MICROSOFT_SCOPES, which only narrows", () => {
    expect(() => loadConfig({ ...base, MICROSOFT_SCOPES: "Mail.Read Mail.Send" })).toThrow(
      /not allowed: Mail.Send.*set ALLOW_SEND_MAIL=true instead/,
    );
    // With sending on, narrowing may keep it, or drop it.
    const keep = loadConfig({
      ...base,
      ALLOW_SEND_MAIL: "true",
      MICROSOFT_SCOPES: "Mail.Read mail.send",
    });
    expect(keep.scopes).toEqual(["openid", "offline_access", "Mail.Read", SEND_SCOPE]);
    const drop = loadConfig({ ...base, ALLOW_SEND_MAIL: "true", MICROSOFT_SCOPES: "Mail.Read" });
    expect(drop.scopes).not.toContain(SEND_SCOPE);
  });
});
