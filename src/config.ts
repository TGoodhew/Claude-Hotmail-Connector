/**
 * Runtime configuration for the connector.
 *
 * Values come from environment variables: set by the MCP client (e.g. the
 * Claude Desktop extension's settings map "Application (client) ID" to
 * `MICROSOFT_CLIENT_ID`), optionally seeded from a local `.env` file in
 * development. {@link loadConfig} is a pure function of its `env` argument so it
 * can be unit-tested without mutating `process.env`.
 */

import { realpathSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { config as loadDotenv } from "dotenv";
import { z } from "zod";
import { ConfigError } from "./util/errors.js";
import { isValidTimeZone } from "./util/time.js";

/**
 * Where a development `.env` may live: the root of the connector's own checkout
 * (the parent of the folder holding the entry script, e.g. `<repo>/dist/index.js`
 * → `<repo>/.env`). It is deliberately **not** the current directory: Claude
 * Code launches MCP servers inside project folders, and a project's `.env` must
 * not be able to swap our client id, tenant, scopes or cache dir (#41, spec v2
 * §8.4). Extension installs have no `.env`.
 *
 * Symlinks are resolved first: Node keeps a symlinked entry script's own path in
 * `process.argv[1]`, so `project/bin/x.js -> <repo>/dist/index.js` would
 * otherwise point at `project/.env`.
 */
export function dotenvPath(entryScript: string | undefined): string | undefined {
  if (!entryScript) return undefined;
  let script = resolve(entryScript);
  try {
    script = realpathSync(script);
  } catch {
    // Not on disk (e.g. a test path): the resolved path is the best we have.
  }
  return join(dirname(dirname(script)), ".env");
}

// Seed process.env from that .env if present. Real environment wins.
// `quiet` avoids writing dotenv's banner to stdout (reserved for MCP JSON-RPC).
const devDotenv = dotenvPath(process.argv[1]);
if (devDotenv) loadDotenv({ path: devDotenv, quiet: true });

/** Delegated Microsoft Graph scopes requested at sign-in (least privilege). */
export const DEFAULT_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Calendars.ReadWrite",
] as const;

/**
 * Sending mail (#57): requested only when the user turns sending on
 * (`ALLOW_SEND_MAIL`). `Mail.Send` can only send, not read, move or delete.
 */
export const SEND_SCOPE = "Mail.Send";

/**
 * Personal Microsoft accounts only (spec v2 §8.1, #62): this connector exists
 * for hotmail/outlook/live accounts. Work/school accounts use Claude's official
 * M365 connector. Set `MICROSOFT_TENANT=common` with a multitenant registration
 * to allow them anyway.
 */
export const DEFAULT_TENANT = "consumers";
/** Used when `DEFAULT_TIMEZONE` is unset and the computer's own zone can't be read. */
export const FALLBACK_TIMEZONE = "UTC";

/**
 * The computer's own IANA time zone, or UTC if the runtime can't tell (#34).
 * It is the default for calendar calls that don't name a zone, so events land
 * at the user's local time wherever they are.
 */
export function systemTimeZone(): string {
  try {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return tz && isValidTimeZone(tz) ? tz : FALLBACK_TIMEZONE;
  } catch {
    return FALLBACK_TIMEZONE;
  }
}
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const LOOPBACK_HOST = "127.0.0.1";

/**
 * Shown when no client id is configured. Each user brings their own Entra app
 * registration (spec v2 §8.1); builds deliberately ship without a shared one.
 */
export const MISSING_CLIENT_ID_MESSAGE =
  "No Microsoft app registration is configured. Create your own Entra app registration " +
  '(see "Step 1: register your own Microsoft Entra application" in the README, or run ' +
  "scripts/register-app.ps1), " +
  "then enter its Application (client) ID in the extension's settings in Claude Desktop " +
  "(Settings > Extensions > Hotmail Connector), or set MICROSOFT_CLIENT_ID.";

const EnvSchema = z.object({
  // Optional in the schema so loadConfig can raise the friendlier MISSING_CLIENT_ID_MESSAGE.
  MICROSOFT_CLIENT_ID: z
    .string()
    .trim()
    .optional()
    .transform((v) => (v && v.length > 0 ? v : undefined)),
  MICROSOFT_TENANT: z.string().trim().min(1).default(DEFAULT_TENANT),
  // Blank means unset: fall back to the computer's own zone.
  DEFAULT_TIMEZONE: z
    .string()
    .trim()
    .optional()
    .transform((v) => (v && v.length > 0 ? v : undefined)),
  MICROSOFT_SCOPES: z.string().trim().optional(),
  // The extension passes its checkbox as "true" / "false"; blank means off.
  ALLOW_SEND_MAIL: z.string().trim().optional(),
  LOG_LEVEL: z.enum(["debug", "info", "warn", "error"]).default("info"),
  CLAUDE_HOTMAIL_CACHE_DIR: z.string().trim().optional(),
});

export interface AppConfig {
  /** Entra app registration (public client) id. */
  microsoftClientId: string;
  /** Tenant segment: `consumers` (personal accounts, default) or `common` (+ work/school). */
  microsoftTenant: string;
  /** Full authority base, e.g. https://login.microsoftonline.com/consumers */
  authority: string;
  /** Default IANA time zone for calendar reads/writes: `DEFAULT_TIMEZONE`, else the system zone. */
  defaultTimeZone: string;
  /** Delegated Graph scopes. */
  scopes: string[];
  /**
   * Whether sending mail is turned on (`ALLOW_SEND_MAIL`, #57). Off by default;
   * when on, `Mail.Send` is requested and the send tools are offered, each send
   * still needing the user's click in the browser.
   */
  allowSendMail: boolean;
  /** Microsoft Graph v1.0 base URL. */
  graphBaseUrl: string;
  /** Loopback host for the OAuth redirect listener. */
  loopbackHost: string;
  /** Directory holding the encrypted token cache. */
  cacheDir: string;
  /** Encrypted token cache file path. */
  tokenCacheFile: string;
  /** Log verbosity. */
  logLevel: "debug" | "info" | "warn" | "error";
}

/** Directory for the token cache (honours `CLAUDE_HOTMAIL_CACHE_DIR`). */
function resolveCacheDir(env: NodeJS.ProcessEnv): string {
  if (env.CLAUDE_HOTMAIL_CACHE_DIR) return env.CLAUDE_HOTMAIL_CACHE_DIR;
  if (process.platform === "win32" && env.LOCALAPPDATA) {
    return join(env.LOCALAPPDATA, "claude-hotmail-connector");
  }
  const xdg = env.XDG_STATE_HOME;
  if (xdg) return join(xdg, "claude-hotmail-connector");
  return join(homedir(), ".claude-hotmail-connector");
}

/**
 * Scopes `MICROSOFT_SCOPES` may choose from: the least-privilege default plus
 * `Calendars.Read`, the read-only narrowing of `Calendars.ReadWrite`. The
 * override can only narrow access, never broaden it (spec v2 §6, decision D11).
 * `Mail.Send` joins both lists only when sending is turned on (#57).
 */
export const ALLOWED_SCOPES = [...DEFAULT_SCOPES, "Calendars.Read"] as const;

function scopeLists(allowSend: boolean): { defaults: string[]; allowed: string[] } {
  const extra = allowSend ? [SEND_SCOPE] : [];
  return { defaults: [...DEFAULT_SCOPES, ...extra], allowed: [...ALLOWED_SCOPES, ...extra] };
}

/**
 * `ALLOW_SEND_MAIL`: true / false (also 1/0, yes/no, on/off); unset or blank is
 * off, and so is a setting the extension host left unfilled (`${...}`): a
 * missing switch must never stop the server, nor turn sending on.
 */
function parseAllowSend(raw: string | undefined): boolean {
  const v = (raw ?? "").toLowerCase();
  if (v.startsWith("${")) return false;
  if (["", "false", "0", "no", "off"].includes(v)) return false;
  if (["true", "1", "yes", "on"].includes(v)) return true;
  throw new ConfigError(`ALLOW_SEND_MAIL must be true or false; got "${raw}".`);
}

/** Always requested: needed for sign-in and silent refresh. */
export const REQUIRED_SCOPES = ["openid", "offline_access"] as const;

function parseScopes(raw: string | undefined, allowSend: boolean): string[] {
  const { defaults, allowed } = scopeLists(allowSend);
  if (!raw) return defaults;
  const requested = raw
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (requested.length === 0) return defaults;

  // Graph scope names are case-insensitive; normalise to the canonical spelling.
  const canonical = new Map<string, string>(allowed.map((s) => [s.toLowerCase(), s]));
  const rejected = requested.filter((s) => !canonical.has(s.toLowerCase()));
  if (rejected.length > 0) {
    const sendHint = rejected.some((s) => s.toLowerCase() === SEND_SCOPE.toLowerCase())
      ? " To send mail, set ALLOW_SEND_MAIL=true instead."
      : "";
    throw new ConfigError(
      `MICROSOFT_SCOPES may only narrow the default scopes; not allowed: ${rejected.join(", ")}. ` +
        `Allowed: ${allowed.join(" ")}.${sendHint}`,
    );
  }
  const chosen = new Set(requested.map((s) => canonical.get(s.toLowerCase())!));
  for (const s of REQUIRED_SCOPES) chosen.add(s);
  return allowed.filter((s) => chosen.has(s));
}

/**
 * Build a validated {@link AppConfig} from an environment object.
 * Throws {@link ConfigError} with a friendly, aggregated message on failure,
 * including when no client id is configured.
 */
export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const parsed = EnvSchema.safeParse(env);
  if (!parsed.success) {
    const details = parsed.error.issues
      .map((i) => (i.path.length ? `${i.path.join(".")}: ${i.message}` : i.message))
      .join("; ");
    throw new ConfigError(`Invalid configuration: ${details}`);
  }

  const e = parsed.data;
  const cacheDir = resolveCacheDir(env);
  const allowSendMail = parseAllowSend(e.ALLOW_SEND_MAIL);
  const clientId = e.MICROSOFT_CLIENT_ID;
  if (!clientId) throw new ConfigError(MISSING_CLIENT_ID_MESSAGE);
  // Checked here so a typo fails at startup, not on every calendar call.
  if (e.DEFAULT_TIMEZONE && !isValidTimeZone(e.DEFAULT_TIMEZONE)) {
    throw new ConfigError(
      `DEFAULT_TIMEZONE "${e.DEFAULT_TIMEZONE}" is not an IANA time zone. Use a name like ` +
        `"Europe/London" or "UTC", or leave it unset to use this computer's time zone.`,
    );
  }

  return {
    microsoftClientId: clientId,
    microsoftTenant: e.MICROSOFT_TENANT,
    authority: `https://login.microsoftonline.com/${e.MICROSOFT_TENANT}`,
    defaultTimeZone: e.DEFAULT_TIMEZONE ?? systemTimeZone(),
    scopes: parseScopes(e.MICROSOFT_SCOPES, allowSendMail),
    allowSendMail,
    graphBaseUrl: GRAPH_BASE_URL,
    loopbackHost: LOOPBACK_HOST,
    cacheDir,
    tokenCacheFile: join(cacheDir, "token-cache.enc"),
    logLevel: e.LOG_LEVEL,
  };
}

let cached: AppConfig | undefined;

/** Memoised config built from `process.env`. Use in application code. */
export function getConfig(): AppConfig {
  cached ??= loadConfig(process.env);
  return cached;
}

/** Reset the memoised config (test helper). */
export function resetConfigCache(): void {
  cached = undefined;
}
