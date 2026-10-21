/**
 * Runtime configuration for the connector.
 *
 * Values come from environment variables (optionally seeded from a local
 * `.env` file, which is git-ignored) and, for the client id, the persisted
 * user settings file (see `settings.ts`). {@link loadConfig} is a pure function
 * of its `env` and `settings` arguments so it can be unit-tested without
 * mutating `process.env` or touching the filesystem.
 */

import { homedir } from "node:os";
import { join } from "node:path";
import { config as loadDotenv } from "dotenv";
import { z } from "zod";
import { readSettings, type UserSettings } from "./settings.js";
import { ConfigError } from "./util/errors.js";

// Seed process.env from a local .env if present. Real environment wins.
// `quiet` avoids writing dotenv's banner to stdout (reserved for MCP JSON-RPC).
loadDotenv({ quiet: true });

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

export const DEFAULT_TENANT = "common";
export const DEFAULT_TIMEZONE = "Australia/Brisbane";
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const LOOPBACK_HOST = "127.0.0.1";

/**
 * Shown when no client id is configured. Each user brings their own Entra app
 * registration (spec v2 §8.1); builds deliberately ship without a shared one.
 */
export const MISSING_CLIENT_ID_MESSAGE =
  "No Microsoft app registration is configured. Create your own Entra app registration " +
  '(see "Register a Microsoft Entra application" in the README, or run scripts/register-app.ps1), ' +
  "then run `setup --client-id <Application (client) ID>` or set MICROSOFT_CLIENT_ID.";

const EnvSchema = z.object({
  // Optional here: falls back to the saved settings; loadConfig errors if neither is set.
  MICROSOFT_CLIENT_ID: z
    .string()
    .trim()
    .optional()
    .transform((v) => (v && v.length > 0 ? v : undefined)),
  MICROSOFT_TENANT: z.string().trim().min(1).default(DEFAULT_TENANT),
  DEFAULT_TIMEZONE: z.string().trim().min(1).default(DEFAULT_TIMEZONE),
  MICROSOFT_SCOPES: z.string().trim().optional(),
  LOG_LEVEL: z.enum(["debug", "info", "warn", "error"]).default("info"),
  CLAUDE_HOTMAIL_CACHE_DIR: z.string().trim().optional(),
});

export interface AppConfig {
  /** Entra app registration (public client) id. */
  microsoftClientId: string;
  /** Tenant segment; `common` enables personal + work/school accounts. */
  microsoftTenant: string;
  /** Full authority base, e.g. https://login.microsoftonline.com/common */
  authority: string;
  /** Default IANA time zone for calendar reads/writes. */
  defaultTimeZone: string;
  /** Delegated Graph scopes. */
  scopes: string[];
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

/** Directory for the token cache and settings file (honours `CLAUDE_HOTMAIL_CACHE_DIR`). */
export function resolveCacheDir(env: NodeJS.ProcessEnv): string {
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
 */
export const ALLOWED_SCOPES = [...DEFAULT_SCOPES, "Calendars.Read"] as const;

/** Always requested: needed for sign-in and silent refresh. */
export const REQUIRED_SCOPES = ["openid", "offline_access"] as const;

function parseScopes(raw: string | undefined): string[] {
  if (!raw) return [...DEFAULT_SCOPES];
  const requested = raw
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (requested.length === 0) return [...DEFAULT_SCOPES];

  // Graph scope names are case-insensitive; normalise to the canonical spelling.
  const canonical = new Map<string, string>(ALLOWED_SCOPES.map((s) => [s.toLowerCase(), s]));
  const rejected = requested.filter((s) => !canonical.has(s.toLowerCase()));
  if (rejected.length > 0) {
    throw new ConfigError(
      `MICROSOFT_SCOPES may only narrow the default scopes; not allowed: ${rejected.join(", ")}. ` +
        `Allowed: ${ALLOWED_SCOPES.join(" ")}.`,
    );
  }
  const chosen = new Set(requested.map((s) => canonical.get(s.toLowerCase())!));
  for (const s of REQUIRED_SCOPES) chosen.add(s);
  return ALLOWED_SCOPES.filter((s) => chosen.has(s));
}

/**
 * Build a validated {@link AppConfig} from an environment object and the saved
 * user settings. `MICROSOFT_CLIENT_ID` wins over `settings.clientId`.
 * Throws {@link ConfigError} with a friendly, aggregated message on failure,
 * including when no client id is configured at all.
 */
export function loadConfig(
  env: NodeJS.ProcessEnv = process.env,
  settings: UserSettings = {},
): AppConfig {
  const parsed = EnvSchema.safeParse(env);
  if (!parsed.success) {
    const details = parsed.error.issues
      .map((i) => (i.path.length ? `${i.path.join(".")}: ${i.message}` : i.message))
      .join("; ");
    throw new ConfigError(`Invalid configuration: ${details}`);
  }

  const e = parsed.data;
  const cacheDir = resolveCacheDir(env);
  const clientId = e.MICROSOFT_CLIENT_ID ?? settings.clientId;
  if (!clientId) throw new ConfigError(MISSING_CLIENT_ID_MESSAGE);

  return {
    microsoftClientId: clientId,
    microsoftTenant: e.MICROSOFT_TENANT,
    authority: `https://login.microsoftonline.com/${e.MICROSOFT_TENANT}`,
    defaultTimeZone: e.DEFAULT_TIMEZONE,
    scopes: parseScopes(e.MICROSOFT_SCOPES),
    graphBaseUrl: GRAPH_BASE_URL,
    loopbackHost: LOOPBACK_HOST,
    cacheDir,
    tokenCacheFile: join(cacheDir, "token-cache.enc"),
    logLevel: e.LOG_LEVEL,
  };
}

let cached: AppConfig | undefined;

/** Memoised config built from `process.env` and the saved settings. Use in application code. */
export function getConfig(): AppConfig {
  cached ??= loadConfig(process.env, readSettings(resolveCacheDir(process.env)));
  return cached;
}

/** Reset the memoised config (test helper). */
export function resetConfigCache(): void {
  cached = undefined;
}
