/**
 * Runtime configuration for the connector.
 *
 * Values come from environment variables (optionally seeded from a local
 * `.env` file, which is git-ignored). {@link loadConfig} is a pure function of
 * its `env` argument so it can be unit-tested without mutating `process.env`.
 */

import { homedir } from "node:os";
import { join } from "node:path";
import { config as loadDotenv } from "dotenv";
import { z } from "zod";
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

const EnvSchema = z.object({
  MICROSOFT_CLIENT_ID: z
    .string({ message: "MICROSOFT_CLIENT_ID is required" })
    .trim()
    .min(1, "MICROSOFT_CLIENT_ID is required (the Entra app registration client id)"),
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

function defaultCacheDir(env: NodeJS.ProcessEnv): string {
  if (env.CLAUDE_HOTMAIL_CACHE_DIR) return env.CLAUDE_HOTMAIL_CACHE_DIR;
  if (process.platform === "win32" && env.LOCALAPPDATA) {
    return join(env.LOCALAPPDATA, "claude-hotmail-connector");
  }
  const xdg = env.XDG_STATE_HOME;
  if (xdg) return join(xdg, "claude-hotmail-connector");
  return join(homedir(), ".claude-hotmail-connector");
}

function parseScopes(raw: string | undefined): string[] {
  if (!raw) return [...DEFAULT_SCOPES];
  const scopes = raw
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  return scopes.length > 0 ? scopes : [...DEFAULT_SCOPES];
}

/**
 * Build a validated {@link AppConfig} from an environment object.
 * Throws {@link ConfigError} with a friendly, aggregated message on failure.
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
  const cacheDir = defaultCacheDir(env);

  return {
    microsoftClientId: e.MICROSOFT_CLIENT_ID,
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

/** Memoised config built from `process.env`. Use in application code. */
export function getConfig(): AppConfig {
  cached ??= loadConfig(process.env);
  return cached;
}

/** Reset the memoised config (test helper). */
export function resetConfigCache(): void {
  cached = undefined;
}
