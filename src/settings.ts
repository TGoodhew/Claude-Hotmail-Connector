/**
 * Persisted user settings: `settings.json` in the connector's cache directory.
 *
 * Holds the user's own Entra **client id** so that commands launched without
 * any environment (the installer's post-install `login`, the Start-menu
 * "Sign in" shortcut, a bare `hotmail-connector whoami`) still know which app
 * registration to sign in with. `setup --client-id <id>` writes it. The id is
 * a public-client identifier, not a secret.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { z } from "zod";
import { ConfigError } from "./util/errors.js";

export const SETTINGS_FILE_NAME = "settings.json";

// Loose: unknown keys survive a read-modify-write.
const SettingsSchema = z
  .object({
    clientId: z.string().trim().min(1).optional(),
  })
  .loose();

export type UserSettings = z.infer<typeof SettingsSchema>;

/** Entra application (client) ids are GUIDs. */
const CLIENT_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isValidClientId(id: string): boolean {
  return CLIENT_ID_RE.test(id.trim());
}

export function settingsFilePath(cacheDir: string): string {
  return join(cacheDir, SETTINGS_FILE_NAME);
}

/** Read settings; a missing file yields `{}`. A corrupt file is a {@link ConfigError}. */
export function readSettings(cacheDir: string): UserSettings {
  const path = settingsFilePath(cacheDir);
  if (!existsSync(path)) return {};
  let raw: unknown;
  try {
    raw = JSON.parse(readFileSync(path, "utf8").replace(/^\uFEFF/, ""));
  } catch (e) {
    throw new ConfigError(`Settings file is not valid JSON: ${path}`, e);
  }
  const parsed = SettingsSchema.safeParse(raw);
  if (!parsed.success) throw new ConfigError(`Settings file is invalid: ${path}`);
  return parsed.data;
}

/** Merge `patch` into the saved settings and write them back. Returns the file path. */
export function writeSettings(cacheDir: string, patch: UserSettings): string {
  const path = settingsFilePath(cacheDir);
  const next = { ...readSettings(cacheDir), ...patch };
  mkdirSync(cacheDir, { recursive: true });
  writeFileSync(path, JSON.stringify(next, null, 2) + "\n");
  return path;
}
