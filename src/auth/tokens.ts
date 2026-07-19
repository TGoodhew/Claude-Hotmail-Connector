/**
 * Secure, at-rest token cache.
 *
 * The Microsoft **refresh token** is the long-lived secret that lets the
 * connector act on the account without re-prompting. It is stored encrypted
 * with AES-256-GCM in the user's profile directory, with the encryption key in
 * a sibling file created with owner-only permissions. This matches the
 * local-hosting spec's "encrypted file (never plaintext, never in the repo)"
 * requirement for Mode A. (An OS keychain backend can be added later behind the
 * same {@link TokenStore} interface.)
 */

import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { chmod, mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { z } from "zod";
import { AuthError } from "../util/errors.js";

const ALGO = "aes-256-gcm";
const KEY_BYTES = 32;
const IV_BYTES = 12;
const CONTAINER_VERSION = 1;

/** A set of Microsoft tokens plus metadata needed to refresh them. */
export interface TokenSet {
  /** Graph access token (short-lived). */
  accessToken: string;
  /** Refresh token (long-lived secret). */
  refreshToken: string;
  /** Epoch milliseconds at which {@link accessToken} expires. */
  expiresAt: number;
  /** Scopes the tokens were granted. */
  scopes: string[];
  /** Non-secret account hint (e.g. username), for diagnostics. */
  account?: string;
  /** Usually "Bearer". */
  tokenType?: string;
}

const TokenSetSchema = z.object({
  accessToken: z.string().min(1),
  refreshToken: z.string().min(1),
  expiresAt: z.number().finite(),
  scopes: z.array(z.string()),
  account: z.string().optional(),
  tokenType: z.string().optional(),
});

/** Persistent store for a single account's {@link TokenSet}. */
export interface TokenStore {
  /** Return the cached tokens, or null if none are stored. */
  load(): Promise<TokenSet | null>;
  /** Encrypt and persist the tokens. */
  save(tokens: TokenSet): Promise<void>;
  /** Remove the cached tokens (used on sign-out or corruption). */
  clear(): Promise<void>;
}

interface Container {
  v: number;
  iv: string;
  tag: string;
  data: string;
}

function encrypt(key: Buffer, plaintext: string): Container {
  const iv = randomBytes(IV_BYTES);
  const cipher = createCipheriv(ALGO, key, iv);
  const ct = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  return {
    v: CONTAINER_VERSION,
    iv: iv.toString("base64"),
    tag: cipher.getAuthTag().toString("base64"),
    data: ct.toString("base64"),
  };
}

function decrypt(key: Buffer, c: Container): string {
  const decipher = createDecipheriv(ALGO, key, Buffer.from(c.iv, "base64"));
  decipher.setAuthTag(Buffer.from(c.tag, "base64"));
  // .final() throws if the GCM auth tag does not verify (tamper/corruption).
  return Buffer.concat([decipher.update(Buffer.from(c.data, "base64")), decipher.final()]).toString(
    "utf8",
  );
}

function isErrnoException(e: unknown): e is NodeJS.ErrnoException {
  return e instanceof Error && "code" in e;
}

async function bestEffortChmod(path: string, mode: number): Promise<void> {
  try {
    await chmod(path, mode);
  } catch {
    // chmod is a no-op / may fail on some Windows filesystems; the file already
    // lives under the ACL-protected user profile. Ignore.
  }
}

/**
 * File-backed, AES-256-GCM encrypted {@link TokenStore}.
 *
 * @param tokenCacheFile absolute path to the ciphertext file (e.g. token-cache.enc)
 * @param keyFile        optional key path; defaults to token-cache.key beside it
 */
export function createFileTokenStore(tokenCacheFile: string, keyFile?: string): TokenStore {
  const dir = dirname(tokenCacheFile);
  const resolvedKeyFile = keyFile ?? join(dir, "token-cache.key");

  async function ensureDir(): Promise<void> {
    await mkdir(dir, { recursive: true, mode: 0o700 });
    await bestEffortChmod(dir, 0o700);
  }

  async function loadOrCreateKey(): Promise<Buffer> {
    await ensureDir();
    try {
      const b64 = (await readFile(resolvedKeyFile, "utf8")).trim();
      const key = Buffer.from(b64, "base64");
      if (key.length === KEY_BYTES) return key;
      // Wrong size => treat as corrupt and regenerate (invalidates old cache).
    } catch (e) {
      if (isErrnoException(e) && e.code !== "ENOENT") throw e;
    }
    const key = randomBytes(KEY_BYTES);
    await writeFile(resolvedKeyFile, key.toString("base64"), { mode: 0o600 });
    await bestEffortChmod(resolvedKeyFile, 0o600);
    return key;
  }

  async function writeAtomic(path: string, contents: string): Promise<void> {
    const tmp = `${path}.tmp-${randomBytes(6).toString("hex")}`;
    await writeFile(tmp, contents, { mode: 0o600 });
    await bestEffortChmod(tmp, 0o600);
    await rename(tmp, path); // atomic replace (POSIX and Windows via MoveFileEx)
  }

  return {
    async load(): Promise<TokenSet | null> {
      let raw: string;
      try {
        raw = await readFile(tokenCacheFile, "utf8");
      } catch (e) {
        if (isErrnoException(e) && e.code === "ENOENT") return null;
        throw e;
      }
      try {
        const container = JSON.parse(raw) as Container;
        const key = await loadOrCreateKey();
        const plaintext = decrypt(key, container);
        return TokenSetSchema.parse(JSON.parse(plaintext));
      } catch (e) {
        throw new AuthError(
          "Token cache is unreadable (corrupt, tampered, or key mismatch). Re-authenticate to reset it.",
          e,
        );
      }
    },

    async save(tokens: TokenSet): Promise<void> {
      const parsed = TokenSetSchema.parse(tokens);
      const key = await loadOrCreateKey();
      const container = encrypt(key, JSON.stringify(parsed));
      await writeAtomic(tokenCacheFile, JSON.stringify(container));
    },

    async clear(): Promise<void> {
      try {
        await unlink(tokenCacheFile);
      } catch (e) {
        if (isErrnoException(e) && e.code !== "ENOENT") throw e;
      }
    },
  };
}
