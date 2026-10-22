/**
 * Secure, at-rest token cache.
 *
 * The Microsoft **refresh token** is the long-lived secret that lets the
 * connector act on the account without re-prompting. It is stored encrypted
 * with AES-256-GCM in the user's profile directory. The AES key lives in a
 * sibling key file, **wrapped by the OS** (DPAPI on Windows, see
 * `protect.ts`) so that copying the folder isn't enough to read the cache
 * (spec v2 §8.3). Where no OS protector is available, or it fails, the key is
 * stored as plain base64 in an owner-only file instead, and a warning is logged.
 *
 * Key file formats:
 *  - legacy / fallback: the base64 key itself;
 *  - protected: `{"v":2,"protection":"<protector>","key":"<base64 wrapped key>"}`.
 * A legacy key is upgraded in place when a protector becomes available, so
 * existing sign-ins survive the upgrade.
 */

import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { chmod, link, mkdir, readFile, rename, stat, unlink, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { z } from "zod";
import { AuthError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";
import type { KeyProtector } from "./protect.js";

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
  /**
   * The shared sign-in count (#95) this sign-in was saved under, kept across
   * refreshes: tells a later browser sign-in whether the cached account is the
   * one that won, or an older one left by a winner whose save failed.
   */
  signInGeneration?: number;
}

const TokenSetSchema = z.object({
  accessToken: z.string().min(1),
  refreshToken: z.string().min(1),
  expiresAt: z.number().finite(),
  scopes: z.array(z.string()),
  account: z.string().optional(),
  tokenType: z.string().optional(),
  signInGeneration: z.number().int().nonnegative().optional(),
});

/** Persistent store for a single account's {@link TokenSet}. */
export interface TokenStore {
  /** Return the cached tokens, or null if none are stored. */
  load(): Promise<TokenSet | null>;
  /** Encrypt and persist the tokens. */
  save(tokens: TokenSet): Promise<void>;
  /**
   * Remove the cached tokens (on corruption, or sign-out). With `includeKey`,
   * also remove the encryption key, leaving nothing of the sign-in on disk.
   */
  clear(options?: { includeKey?: boolean }): Promise<void>;
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

function safeJson(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return undefined;
  }
}

async function bestEffortChmod(path: string, mode: number): Promise<void> {
  try {
    await chmod(path, mode);
  } catch {
    // chmod is a no-op / may fail on some Windows filesystems; the file already
    // lives under the ACL-protected user profile. Ignore.
  }
}

const ProtectedKeyFileSchema = z.object({
  v: z.literal(2),
  protection: z.string().min(1),
  key: z.string().min(1),
});

export interface FileTokenStoreOptions {
  /** Key file path; defaults to `token-cache.key` beside the cache. */
  keyFile?: string;
  /** OS protector for the key (see `defaultKeyProtector`); `null`/omitted = plain key file. */
  protector?: KeyProtector | null;
  logger?: Logger;
}

/**
 * File-backed, AES-256-GCM encrypted {@link TokenStore}.
 *
 * @param tokenCacheFile absolute path to the ciphertext file (e.g. token-cache.enc)
 */
export function createFileTokenStore(
  tokenCacheFile: string,
  opts: FileTokenStoreOptions = {},
): TokenStore {
  const dir = dirname(tokenCacheFile);
  const resolvedKeyFile = opts.keyFile ?? join(dir, "token-cache.key");
  const protector = opts.protector ?? null;
  const log = opts.logger ?? createLogger({ name: "tokens" });

  async function ensureDir(): Promise<void> {
    await mkdir(dir, { recursive: true, mode: 0o700 });
    await bestEffortChmod(dir, 0o700);
  }

  /**
   * Write the key file. `exclusive` publishes it only if none exists yet
   * (atomically, via a hard link), and returns false if another process won
   * the race to create it. Otherwise it replaces the file atomically.
   */
  async function writeKeyFile(key: Buffer, exclusive: boolean): Promise<boolean> {
    let contents = key.toString("base64");
    if (protector) {
      try {
        const wrapped = await protector.protect(key);
        contents = JSON.stringify({
          v: 2,
          protection: protector.name,
          key: wrapped.toString("base64"),
        });
      } catch (e) {
        log.warn("OS key protection failed; storing the token-cache key unprotected", {
          protection: protector.name,
          error: (e as Error).message,
        });
      }
    }
    if (exclusive) return publishExclusive(resolvedKeyFile, contents);
    await writeAtomic(resolvedKeyFile, contents);
    return true;
  }

  type KeyRead =
    | { status: "ok"; key: Buffer; rewrote?: boolean; legacy?: boolean }
    | { status: "absent" }
    | { status: "unusable" };

  /**
   * Read the key file. "unusable" = present but corrupt or foreign (a writer
   * replaces it). `upgrade` wraps a legacy plain key in place, so only writers
   * pass it: loads never write to disk.
   */
  async function readKey(upgrade: boolean): Promise<KeyRead> {
    let raw: string;
    try {
      raw = (await readFile(resolvedKeyFile, "utf8")).trim();
    } catch (e) {
      if (isErrnoException(e) && e.code === "ENOENT") return { status: "absent" };
      throw e;
    }

    if (raw.startsWith("{")) {
      const parsed = ProtectedKeyFileSchema.safeParse(safeJson(raw));
      if (!parsed.success || !protector || protector.name !== parsed.data.protection) {
        log.warn("Token-cache key is protected in a way this machine can't unwrap; resetting it");
        return { status: "unusable" };
      }
      try {
        const key = await protector.unprotect(Buffer.from(parsed.data.key, "base64"));
        return key.length === KEY_BYTES ? { status: "ok", key } : { status: "unusable" };
      } catch (e) {
        // e.g. copied from another user/machine: the cache can't be read here.
        log.warn("Could not unwrap the token-cache key; resetting it", {
          error: (e as Error).message,
        });
        return { status: "unusable" };
      }
    }

    const key = Buffer.from(raw, "base64");
    // Corrupt => regenerate (invalidates the old cache).
    if (key.length !== KEY_BYTES) return { status: "unusable" };
    if (protector && upgrade) {
      await writeKeyFile(key, false); // upgrade a legacy plain key in place
      return { status: "ok", key, rewrote: true };
    }
    return { status: "ok", key, legacy: protector !== null };
  }

  /** Identity of the key file on disk; changes whenever any process replaces it. */
  async function keyFileSignature(): Promise<string | null> {
    try {
      const st = await stat(resolvedKeyFile);
      return `${st.ino}:${st.size}:${st.mtimeMs}`;
    } catch (e) {
      if (isErrnoException(e) && e.code === "ENOENT") return null;
      throw e;
    }
  }

  // The key is cached against the key file's identity: unwrapping may spawn a
  // process (DPAPI), so it happens again only if another process replaced the
  // file. A stale key would write a cache that fresh processes can't decrypt.
  // `needsUpgrade`: a legacy plain key read by a load; the next save wraps it.
  let cached: { key: Buffer; signature: string | null; needsUpgrade?: boolean } | null = null;
  // Key work in progress, per mode: "read" (load) never writes; "write" (save)
  // may create, reset or upgrade the key file.
  const inflight: { read: Promise<Buffer | null> | null; write: Promise<Buffer> | null } = {
    read: null,
    write: null,
  };

  async function remember(key: Buffer): Promise<Buffer> {
    cached = { key, signature: await keyFileSignature() };
    return key;
  }

  /** The key for reading the cache, or null if there's no usable one. Never writes. */
  function readCurrentKey(): Promise<Buffer | null> {
    inflight.read ??= (async () => {
      const signature = await keyFileSignature();
      if (signature === null) return null;
      if (cached && signature === cached.signature) return cached.key;
      for (let attempt = 0; attempt < 3; attempt++) {
        const before = await keyFileSignature();
        const read = await readKey(false);
        if (read.status !== "ok") return null;
        const after = await keyFileSignature();
        if (before !== after) continue; // replaced while we read it: read again
        cached = { key: read.key, signature: after, needsUpgrade: read.legacy };
        return read.key;
      }
      return null;
    })().finally(() => {
      inflight.read = null;
    });
    return inflight.read;
  }

  /** The key for writing the cache: creates, resets or upgrades the key file as needed. */
  function loadOrCreateKey(): Promise<Buffer> {
    inflight.write ??= (async () => {
      await ensureDir();
      const signature = await keyFileSignature();
      if (cached && signature !== null && signature === cached.signature && !cached.needsUpgrade) {
        return cached.key;
      }
      for (let attempt = 0; attempt < 3; attempt++) {
        // Note the file's identity around the read: if another process replaced
        // it meanwhile (e.g. while DPAPI was unwrapping), the key we got belongs
        // to the old file, so read again rather than pair it with the new one.
        const before = await keyFileSignature();
        const read = await readKey(true);
        if (read.status === "ok") {
          const after = await keyFileSignature();
          if (before !== after && !read.rewrote) continue;
          cached = { key: read.key, signature: after };
          return read.key;
        }
        const key = randomBytes(KEY_BYTES);
        if (read.status === "unusable") {
          // Only reset the file we judged unusable: if another process replaced it
          // meanwhile (e.g. while our unwrap was failing), adopt its key instead.
          if ((await keyFileSignature()) !== before) continue;
          await writeKeyFile(key, false);
          return remember(key);
        }
        // Absent: create it exclusively. If another process got there first,
        // loop round and adopt its key rather than overwriting it.
        if (await writeKeyFile(key, true)) return remember(key);
      }
      throw new AuthError("Could not establish the token-cache key.");
    })().finally(() => {
      inflight.write = null;
    });
    return inflight.write;
  }

  async function writeAtomic(path: string, contents: string): Promise<void> {
    const tmp = `${path}.tmp-${randomBytes(6).toString("hex")}`;
    await writeFile(tmp, contents, { mode: 0o600 });
    await bestEffortChmod(tmp, 0o600);
    await rename(tmp, path); // atomic replace (POSIX and Windows via MoveFileEx)
  }

  /** Create `path` only if it doesn't exist, atomically. Returns false if it already exists. */
  async function publishExclusive(path: string, contents: string): Promise<boolean> {
    const tmp = `${path}.tmp-${randomBytes(6).toString("hex")}`;
    await writeFile(tmp, contents, { mode: 0o600 });
    await bestEffortChmod(tmp, 0o600);
    try {
      await link(tmp, path); // fails with EEXIST if another process published first
      return true;
    } catch (e) {
      if (isErrnoException(e) && e.code === "EEXIST") return false;
      if (isErrnoException(e) && ["EPERM", "ENOTSUP", "EXDEV", "ENOSYS"].includes(e.code ?? "")) {
        // No hard links here (e.g. some FAT/network volumes). An exclusive `wx`
        // create would expose a half-written file to concurrent readers, so use an
        // atomic rename instead: readers never see partial content, but two
        // processes creating a key at the same instant can race (last one wins).
        log.warn("Hard links unavailable for the token cache; key creation isn't exclusive here");
        await writeAtomic(path, contents);
        return true;
      }
      throw e;
    } finally {
      await unlink(tmp).catch(() => undefined);
    }
  }

  return {
    async load(): Promise<TokenSet | null> {
      // Another process may replace the key *and* the cache between our reading
      // the ciphertext and getting the key. The pair then doesn't match even though
      // both files are fine, so read both again once before calling it corrupt
      // (which makes the caller clear the cache and lose the fresh tokens).
      for (let attempt = 0; ; attempt++) {
        let raw: string;
        try {
          raw = await readFile(tokenCacheFile, "utf8");
        } catch (e) {
          if (isErrnoException(e) && e.code === "ENOENT") return null;
          throw e;
        }
        try {
          const container = JSON.parse(raw) as Container;
          // Read-only: a load never creates, resets or upgrades the key file, so
          // a load that outlives a full cleanup can't put a key back on disk.
          const key = await readCurrentKey();
          if (!key) throw new AuthError("No usable token-cache key.");
          const plaintext = decrypt(key, container);
          return TokenSetSchema.parse(JSON.parse(plaintext));
        } catch (e) {
          if (attempt === 0) continue;
          throw new AuthError(
            "Token cache is unreadable (corrupt, tampered, or key mismatch). Re-authenticate to reset it.",
            e,
          );
        }
      }
    },

    async save(tokens: TokenSet): Promise<void> {
      const parsed = TokenSetSchema.parse(tokens);
      const key = await loadOrCreateKey();
      const container = encrypt(key, JSON.stringify(parsed));
      await writeAtomic(tokenCacheFile, JSON.stringify(container));
    },

    async clear(options: { includeKey?: boolean } = {}): Promise<void> {
      // Let key work already under way in this store finish first, so it can't
      // land after the deletion (a full cleanup must leave nothing behind).
      await Promise.allSettled([inflight.read, inflight.write].filter((p) => p !== null));
      const files = options.includeKey ? [tokenCacheFile, resolvedKeyFile] : [tokenCacheFile];
      for (const path of files) {
        try {
          await unlink(path);
        } catch (e) {
          if (isErrnoException(e) && e.code !== "ENOENT") throw e;
        }
      }
      if (options.includeKey) cached = null;
    },
  };
}
