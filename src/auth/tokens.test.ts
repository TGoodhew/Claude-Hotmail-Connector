import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AuthError } from "../util/errors.js";
import { createDpapiProtector, type KeyProtector } from "./protect.js";
import { createFileTokenStore, type TokenSet } from "./tokens.js";

let dir: string;
let file: string;

const sample: TokenSet = {
  accessToken: "ACCESS-TOKEN-abc123-secretvalue",
  refreshToken: "REFRESH-TOKEN-xyz789-secretvalue",
  expiresAt: 1_800_000_000_000,
  scopes: ["Mail.Read", "Calendars.ReadWrite"],
  account: "tony_goodhew@hotmail.com",
  tokenType: "Bearer",
};

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "cht-tokens-"));
  file = join(dir, "token-cache.enc");
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("createFileTokenStore", () => {
  it("returns null when nothing is stored", async () => {
    const store = createFileTokenStore(file);
    expect(await store.load()).toBeNull();
  });

  it("round-trips a token set through encryption", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    expect(await store.load()).toEqual(sample);
  });

  it("never writes plaintext tokens to disk", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    const onDisk = readFileSync(file, "utf8");
    expect(onDisk).not.toContain(sample.accessToken);
    expect(onDisk).not.toContain(sample.refreshToken);
    expect(onDisk).not.toContain("secretvalue");
    // The container is JSON with iv/tag/data fields.
    const parsed = JSON.parse(onDisk);
    expect(parsed).toHaveProperty("iv");
    expect(parsed).toHaveProperty("tag");
    expect(parsed).toHaveProperty("data");
  });

  it("detects tampering via the GCM auth tag", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    const container = JSON.parse(readFileSync(file, "utf8"));
    // Flip a byte in the ciphertext.
    const bytes = Buffer.from(container.data, "base64");
    bytes[0] = bytes[0]! ^ 0xff;
    container.data = bytes.toString("base64");
    writeFileSync(file, JSON.stringify(container));

    await expect(store.load()).rejects.toBeInstanceOf(AuthError);
  });

  it("throws AuthError when the container is not valid JSON", async () => {
    writeFileSync(file, "not-json");
    const store = createFileTokenStore(file);
    await expect(store.load()).rejects.toBeInstanceOf(AuthError);
  });

  it("clear() removes the cache file", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    expect(existsSync(file)).toBe(true);
    await store.clear();
    expect(existsSync(file)).toBe(false);
    expect(await store.load()).toBeNull();
  });

  it("clear({ includeKey }) also removes the key file, and a new sign-in starts afresh", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    const key = join(dir, "token-cache.key");
    expect(existsSync(key)).toBe(true);
    await store.clear({ includeKey: true });
    expect(existsSync(file)).toBe(false);
    expect(existsSync(key)).toBe(false);
    await store.save(sample); // same store, no stale cached key
    await expect(createFileTokenStore(file).load()).resolves.toEqual(sample);
  });

  it("clear() is a no-op when nothing is stored", async () => {
    const store = createFileTokenStore(file);
    await expect(store.clear()).resolves.toBeUndefined();
  });

  it.runIf(process.platform !== "win32")("stores the key with owner-only perms", async () => {
    const store = createFileTokenStore(file);
    await store.save(sample);
    const keyMode = statSync(join(dir, "token-cache.key")).mode & 0o777;
    expect(keyMode & 0o077).toBe(0); // no group/other permissions
  });
});

describe("createFileTokenStore with an OS key protector", () => {
  const MARK = Buffer.from("wrapped:");
  /** Reversible stand-in for DPAPI that counts calls. */
  function fakeProtector(overrides: Partial<KeyProtector> = {}) {
    return {
      name: "fake",
      protect: vi.fn(async (d: Buffer) => Buffer.concat([MARK, Buffer.from(d).reverse()])),
      unprotect: vi.fn(async (d: Buffer) => Buffer.from(d.subarray(MARK.length)).reverse()),
      ...overrides,
    };
  }
  const keyFile = () => join(dir, "token-cache.key");

  it("writes only the wrapped key, and round-trips across store instances", async () => {
    await createFileTokenStore(file, { protector: fakeProtector() }).save(sample);
    const stored = JSON.parse(readFileSync(keyFile(), "utf8"));
    expect(stored).toMatchObject({ v: 2, protection: "fake" });

    const unwrapped = Buffer.from(stored.key, "base64").subarray(MARK.length).reverse();
    expect(readFileSync(keyFile(), "utf8")).not.toContain(unwrapped.toString("base64"));

    expect(await createFileTokenStore(file, { protector: fakeProtector() }).load()).toEqual(sample);
  });

  it("upgrades a legacy plain key in place on the next save, keeping the existing sign-in", async () => {
    await createFileTokenStore(file).save(sample); // legacy: plain base64 key
    expect(readFileSync(keyFile(), "utf8")).not.toContain("{");

    const protector = fakeProtector();
    const store = createFileTokenStore(file, { protector });
    // Loads are read-only: the legacy key still decrypts the cache, untouched.
    expect(await store.load()).toEqual(sample);
    expect(readFileSync(keyFile(), "utf8")).not.toContain("{");
    expect(protector.protect).not.toHaveBeenCalled();

    // The next save (e.g. a token refresh) wraps it.
    await store.save(sample);
    expect(JSON.parse(readFileSync(keyFile(), "utf8")).protection).toBe("fake");
    expect(protector.protect).toHaveBeenCalledOnce();
    expect(await createFileTokenStore(file, { protector }).load()).toEqual(sample);
  });

  it("a load paused across a full cleanup can't put a key back on disk", async () => {
    await createFileTokenStore(file, { protector: fakeProtector() }).save(sample);
    let release!: () => void;
    const paused = new Promise<void>((r) => (release = r));
    let entered!: () => void;
    const inUnwrap = new Promise<void>((r) => (entered = r));
    const slow = fakeProtector({
      unprotect: vi.fn(async (d: Buffer) => {
        entered();
        await paused;
        return Buffer.from(d.subarray(MARK.length)).reverse();
      }),
    });
    const store = createFileTokenStore(file, { protector: slow });
    const loading = store.load().catch((e: unknown) => e);
    await inUnwrap;
    const clearing = store.clear({ includeKey: true }); // waits for the load's key work
    // Grace period: a cleanup that did NOT wait would delete both files now, and the
    // resumed load would then recreate the key. (With the fix, nothing is deleted yet.)
    await new Promise((r) => setTimeout(r, 50));
    release();
    await clearing;
    await loading;
    expect(existsSync(file)).toBe(false);
    expect(existsSync(keyFile())).toBe(false);
    expect(await store.load()).toBeNull();
    expect(existsSync(keyFile())).toBe(false); // a load with no cache creates nothing
  });

  it("falls back to a plain key file when protection fails", async () => {
    const protector = fakeProtector({
      protect: vi.fn(async () => {
        throw new Error("DPAPI unavailable");
      }),
    });
    const store = createFileTokenStore(file, { protector });
    await store.save(sample);
    expect(readFileSync(keyFile(), "utf8")).not.toContain("{");
    expect(await store.load()).toEqual(sample);
  });

  it("resets the key when it can't be unwrapped (e.g. another user/machine)", async () => {
    await createFileTokenStore(file, { protector: fakeProtector() }).save(sample);
    const foreign = fakeProtector({
      unprotect: vi.fn(async () => {
        throw new Error("Key not valid for use in specified state.");
      }),
    });
    // The old cache can no longer be decrypted: reported as unreadable (=> re-authenticate).
    await expect(createFileTokenStore(file, { protector: foreign }).load()).rejects.toBeInstanceOf(
      AuthError,
    );
  });

  it("unwraps the key once per store, not on every load", async () => {
    await createFileTokenStore(file, { protector: fakeProtector() }).save(sample);
    const protector = fakeProtector();
    const store = createFileTokenStore(file, { protector });
    await store.load();
    await store.load();
    await store.save(sample);
    expect(protector.unprotect).toHaveBeenCalledOnce();
  });

  it.runIf(process.platform === "win32")(
    "works end to end with real Windows DPAPI",
    async () => {
      const dpapi = createDpapiProtector();
      await createFileTokenStore(file, { protector: dpapi }).save(sample);
      expect(JSON.parse(readFileSync(keyFile(), "utf8")).protection).toBe("dpapi");
      expect(await createFileTokenStore(file, { protector: dpapi }).load()).toEqual(sample);
    },
    30_000,
  );
});

describe("createFileTokenStore across processes (shared key file)", () => {
  /** A protector slow enough to widen the window between reading and writing the key. */
  const slowProtector = (): KeyProtector => ({
    name: "slow",
    protect: async (d) => {
      await new Promise((r) => setTimeout(r, 15));
      return Buffer.from(d).reverse();
    },
    unprotect: async (d) => Buffer.from(d).reverse(),
  });

  it("two stores racing to create the key end up sharing one key", async () => {
    // Like two processes (Claude Desktop + Claude Code) starting on an empty cache.
    const a = createFileTokenStore(file, { protector: slowProtector() });
    const b = createFileTokenStore(file, { protector: slowProtector() });
    await Promise.all([a.save(sample), b.save({ ...sample, account: "b@example.com" })]);
    // Whichever cache write landed last, a fresh process can read it, and so can a and b.
    const fresh = await createFileTokenStore(file, { protector: slowProtector() }).load();
    expect(fresh).not.toBeNull();
    await expect(a.load()).resolves.toEqual(fresh);
    await expect(b.load()).resolves.toEqual(fresh);
    expect(readdirSync(dir).filter((f) => f.includes(".tmp-"))).toEqual([]);
  });

  it("doesn't cache a key whose file was replaced while it was being unwrapped", async () => {
    const reversing: KeyProtector = {
      name: "slow",
      protect: async (d) => Buffer.from(d).reverse(),
      unprotect: async (d) => Buffer.from(d).reverse(),
    };
    await createFileTokenStore(file, { protector: reversing }).save(sample);

    // Pause the first unwrap; meanwhile another process resets the key and cache.
    let release!: () => void;
    const paused = new Promise<void>((r) => (release = r));
    let entered!: () => void;
    const inFirstUnwrap = new Promise<void>((r) => (entered = r));
    let calls = 0;
    const pausing: KeyProtector = {
      ...reversing,
      unprotect: async (d) => {
        if (calls++ === 0) {
          entered();
          await paused;
        }
        return reversing.unprotect(d);
      },
    };
    const a = createFileTokenStore(file, { protector: pausing });
    const loading = a.load();
    await inFirstUnwrap; // a is now holding the old key file's contents
    writeFileSync(join(dir, "token-cache.key"), "corrupt");
    const b = createFileTokenStore(file, { protector: reversing });
    await b.save({ ...sample, account: "after-reset@example.com" });
    release();
    // The paused load re-reads the replacement cache rather than calling it corrupt
    // (which would make auth clear the fresh tokens).
    await expect(loading).resolves.toMatchObject({ account: "after-reset@example.com" });

    // `a` must now hold b's key: what it saves is readable by a fresh process.
    await a.save(sample);
    await expect(createFileTokenStore(file, { protector: reversing }).load()).resolves.toEqual(
      sample,
    );
  });

  it("adopts a replacement key instead of resetting it when its own unwrap failed", async () => {
    const reversing: KeyProtector = {
      name: "slow",
      protect: async (d) => Buffer.from(d).reverse(),
      unprotect: async (d) => Buffer.from(d).reverse(),
    };
    await createFileTokenStore(file, { protector: reversing }).save(sample);

    // a's first unwrap pauses, then fails (as if the key were foreign); meanwhile
    // b resets the key and saves fresh tokens.
    let release!: () => void;
    const paused = new Promise<void>((r) => (release = r));
    let entered!: () => void;
    const inFirstUnwrap = new Promise<void>((r) => (entered = r));
    let calls = 0;
    const failingOnce: KeyProtector = {
      ...reversing,
      unprotect: async (d) => {
        if (calls++ === 0) {
          entered();
          await paused;
          throw new Error("Key not valid for use in specified state.");
        }
        return reversing.unprotect(d);
      },
    };
    const a = createFileTokenStore(file, { protector: failingOnce });
    const loading = a.load();
    await inFirstUnwrap;
    writeFileSync(join(dir, "token-cache.key"), "corrupt");
    const b = createFileTokenStore(file, { protector: reversing });
    await b.save({ ...sample, account: "fresh@example.com" });
    release();

    // a must not overwrite b's key: b's fresh tokens stay readable for everyone.
    await expect(loading).resolves.toMatchObject({ account: "fresh@example.com" });
    await expect(
      createFileTokenStore(file, { protector: reversing }).load(),
    ).resolves.toMatchObject({ account: "fresh@example.com" });
  });

  it("re-reads the key when another process replaced the key file", async () => {
    const a = createFileTokenStore(file);
    await a.save(sample);
    // Another process finds the key corrupt, regenerates it, and writes a new cache.
    writeFileSync(join(dir, "token-cache.key"), "corrupt");
    const b = createFileTokenStore(file);
    await b.save({ ...sample, account: "after-reset@example.com" });
    // `a` must not keep using its stale key, for reading or for writing.
    await expect(a.load()).resolves.toMatchObject({ account: "after-reset@example.com" });
    await a.save(sample);
    await expect(createFileTokenStore(file).load()).resolves.toEqual(sample);
  });
});
