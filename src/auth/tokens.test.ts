import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync, statSync } from "node:fs";
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

  it("upgrades a legacy plain key in place, keeping the existing sign-in", async () => {
    await createFileTokenStore(file).save(sample); // legacy: plain base64 key
    expect(readFileSync(keyFile(), "utf8")).not.toContain("{");

    const protector = fakeProtector();
    expect(await createFileTokenStore(file, { protector }).load()).toEqual(sample);
    expect(JSON.parse(readFileSync(keyFile(), "utf8")).protection).toBe("fake");
    expect(protector.protect).toHaveBeenCalledOnce();
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
