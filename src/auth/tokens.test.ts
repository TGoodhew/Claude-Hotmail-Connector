import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { AuthError } from "../util/errors.js";
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
