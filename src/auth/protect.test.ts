import { randomBytes } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import { createDpapiProtector, defaultKeyProtector, type PowerShellRunner } from "./protect.js";

describe("createDpapiProtector (injected runner)", () => {
  it("sends the secret on stdin (never in the command) and decodes the base64 reply", async () => {
    const run = vi.fn<PowerShellRunner>(async () => Buffer.from("BLOB").toString("base64"));
    const key = Buffer.from("super-secret-key-material-123456");
    const out = await createDpapiProtector(run).protect(key);

    expect(out.toString()).toBe("BLOB");
    const [command, stdin] = run.mock.calls[0]!;
    expect(command).toContain("ProtectedData]::Protect(");
    expect(command).toContain("CurrentUser");
    expect(stdin).toBe(key.toString("base64"));
    expect(command).not.toContain(key.toString("base64"));
  });

  it("uses Unprotect for unwrapping", async () => {
    const run = vi.fn<PowerShellRunner>(async () => Buffer.from("KEY").toString("base64"));
    await createDpapiProtector(run).unprotect(Buffer.from("BLOB"));
    expect(run.mock.calls[0]![0]).toContain("ProtectedData]::Unprotect(");
  });

  it("rejects an empty reply rather than returning an empty key", async () => {
    const run = vi.fn<PowerShellRunner>(async () => "");
    await expect(createDpapiProtector(run).protect(Buffer.from("x"))).rejects.toThrow(
      /returned no data/,
    );
  });
});

describe("defaultKeyProtector", () => {
  it("is DPAPI on Windows and absent elsewhere (for now)", () => {
    expect(defaultKeyProtector("win32")?.name).toBe("dpapi");
    expect(defaultKeyProtector("darwin")).toBeNull();
    expect(defaultKeyProtector("linux")).toBeNull();
  });
});

describe.runIf(process.platform === "win32")("real Windows DPAPI", () => {
  it("round-trips a key, and the wrapped blob differs from the key", async () => {
    const dpapi = createDpapiProtector();
    const key = randomBytes(32);
    const wrapped = await dpapi.protect(key);
    expect(wrapped.equals(key)).toBe(false);
    expect(wrapped.length).toBeGreaterThan(key.length);
    expect((await dpapi.unprotect(wrapped)).equals(key)).toBe(true);
  }, 30_000);
});
