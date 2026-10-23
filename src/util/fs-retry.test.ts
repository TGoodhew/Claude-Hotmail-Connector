import { describe, expect, it, vi } from "vitest";
import { isErrno, isTransientFsError, retrying } from "./fs-retry.js";

const errno = (code: string) => Object.assign(new Error(code), { code });

describe("retrying (#142)", () => {
  it("retries Windows' transient sharing errors until the operation succeeds", async () => {
    const op = vi
      .fn<() => Promise<string>>()
      .mockRejectedValueOnce(errno("EPERM"))
      .mockRejectedValueOnce(errno("EBUSY"))
      .mockRejectedValueOnce(errno("EACCES"))
      .mockResolvedValue("done");
    await expect(retrying(op)).resolves.toBe("done");
    expect(op).toHaveBeenCalledTimes(4);
  });

  it("throws any other error at once", async () => {
    const op = vi.fn<() => Promise<void>>().mockRejectedValue(errno("ENOENT"));
    await expect(retrying(op)).rejects.toMatchObject({ code: "ENOENT" });
    expect(op).toHaveBeenCalledOnce();
  });

  it("gives up after about half a second of transient errors", async () => {
    const op = vi.fn<() => Promise<void>>().mockRejectedValue(errno("EPERM"));
    await expect(retrying(op)).rejects.toMatchObject({ code: "EPERM" });
    expect(op).toHaveBeenCalledTimes(50);
  });

  it("recognises error codes", () => {
    expect(isErrno(errno("EEXIST"), "EEXIST")).toBe(true);
    expect(isErrno(new Error("x"), "EEXIST")).toBe(false);
    expect(isErrno(null, "EEXIST")).toBe(false);
    expect(isTransientFsError(errno("EBUSY"))).toBe(true);
    expect(isTransientFsError(errno("ENOENT"))).toBe(false);
  });
});
