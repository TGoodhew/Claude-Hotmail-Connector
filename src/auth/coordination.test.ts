import { existsSync, mkdtempSync, readFileSync, rmSync, utimesSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { AuthError } from "../util/errors.js";
import {
  createFileCoordinator,
  createMemoryCoordinator,
  type AuthCoordinator,
} from "./coordination.js";

const dirs: string[] = [];
function tempDir(): string {
  const d = mkdtempSync(join(tmpdir(), "hotmail-coord-"));
  dirs.push(d);
  return d;
}
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

/** Run two overlapping locked sections and record whether they ever overlapped. */
async function overlapCheck(a: AuthCoordinator, b: AuthCoordinator): Promise<boolean> {
  let active = 0;
  let overlapped = false;
  const section = (c: AuthCoordinator) =>
    c.withLock(async (s) => {
      active += 1;
      if (active > 1) overlapped = true;
      await new Promise((r) => setTimeout(r, 15));
      active -= 1;
      return { value: undefined, next: { ...s, signIns: s.signIns + 1 } };
    });
  await Promise.all([section(a), section(b), section(a), section(b)]);
  return overlapped;
}

describe("createMemoryCoordinator", () => {
  it("runs locked sections one at a time and keeps their updates", async () => {
    const c = createMemoryCoordinator();
    expect(await overlapCheck(c, c)).toBe(false);
    expect(await c.snapshot()).toEqual({ signOuts: 0, signIns: 4 });
  });
});

describe("createFileCoordinator (#95)", () => {
  it("excludes two processes from the locked section at once", async () => {
    const dir = tempDir();
    // Two coordinators on one directory stand in for two processes.
    expect(await overlapCheck(createFileCoordinator(dir), createFileCoordinator(dir))).toBe(false);
    expect(await createFileCoordinator(dir).snapshot()).toEqual({ signOuts: 0, signIns: 4 });
    expect(existsSync(join(dir, "auth.lock"))).toBe(false); // released
  });

  it("keeps the counters across processes, and starts from zero without a state file", async () => {
    const dir = tempDir();
    const a = createFileCoordinator(dir);
    expect(await a.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
    await a.withLock(async (s) => ({ value: 1, next: { ...s, signOuts: s.signOuts + 1 } }));
    expect(await createFileCoordinator(dir).snapshot()).toEqual({ signOuts: 1, signIns: 0 });
  });

  it("reads an unreadable or malformed state file as zeros, then rewrites it", async () => {
    const dir = tempDir();
    writeFileSync(join(dir, "auth-state.json"), "{not json");
    const c = createFileCoordinator(dir);
    expect(await c.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
    writeFileSync(join(dir, "auth-state.json"), JSON.stringify({ signOuts: -3, signIns: "x" }));
    expect(await c.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
    await c.withLock(async (s) => ({ value: 0, next: { ...s, signIns: 1 } }));
    expect(JSON.parse(readFileSync(join(dir, "auth-state.json"), "utf8"))).toMatchObject({
      signOuts: 0,
      signIns: 1,
    });
  });

  it("breaks a lock left by a process that died holding it", async () => {
    const dir = tempDir();
    const lock = join(dir, "auth.lock");
    writeFileSync(lock, "dead-process");
    const old = new Date(Date.now() - 60_000);
    utimesSync(lock, old, old);
    const value = await createFileCoordinator(dir, { staleLockMs: 30_000 }).withLock(async () => ({
      value: "ran",
    }));
    expect(value).toBe("ran");
    expect(existsSync(lock)).toBe(false);
  });

  it("gives up on a live lock after the timeout, with an AuthError", async () => {
    const dir = tempDir();
    writeFileSync(join(dir, "auth.lock"), "busy-process"); // fresh: not stale
    const c = createFileCoordinator(dir, { lockTimeoutMs: 60 });
    await expect(c.withLock(async () => ({ value: 1 }))).rejects.toBeInstanceOf(AuthError);
    expect(readFileSync(join(dir, "auth.lock"), "utf8")).toBe("busy-process"); // left alone
  });

  it("doesn't remove a lock that another process has taken since", async () => {
    const dir = tempDir();
    const lock = join(dir, "auth.lock");
    await createFileCoordinator(dir).withLock(async () => {
      // Our lock was broken as stale and someone else holds the file now.
      writeFileSync(lock, "someone-else");
      return { value: undefined };
    });
    expect(readFileSync(lock, "utf8")).toBe("someone-else");
  });

  it("releases the lock when the locked section throws, and stores nothing", async () => {
    const dir = tempDir();
    const c = createFileCoordinator(dir);
    await expect(
      c.withLock(async () => {
        throw new Error("boom");
      }),
    ).rejects.toThrow("boom");
    expect(existsSync(join(dir, "auth.lock"))).toBe(false);
    expect(await c.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
  });
});
