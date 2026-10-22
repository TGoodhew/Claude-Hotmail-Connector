import { spawn, spawnSync } from "node:child_process";
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { AuthError } from "../util/errors.js";
import {
  createFileCoordinator,
  createMemoryCoordinator,
  osLockName,
  removeIfUnchanged,
  type AuthCoordinator,
  type FileCoordinatorOptions,
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

/** Run overlapping locked sections and record whether any two ever overlapped. */
async function overlapCheck(...cs: AuthCoordinator[]): Promise<boolean> {
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
  await Promise.all([...cs, ...cs].map(section));
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
    expect(existsSync(join(dir, "auth.lock"))).toBe(false); // released (or never a file)
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

  it("releases the lock when the locked section throws, and stores nothing", async () => {
    const dir = tempDir();
    const c = createFileCoordinator(dir);
    await expect(
      c.withLock(async () => {
        throw new Error("boom");
      }),
    ).rejects.toThrow("boom");
    expect(await c.withLock(async () => ({ value: "free again" }))).toBe("free again");
    expect(await c.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
  });
});

describe("osLockName (Copilot review on #126)", () => {
  it("names a per-directory pipe on Windows and an abstract socket on Linux, nothing elsewhere", () => {
    const win = osLockName("C:\\Users\\A\\cache", "win32")!;
    expect(win).toMatch(/^\\\\\.\\pipe\\claude-hotmail-connector-auth-[0-9a-f]{32}$/);
    expect(osLockName("c:\\users\\a\\CACHE", "win32")).toBe(win); // paths ignore case there
    expect(osLockName("C:\\Users\\B\\cache", "win32")).not.toBe(win);
    const linux = osLockName("/home/a/cache", "linux")!;
    expect(linux.charCodeAt(0)).toBe(0); // abstract: no file to leave behind
    expect(linux.slice(1)).toMatch(/^claude-hotmail-connector-auth-[0-9a-f]{32}$/);
    expect(osLockName("/Users/a/cache", "darwin")).toBeNull();
  });
});

const name = osLockName(tmpdir());
describe.skipIf(name === null)("the OS-held lock (Windows, Linux; Copilot review on #126)", () => {
  /** Another process holding the lock for `dir` until it is killed. */
  async function holderProcess(dir: string) {
    const script =
      'require("net").createServer().listen(process.argv[1], () => console.log("held"));' +
      "setInterval(() => {}, 1000);";
    const child = spawn(process.execPath, ["-e", script, osLockName(dir)!], {
      stdio: ["ignore", "pipe", "inherit"],
    });
    await new Promise((r) => child.stdout.once("data", r));
    return child;
  }

  it("uses no lock file", async () => {
    const dir = tempDir();
    await createFileCoordinator(dir).withLock(async () => {
      expect(readdirSync(dir)).toEqual([]);
      return { value: undefined };
    });
  });

  it("waits for another process holding it, and has it once that process dies", async () => {
    const dir = tempDir();
    const child = await holderProcess(dir);
    try {
      const c = createFileCoordinator(dir, { lockTimeoutMs: 100 });
      await expect(c.withLock(async () => ({ value: 1 }))).rejects.toBeInstanceOf(AuthError);
    } finally {
      child.kill("SIGKILL"); // a crash: nothing is cleaned up
      await new Promise((r) => child.once("exit", r));
    }
    // The OS let go of it with the process.
    const c = createFileCoordinator(dir, { lockTimeoutMs: 2_000 });
    expect(await c.withLock(async () => ({ value: "ran" }))).toBe("ran");
  });

  it("excludes many waiters at once", async () => {
    const dir = tempDir();
    const cs = Array.from({ length: 4 }, () => createFileCoordinator(dir));
    expect(await overlapCheck(...cs)).toBe(false);
  });
});

// The lock file is used only where there's no OS-held lock (macOS). Windows
// renames through file handles, so two waiters can both "move aside" one file
// there, which is one reason Windows uses the OS-held lock instead: these
// tests of the file protocol run elsewhere.
describe.skipIf(process.platform === "win32")("the lock file (#95)", () => {
  const file = (dir: string, o: FileCoordinatorOptions = {}) =>
    createFileCoordinator(dir, { ...o, mechanism: "file" });

  it("excludes two processes from the locked section at once", async () => {
    const dir = tempDir();
    expect(await overlapCheck(file(dir), file(dir))).toBe(false);
    expect(existsSync(join(dir, "auth.lock"))).toBe(false); // released
  });

  it("breaks a lock left by a process that died holding it", async () => {
    const dir = tempDir();
    const lock = join(dir, "auth.lock");
    writeFileSync(lock, "dead-process");
    const old = new Date(Date.now() - 60_000);
    utimesSync(lock, old, old);
    const value = await file(dir, { staleLockMs: 30_000 }).withLock(async () => ({
      value: "ran",
    }));
    expect(value).toBe("ran");
    expect(existsSync(lock)).toBe(false);
  });

  it("gives up on a live lock after the timeout, with an AuthError", async () => {
    const dir = tempDir();
    writeFileSync(join(dir, "auth.lock"), "busy-process"); // fresh: not stale
    const c = file(dir, { lockTimeoutMs: 60 });
    await expect(c.withLock(async () => ({ value: 1 }))).rejects.toBeInstanceOf(AuthError);
    expect(readFileSync(join(dir, "auth.lock"), "utf8")).toBe("busy-process"); // left alone
  });

  it("doesn't remove a lock that another process has taken since", async () => {
    const dir = tempDir();
    const lock = join(dir, "auth.lock");
    await file(dir).withLock(async () => {
      // Our lock was broken as stale and someone else holds the file now.
      writeFileSync(lock, "someone-else");
      return { value: undefined };
    });
    expect(readFileSync(lock, "utf8")).toBe("someone-else");
  });

  it("breaks a fresh lock at once when the process that took it has exited", async () => {
    const dir = tempDir();
    const exited = spawnSync(process.execPath, ["-e", ""]).pid!; // a process id no longer running
    writeFileSync(join(dir, "auth.lock"), `${exited}-deadbeef`);
    const c = file(dir, { lockTimeoutMs: 2_000 });
    expect(await c.withLock(async () => ({ value: "ran" }))).toBe("ran");
  });

  it("never lets two waiters reaping the same stale lock both hold it (Copilot review on #126)", async () => {
    const dir = tempDir();
    const exited = spawnSync(process.execPath, ["-e", ""]).pid!;
    writeFileSync(join(dir, "auth.lock"), `${exited}-deadbeef`);
    // Several "processes" all find the dead holder's lock at once.
    const cs = Array.from({ length: 4 }, () => file(dir));
    expect(await overlapCheck(...cs)).toBe(false);
    expect(existsSync(join(dir, "auth.lock.reap"))).toBe(false);
  });

  it("doesn't break a lock whose holder is still running, however long it waits", async () => {
    const dir = tempDir();
    writeFileSync(join(dir, "auth.lock"), `${process.pid}-other-instance`); // this process: alive
    const c = file(dir, { lockTimeoutMs: 80 });
    await expect(c.withLock(async () => ({ value: 1 }))).rejects.toBeInstanceOf(AuthError);
    expect(readFileSync(join(dir, "auth.lock"), "utf8")).toBe(`${process.pid}-other-instance`);
  });

  it("doesn't break an old lock whose holder is still running, e.g. paused across sleep (Copilot review on #126)", async () => {
    const dir = tempDir();
    const lock = join(dir, "auth.lock");
    writeFileSync(lock, `${process.ppid}-paused-holder`); // another process, still running
    const old = new Date(Date.now() - 60 * 60_000);
    utimesSync(lock, old, old);
    const c = file(dir, { lockTimeoutMs: 80, staleLockMs: 1_000 });
    await expect(c.withLock(async () => ({ value: 1 }))).rejects.toBeInstanceOf(AuthError);
    expect(readFileSync(lock, "utf8")).toBe(`${process.ppid}-paused-holder`);
  });

  it("recovers when a reaper also died holding its guard, without two waiters reaping at once (Copilot review on #126)", async () => {
    const dir = tempDir();
    const exited = spawnSync(process.execPath, ["-e", ""]).pid!;
    writeFileSync(join(dir, "auth.lock"), `${exited}-holder`);
    writeFileSync(join(dir, "auth.lock.reap"), `${exited}-reaper`);
    const cs = Array.from({ length: 4 }, () => file(dir));
    expect(await overlapCheck(...cs)).toBe(false);
    // Lock, guard and set-aside files are all gone; only the counters remain.
    expect(readdirSync(dir)).toEqual(["auth-state.json"]);
  });
});

describe("removeIfUnchanged (Copilot review on #126)", () => {
  it("removes the file when it still holds what was judged stale", async () => {
    const dir = tempDir();
    const f = join(dir, "auth.lock");
    writeFileSync(f, "stale");
    await removeIfUnchanged(f, "stale");
    expect(readdirSync(dir)).toEqual([]);
  });

  it("puts back a file that replaced the stale one since", async () => {
    const dir = tempDir();
    const f = join(dir, "auth.lock");
    writeFileSync(f, "new-holder");
    await removeIfUnchanged(f, "stale");
    expect(readFileSync(f, "utf8")).toBe("new-holder");
    expect(readdirSync(dir)).toEqual(["auth.lock"]);
  });

  it("does nothing when the file is already gone", async () => {
    const dir = tempDir();
    await removeIfUnchanged(join(dir, "auth.lock"), "stale");
    expect(readdirSync(dir)).toEqual([]);
  });
});
