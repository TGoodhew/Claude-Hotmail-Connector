/**
 * Cross-process coordination of sign-in and sign-out (#95).
 *
 * Several processes can share one token cache: Claude Desktop's extension,
 * Claude Code's server, and the `login` / `logout` commands. Within a process,
 * `createMicrosoftAuth` already orders saves against sign-out. Across processes
 * two counters in a small state file next to the cache do the same:
 *
 * - `signOuts` is bumped by every sign-out. An operation (a refresh, a browser
 *   sign-in) notes it when it starts and saves only if it hasn't changed, so a
 *   sign-out in any process wins over anything that started before it.
 * - `signIns` is bumped by every completed browser sign-in. A sign-in saves only
 *   if no other sign-in completed since it started, so of two pending browser
 *   flows the first to finish wins and the later one can't overwrite it.
 *
 * The check and the save (or the sign-out's bump and cleanup) run under an
 * exclusive lock file, so they can't interleave with another process's.
 */

import { randomBytes } from "node:crypto";
import { mkdir, open, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { AuthError } from "../util/errors.js";

export interface SharedAuthState {
  signOuts: number;
  signIns: number;
}

export interface AuthCoordinator {
  /** The shared counters now, without the lock: note them when an operation starts. */
  snapshot(): Promise<SharedAuthState>;
  /**
   * Run `fn` holding the lock, with the counters as they are now. If it returns
   * `next`, those counters are stored before the lock is released.
   */
  withLock<T>(
    fn: (state: SharedAuthState) => Promise<{ value: T; next?: SharedAuthState }>,
  ): Promise<T>;
}

const EMPTY: SharedAuthState = { signOuts: 0, signIns: 0 };

/** For a single process (and tests): the same contract, kept in memory. */
export function createMemoryCoordinator(): AuthCoordinator {
  let state = { ...EMPTY };
  let queue: Promise<unknown> = Promise.resolve();
  return {
    snapshot: async () => ({ ...state }),
    withLock(fn) {
      const run = queue.then(async () => {
        const { value, next } = await fn({ ...state });
        if (next) state = { ...next };
        return value;
      });
      queue = run.catch(() => undefined);
      return run;
    },
  };
}

export interface FileCoordinatorOptions {
  /** Give up waiting for the lock after this long (default 10s). */
  lockTimeoutMs?: number;
  /**
   * A lock is stale, and removed, when the process that took it has exited. A
   * running holder keeps it however long it takes (it may just be paused, e.g.
   * across sleep). Only a lock that never got a process id (its taker died
   * creating it) goes by age: stale once it is this old (default 2 minutes).
   */
  staleLockMs?: number;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
}

const STATE_FILE = "auth-state.json";
const LOCK_FILE = "auth.lock";
/** Guards removing a stale lock, so only one waiter does it at a time. */
const REAP_FILE = "auth.lock.reap";
/** A reaper holds its guard for milliseconds; an id-less one this old was left by a crash. */
const REAP_STALE_MS = 5_000;

function isErrno(e: unknown, code: string): boolean {
  return typeof e === "object" && e !== null && (e as NodeJS.ErrnoException).code === code;
}

/** Whether a process with this id is running (EPERM: it is, but isn't ours to signal). */
function isRunning(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    return isErrno(e, "EPERM");
  }
}

/** A lock or guard file as found: what it holds and when it was written. */
interface Holder {
  content: string;
  mtimeMs: number;
}

async function inspect(path: string): Promise<Holder | null> {
  try {
    const { mtimeMs } = await stat(path);
    return { content: await readFile(path, "utf8"), mtimeMs };
  } catch (e) {
    if (isErrno(e, "ENOENT")) return null;
    throw e;
  }
}

/**
 * Whether a lock or guard's holder is gone. The holder's process id leads the
 * file (`<pid>-<random>`): it is gone once that process has exited. A file
 * without one was left by a taker that died creating it, once it is older than
 * `maxAgeMs` (until then it may still be being written).
 */
function holderGone(h: Holder, nowMs: number, maxAgeMs: number): boolean {
  const pid = Number(h.content.split("-")[0]);
  if (!Number.isSafeInteger(pid) || pid <= 0) return nowMs - h.mtimeMs > maxAgeMs;
  // This process's own holders are alive: it doesn't keep locks it isn't using.
  return pid !== process.pid && !isRunning(pid);
}

/**
 * Remove `path` only if it still holds `expected`. It is moved aside first
 * (atomic), so what is checked is what was taken. Anything else, i.e. a file
 * that replaced the expected one since, is put back unless yet another has
 * appeared, so this can't delete a replacement the way a check-then-`rm` could.
 * Exported for tests.
 */
export async function removeIfUnchanged(path: string, expected: string): Promise<void> {
  const aside = `${path}.${process.pid}.${randomBytes(4).toString("hex")}.old`;
  try {
    await rename(path, aside);
  } catch (e) {
    // Gone already, or (Windows) open in its writer for a moment: leave it.
    if (isErrno(e, "ENOENT") || isErrno(e, "EPERM") || isErrno(e, "EBUSY")) return;
    throw e;
  }
  try {
    const content = await readFile(aside, "utf8");
    if (content !== expected) {
      try {
        const fh = await open(path, "wx");
        try {
          await fh.writeFile(content);
        } finally {
          await fh.close();
        }
      } catch (e) {
        if (!isErrno(e, "EEXIST")) throw e;
      }
    }
  } finally {
    await rm(aside, { force: true });
  }
}

function parseState(raw: string): SharedAuthState {
  try {
    const v = JSON.parse(raw) as Partial<SharedAuthState>;
    const n = (x: unknown) => (Number.isSafeInteger(x) && (x as number) >= 0 ? (x as number) : 0);
    return { signOuts: n(v.signOuts), signIns: n(v.signIns) };
  } catch {
    return { ...EMPTY }; // unreadable: start again (the next update rewrites it)
  }
}

/** Shared between the processes using `dir` (the token cache directory). */
export function createFileCoordinator(
  dir: string,
  options: FileCoordinatorOptions = {},
): AuthCoordinator {
  const statePath = join(dir, STATE_FILE);
  const lockPath = join(dir, LOCK_FILE);
  const reapPath = join(dir, REAP_FILE);
  const timeoutMs = options.lockTimeoutMs ?? 10_000;
  const staleMs = options.staleLockMs ?? 120_000;
  const now = options.now ?? Date.now;
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((r) => setTimeout(r, ms)));
  // One lock attempt at a time within this process, so local callers queue here
  // instead of polling the file against each other.
  let local: Promise<unknown> = Promise.resolve();

  async function read(): Promise<SharedAuthState> {
    try {
      return parseState(await readFile(statePath, "utf8"));
    } catch (e) {
      if (isErrno(e, "ENOENT")) return { ...EMPTY };
      throw e;
    }
  }

  async function write(state: SharedAuthState): Promise<void> {
    const tmp = `${statePath}.${process.pid}.${randomBytes(4).toString("hex")}.tmp`;
    await writeFile(tmp, JSON.stringify({ v: 1, ...state }), { mode: 0o600 });
    await rename(tmp, statePath); // atomic replace: readers see the old or the new file
  }

  /** Create `path` exclusively holding a new token, or return null if it exists. */
  async function take(path: string): Promise<string | null> {
    const token = `${process.pid}-${randomBytes(8).toString("hex")}`;
    try {
      const fh = await open(path, "wx");
      try {
        await fh.writeFile(token);
      } finally {
        await fh.close();
      }
      return token;
    } catch (e) {
      if (isErrno(e, "EEXIST")) return null;
      throw e;
    }
  }

  /** Take the lock, returning the token that proves this holder owns it. */
  async function acquire(): Promise<string> {
    await mkdir(dir, { recursive: true });
    const deadline = now() + timeoutMs;
    for (;;) {
      const token = await take(lockPath);
      if (token) return token;
      const held = await inspect(lockPath);
      if (!held) continue; // released between our attempts: try again at once
      if (holderGone(held, now(), staleMs) && (await reap(held.content))) continue;
      if (now() >= deadline) {
        throw new AuthError(
          "Another Hotmail Connector process is busy with the sign-in. Try again in a moment.",
        );
      }
      await sleep(20 + Math.floor(Math.random() * 30));
    }
  }

  /**
   * Remove the stale lock that held `stale`; false if another waiter is
   * reaping. One waiter reaps at a time (an exclusive guard file), it checks
   * again under the guard, and both the lock and a guard left by a crashed
   * reaper are removed only if they still hold what was judged stale (see
   * {@link removeIfUnchanged}), so no waiter can remove a lock or guard that
   * another has taken since it looked.
   */
  async function reap(stale: string): Promise<boolean> {
    const guard = await take(reapPath);
    if (!guard) {
      const reaper = await inspect(reapPath);
      if (reaper && holderGone(reaper, now(), REAP_STALE_MS)) {
        await removeIfUnchanged(reapPath, reaper.content);
      }
      return false;
    }
    try {
      const held = await inspect(lockPath);
      if (held?.content === stale && holderGone(held, now(), staleMs)) {
        await removeIfUnchanged(lockPath, stale);
      }
      return true;
    } finally {
      await removeIfUnchanged(reapPath, guard);
    }
  }

  /** Release the lock only if it is still ours (it may have been broken as stale). */
  async function release(token: string): Promise<void> {
    try {
      if ((await readFile(lockPath, "utf8")) === token) await rm(lockPath, { force: true });
    } catch (e) {
      if (!isErrno(e, "ENOENT")) throw e;
    }
  }

  return {
    snapshot: read,
    withLock(fn) {
      const run = local.then(async () => {
        const token = await acquire();
        try {
          const { value, next } = await fn(await read());
          if (next) await write(next);
          return value;
        } finally {
          await release(token);
        }
      });
      local = run.catch(() => undefined);
      return run;
    },
  };
}
