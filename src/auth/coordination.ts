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
   * A lock older than this is taken to belong to a process that died while
   * holding it, and is removed (default 30s; locks are held for milliseconds).
   */
  staleLockMs?: number;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
}

const STATE_FILE = "auth-state.json";
const LOCK_FILE = "auth.lock";

function isErrno(e: unknown, code: string): boolean {
  return typeof e === "object" && e !== null && (e as NodeJS.ErrnoException).code === code;
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
  const timeoutMs = options.lockTimeoutMs ?? 10_000;
  const staleMs = options.staleLockMs ?? 30_000;
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

  /** Take the lock, returning the token that proves this holder owns it. */
  async function acquire(): Promise<string> {
    await mkdir(dir, { recursive: true });
    const token = `${process.pid}-${randomBytes(8).toString("hex")}`;
    const deadline = now() + timeoutMs;
    for (;;) {
      try {
        const fh = await open(lockPath, "wx"); // exclusive create: fails if held
        try {
          await fh.writeFile(token);
        } finally {
          await fh.close();
        }
        return token;
      } catch (e) {
        if (!isErrno(e, "EEXIST")) throw e;
      }
      try {
        if (now() - (await stat(lockPath)).mtimeMs > staleMs) {
          // Its holder died mid-operation: break the lock and try again.
          await rm(lockPath, { force: true });
          continue;
        }
      } catch (e) {
        if (!isErrno(e, "ENOENT")) throw e;
        continue; // released between our attempts: try again at once
      }
      if (now() >= deadline) {
        throw new AuthError(
          "Another Hotmail Connector process is busy with the sign-in. Try again in a moment.",
        );
      }
      await sleep(20 + Math.floor(Math.random() * 30));
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
