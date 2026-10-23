/**
 * The optional log file (#178). Claude Desktop doesn't keep an extension's
 * stderr, so with `LOG_FILE` on (the extension's "Write a log file" setting)
 * every log line also goes to a file in the cache directory.
 *
 * The lines are the logger's, already redacted (log.ts), with a timestamp and
 * the process id in front. The file never exceeds `maxBytes` (2 MB): when the
 * next line would take it over, the older half is cut off, at a line
 * boundary, and the newer half kept (owner decision on #178). Sign-out leaves
 * the file in place (also #178): it holds no tokens or account names.
 *
 * Processes sharing the cache directory (a source checkout's CLI and Claude
 * Code, say) share the file, so each write (size check, cut, append) runs
 * under a lock file beside it. A line that can't get the lock within a few
 * milliseconds is dropped rather than wait or pass the cap.
 *
 * Writes are synchronous, so a line is on disk before the process can exit.
 * If the file can't be written, the sink says so once on stderr and stops,
 * rather than failing the request that logged.
 */

import {
  appendFileSync,
  closeSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { dirname } from "node:path";
import { isErrno, isTransientFsError } from "./fs-retry.js";

/** The most the file may hold. */
export const DEFAULT_MAX_LOG_BYTES = 2 * 1024 * 1024;

/** File name of the log in the cache directory. */
export const LOG_FILE_NAME = "connector.log";

/** Tries for the lock, a millisecond apart, before a line is dropped. */
const LOCK_TRIES = 20;
/** A lock older than this was left by a process that stopped mid-write. */
const LOCK_STALE_MS = 5_000;
/** Brief Windows sharing errors in a row (dropping those lines) before giving up. */
const MAX_TRANSIENT_FAILURES = 50;

export interface LogFileOptions {
  maxBytes?: number;
  /** Injectable clock for tests. */
  now?: () => Date;
  /** Where the one "can't write the log file" warning goes; defaults to stderr. */
  warn?: (message: string) => void;
}

/** Block for `ms` milliseconds (the sink is synchronous by design). */
function pause(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** A sink that appends each log line to `path`, never letting it pass `maxBytes`. */
export function createLogFileSink(
  path: string,
  options: LogFileOptions = {},
): (line: string) => void {
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_LOG_BYTES;
  const now = options.now ?? (() => new Date());
  const warn = options.warn ?? ((m: string) => process.stderr.write(m + "\n"));
  const lockPath = `${path}.lock`;
  let failed = false;
  let transientFailures = 0;

  function currentSize(): number {
    try {
      return statSync(path).size;
    } catch {
      return 0;
    }
  }

  function stop(e: unknown): void {
    failed = true;
    const reason = e instanceof Error ? e.message : String(e);
    warn(
      `Hotmail Connector: can't write the log file ${path} (${reason}); logging to stderr only.`,
    );
  }

  /** Take the lock file, breaking a stale one; false if another process holds it. */
  function lock(): boolean {
    for (let attempt = 1; attempt <= LOCK_TRIES; attempt++) {
      try {
        closeSync(openSync(lockPath, "wx", 0o600));
        return true;
      } catch (e) {
        if (!isErrno(e, "EEXIST") && !isTransientFsError(e)) throw e;
        if (isErrno(e, "EEXIST")) {
          try {
            if (Date.now() - statSync(lockPath).mtimeMs > LOCK_STALE_MS) {
              unlinkSync(lockPath);
              continue;
            }
          } catch {
            // Released meanwhile: try again.
          }
        }
        pause(1);
      }
    }
    return false;
  }

  function unlock(): void {
    try {
      unlinkSync(lockPath);
    } catch {
      // Already gone (broken as stale): nothing to release.
    }
  }

  /** Keep only the newer half of the file, starting at a whole line. */
  function cutOlderHalf(): void {
    const data = readFileSync(path);
    let start = data.length - Math.floor(maxBytes / 2);
    if (start <= 0) return;
    const newline = data.indexOf(0x0a, start - 1);
    start = newline === -1 ? data.length : newline + 1;
    const tmp = `${path}.${process.pid}.tmp`;
    try {
      writeFileSync(tmp, data.subarray(start), { mode: 0o600 });
      renameSync(tmp, path);
    } catch (e) {
      try {
        unlinkSync(tmp);
      } catch {
        // Never written.
      }
      throw e;
    }
  }

  function append(entry: string, bytes: number): void {
    if (currentSize() + bytes > maxBytes) {
      cutOlderHalf();
      if (currentSize() + bytes > maxBytes) return;
    }
    appendFileSync(path, entry, { mode: 0o600 });
  }

  try {
    mkdirSync(dirname(path), { recursive: true });
  } catch (e) {
    stop(e);
  }

  return (line: string) => {
    if (failed) return;
    const entry = `${now().toISOString()} [${process.pid}] ${line}\n`;
    const bytes = Buffer.byteLength(entry);
    if (bytes > maxBytes) return; // A single line can't fit: drop it rather than pass the cap.
    try {
      if (!lock()) return; // Busy with another process's line: drop this one.
      try {
        append(entry, bytes);
      } finally {
        unlock();
      }
      transientFailures = 0;
    } catch (e) {
      // Windows refuses for a moment while another process has the file open:
      // drop this line and try again on the next, unless it keeps happening.
      if (isTransientFsError(e) && ++transientFailures < MAX_TRANSIENT_FAILURES) return;
      stop(e);
    }
  };
}
