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
 * Writes are synchronous appends, so a line is on disk before the process can
 * exit. If the file can't be written, the sink says so once on stderr and
 * stops, rather than failing the request that logged.
 */

import {
  appendFileSync,
  mkdirSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { dirname } from "node:path";

/** The most the file may hold. */
export const DEFAULT_MAX_LOG_BYTES = 2 * 1024 * 1024;

/** File name of the log in the cache directory. */
export const LOG_FILE_NAME = "connector.log";

export interface LogFileOptions {
  maxBytes?: number;
  /** Injectable clock for tests. */
  now?: () => Date;
  /** Where the one "can't write the log file" warning goes; defaults to stderr. */
  warn?: (message: string) => void;
}

/** A sink that appends each log line to `path`, never letting it pass `maxBytes`. */
export function createLogFileSink(
  path: string,
  options: LogFileOptions = {},
): (line: string) => void {
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_LOG_BYTES;
  const now = options.now ?? (() => new Date());
  const warn = options.warn ?? ((m: string) => process.stderr.write(m + "\n"));
  let failed = false;

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
        // Already gone.
      }
      throw e;
    }
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
    // Read the size each time: another process (the CLI, Claude Code) may share the file.
    if (currentSize() + bytes > maxBytes) {
      try {
        cutOlderHalf();
      } catch {
        // Windows can refuse while another process has the file open. Drop this
        // line rather than pass the cap, and try again on the next one.
        return;
      }
      if (currentSize() + bytes > maxBytes) return;
    }
    try {
      appendFileSync(path, entry, { mode: 0o600 });
    } catch (e) {
      stop(e);
    }
  };
}
