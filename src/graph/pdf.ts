/**
 * Reading a PDF attachment's text in a worker thread (#156). The worker (see
 * src/pdf-worker.ts) gets a heap cap and a time limit, and is terminated when
 * either is passed, so a crafted PDF can't stall or exhaust the server. Its
 * stdout is drained and discarded: the server's own stdout is the MCP channel.
 */

import { join } from "node:path";
import { Worker } from "node:worker_threads";
import type { PdfExtract } from "./pdf-text.js";

/** Limits for one PDF: pages read, time allowed, and the worker's heap. */
export const PDF_LIMITS = { maxPages: 30, timeoutMs: 10_000, maxHeapMb: 256 } as const;

/** What reading a PDF found, including the worker running out of time or memory. */
export type PdfResult = PdfExtract | { status: "too-slow" } | { status: "too-complex" };

export interface ReadPdfOptions {
  /** The worker script; the built one next to the server by default. */
  workerFile?: string | URL;
  maxPages?: number;
  timeoutMs?: number;
  maxHeapMb?: number;
}

/**
 * Where the built worker is. The Claude Desktop extension's server is one
 * CommonJS file with pdf-worker.cjs beside it; a source checkout's dist/ is
 * ESM, with pdf-worker.js beside index.js.
 */
export function defaultWorkerFile(): string | URL {
  return typeof __dirname === "string"
    ? join(__dirname, "pdf-worker.cjs")
    : new URL("./pdf-worker.js", import.meta.url);
}

const STATUSES = new Set(["ok", "encrypted", "no-text", "unreadable"]);

function isExtract(m: unknown): m is PdfExtract {
  if (typeof m !== "object" || m === null) return false;
  const { status, text, pagesRead, totalPages } = m as Record<string, unknown>;
  if (!STATUSES.has(status as string)) return false;
  return (
    status !== "ok" ||
    (typeof text === "string" && typeof pagesRead === "number" && typeof totalPages === "number")
  );
}

/** Read the text of a PDF in a worker. Never throws: every failure is a status. */
export function readPdfText(data: Uint8Array, options: ReadPdfOptions = {}): Promise<PdfResult> {
  return new Promise((resolve) => {
    let worker: Worker;
    try {
      worker = new Worker(options.workerFile ?? defaultWorkerFile(), {
        workerData: { data, maxPages: options.maxPages ?? PDF_LIMITS.maxPages },
        resourceLimits: {
          maxOldGenerationSizeMb: options.maxHeapMb ?? PDF_LIMITS.maxHeapMb,
          maxYoungGenerationSizeMb: 32,
        },
        // Captured rather than inherited: nothing the worker prints may reach
        // the MCP channel. Drained so it can't block on a full pipe.
        stdout: true,
        stderr: true,
      });
    } catch {
      resolve({ status: "unreadable" });
      return;
    }
    worker.stdout.resume();
    worker.stderr.resume();
    let done = false;
    const finish = (result: PdfResult): void => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      void worker.terminate();
      resolve(result);
    };
    const timer = setTimeout(
      () => finish({ status: "too-slow" }),
      options.timeoutMs ?? PDF_LIMITS.timeoutMs,
    );
    worker.once("message", (m: unknown) => finish(isExtract(m) ? m : { status: "unreadable" }));
    worker.once("error", (e: Error & { code?: string }) =>
      finish(
        e.code === "ERR_WORKER_OUT_OF_MEMORY"
          ? { status: "too-complex" }
          : { status: "unreadable" },
      ),
    );
    worker.once("exit", () => finish({ status: "unreadable" }));
  });
}
