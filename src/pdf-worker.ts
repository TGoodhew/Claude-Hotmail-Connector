/**
 * The PDF worker (#156): reads a PDF's text in its own thread, so a hostile
 * PDF can't stall or exhaust the server. Started by src/graph/pdf.ts, which
 * gives it a memory cap and a time limit and discards anything it prints.
 * Built as its own file next to the server (dist/pdf-worker.js, and
 * server/pdf-worker.cjs in the Claude Desktop extension).
 */

import { parentPort, workerData } from "node:worker_threads";
import { extractPdfText } from "./graph/pdf-text.js";

const { data, maxPages } = workerData as { data: Uint8Array; maxPages: number };
extractPdfText(data, maxPages).then(
  (result) => parentPort?.postMessage(result),
  () => parentPort?.postMessage({ status: "unreadable" }),
);
