import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it, vi } from "vitest";
import { makePdf } from "../test-support/pdf.js";
import { defaultWorkerFile, readPdfText } from "./pdf.js";

// Stand-in workers, to check the limits and the plumbing without pdf.js.
const dir = mkdtempSync(join(tmpdir(), "hotmail-pdf-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
function worker(name: string, code: string): string {
  const file = join(dir, `${name}.mjs`);
  writeFileSync(file, `import { parentPort, workerData } from "node:worker_threads";\n${code}\n`);
  return file;
}

describe("readPdfText (#156)", () => {
  it("passes on what the worker found, with the data and the page limit", async () => {
    const echo = worker(
      "echo",
      `parentPort.postMessage({ status: "ok", text: String(workerData.data.length), pagesRead: workerData.maxPages, totalPages: workerData.maxChars, textCut: false });`,
    );
    expect(
      await readPdfText(new Uint8Array(7), { workerFile: echo, maxPages: 3, maxChars: 9 }),
    ).toEqual({ status: "ok", text: "7", pagesRead: 3, totalPages: 9, textCut: false });
  });

  it("stops a worker that runs too long", async () => {
    const hang = worker("hang", "setInterval(() => {}, 1000);");
    expect(await readPdfText(new Uint8Array(1), { workerFile: hang, timeoutMs: 100 })).toEqual({
      status: "too-slow",
    });
  });

  it("stops a worker that uses too much memory", async () => {
    const hog = worker(
      "hog",
      "const keep = []; for (;;) keep.push(new Array(1_000_000).fill(Math.random()));",
    );
    const res = await readPdfText(new Uint8Array(1), {
      workerFile: hog,
      maxHeapMb: 16,
      timeoutMs: 20_000,
    });
    expect(res).toEqual({ status: "too-complex" });
  });

  it("stops a worker whose buffers grow too far, which the heap cap doesn't cover", async () => {
    // pdf.js decodes compressed streams into typed arrays, outside the JS heap
    // (Copilot review on #171).
    const buffers = worker(
      "buffers",
      "const keep = []; setInterval(() => keep.push(Buffer.alloc(16 * 1048576, 1)), 5);",
    );
    const res = await readPdfText(new Uint8Array(1), {
      workerFile: buffers,
      maxGrowthMb: 64,
      timeoutMs: 20_000,
    });
    expect(res).toEqual({ status: "too-complex" });
  });

  it("treats a malformed reply, a crash or a silent exit as unreadable", async () => {
    const odd = worker("odd", `parentPort.postMessage({ status: "ok", text: 42 });`);
    const crash = worker("crash", `throw new Error("boom");`);
    const quiet = worker("quiet", "");
    for (const workerFile of [odd, crash, quiet, join(dir, "missing.mjs")]) {
      expect(await readPdfText(new Uint8Array(1), { workerFile }), workerFile).toEqual({
        status: "unreadable",
      });
    }
  });

  it("keeps whatever the worker prints off the server's stdout (the MCP channel)", async () => {
    const write = vi.spyOn(process.stdout, "write");
    const noisy = worker(
      "noisy",
      `console.log("NOISE-1234"); process.stdout.write("NOISE-5678"); parentPort.postMessage({ status: "no-text" });`,
    );
    expect(await readPdfText(new Uint8Array(1), { workerFile: noisy })).toEqual({
      status: "no-text",
    });
    expect(write.mock.calls.map(([c]) => String(c)).join("")).not.toContain("NOISE");
    write.mockRestore();
  });

  it("finds the built worker next to the server", () => {
    expect(String(defaultWorkerFile())).toMatch(/pdf-worker\.(cjs|js)$/);
  });
});

// The real workers, once built. CI builds both before testing and sets
// REQUIRE_BUILT_PDF_WORKERS, so a missing worker fails there instead of being
// skipped (Copilot review on #171).
const required = process.env.REQUIRE_BUILT_PDF_WORKERS === "1";
describe.each([
  ["dist/pdf-worker.js", "source checkout"],
  ["build/mcpb/server/pdf-worker.cjs", "Claude Desktop extension"],
])("built PDF worker %s (%s)", (file) => {
  it.runIf(required || existsSync(file))("reads a PDF's text", async () => {
    expect(existsSync(file), `${file} should be built`).toBe(true);
    const res = await readPdfText(makePdf(["Your booking: ABC123", "Total $42"]), {
      workerFile: join(process.cwd(), file),
    });
    expect(res).toEqual({
      status: "ok",
      text: "Your booking: ABC123\n\n[Page 2]\nTotal $42",
      pagesRead: 2,
      totalPages: 2,
      textCut: false,
    });
  });
});
