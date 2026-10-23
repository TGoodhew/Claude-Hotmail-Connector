/**
 * Text from a PDF (#156), using pdf.js through unpdf with everything the text
 * doesn't need switched off: no font loading or system fonts, no WebAssembly
 * decoders, no canvas or image decoding, no network (only the bytes given), no
 * logging. (pdf.js 6 no longer compiles font code with eval at all, so there's
 * no isEvalSupported to turn off.) This runs inside the
 * PDF worker (see ./pdf.ts), never in the server's own thread: a PDF is
 * untrusted input and parsing one can be made arbitrarily slow.
 */

import { getDocumentProxy } from "unpdf";

/** What reading a PDF found. */
export type PdfExtract =
  | {
      status: "ok";
      /** The text, page by page; pages after the first are headed "[Page n]". */
      text: string;
      /** Pages read: the first `pagesRead` of `totalPages`. */
      pagesRead: number;
      totalPages: number;
    }
  /** Password-protected. */
  | { status: "encrypted" }
  /** No text layer (e.g. a scan): there's nothing to read without OCR. */
  | { status: "no-text" }
  /** Not a PDF, or too broken to read. */
  | { status: "unreadable" };

/** The text of the first `maxPages` pages of `data`. Never throws. */
export async function extractPdfText(data: Uint8Array, maxPages: number): Promise<PdfExtract> {
  let pdf: Awaited<ReturnType<typeof getDocumentProxy>>;
  try {
    // A plain copy: pdf.js refuses Node Buffers, and may detach what it's given.
    pdf = await getDocumentProxy(new Uint8Array(data), {
      disableFontFace: true,
      useSystemFonts: false,
      useWasm: false,
      useWorkerFetch: false,
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      verbosity: 0,
    });
  } catch (e) {
    return (e as { name?: string } | null)?.name === "PasswordException"
      ? { status: "encrypted" }
      : { status: "unreadable" };
  }
  try {
    const totalPages = pdf.numPages;
    const pagesRead = Math.min(totalPages, maxPages);
    const pages: string[] = [];
    for (let n = 1; n <= pagesRead; n++) {
      const page = await pdf.getPage(n);
      const content = await page.getTextContent();
      const lines: string[] = [];
      let line = "";
      for (const item of content.items) {
        if (!("str" in item)) continue;
        line += item.str;
        if (item.hasEOL) {
          lines.push(line.trimEnd());
          line = "";
        }
      }
      if (line) lines.push(line.trimEnd());
      pages.push(lines.join("\n").trim());
      page.cleanup();
    }
    if (pages.every((p) => p === "")) return { status: "no-text" };
    const text = pages.map((p, i) => (i === 0 ? p : `[Page ${i + 1}]\n${p}`)).join("\n\n");
    return { status: "ok", text, pagesRead, totalPages };
  } catch {
    return { status: "unreadable" };
  } finally {
    await pdf.loadingTask.destroy().catch(() => undefined);
  }
}
