/**
 * Reading an email attachment's content (#54, Mail.Read). Only what can be
 * returned as text: plain text and HTML (converted to text), calendar files
 * (parsed into events), and PDFs (their text, read in an isolated worker,
 * #156). Images and any other binary are refused, never returned raw.
 */

import { htmlToText } from "../util/html.js";
import { parseIcs, type IcsCalendar } from "../util/ics.js";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { sliceBody, type MessageBody } from "./mail.js";
import { readPdfText, type PdfResult } from "./pdf.js";

/** Largest text, HTML or calendar attachment read, in bytes; booking texts and invitations are far smaller. */
export const MAX_ATTACHMENT_BYTES = 1_048_576;
/** Largest PDF read: booking PDFs with a logo or a map often run to a few MB (#156). */
export const MAX_PDF_BYTES = 5 * 1_048_576;

export interface GetAttachmentInput {
  messageId: string;
  attachmentId: string;
  /** For text and HTML: most characters to return, and where to start (as get_message). */
  maxChars?: number;
  startChar?: number;
}

export type AttachmentKind = "text" | "html" | "calendar" | "pdf";

export interface AttachmentContent {
  id: string;
  name: string | null;
  contentType: string | null;
  size: number | null;
  kind: AttachmentKind;
  /** Text, HTML and PDF: the text (HTML converted), cut at maxChars like a message body. */
  text?: MessageBody;
  /** PDFs: the first `pagesRead` of `totalPages` pages were read. */
  pdf?: { pagesRead: number; totalPages: number };
  /** Calendar files: the events in them. */
  calendar?: IcsCalendar;
}

interface GraphAttachment {
  "@odata.type"?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  contentBytes?: string;
}

const FILE_ATTACHMENT = "#microsoft.graph.fileAttachment";

/** Plain-text types read as text (besides text/*). */
const TEXT_TYPES = new Set(["application/json", "application/xml"]);
const TEXT_EXTENSIONS = new Set(["txt", "csv", "tsv", "md", "log", "json", "xml"]);

/** The media type, lower case and without parameters. */
function mediaType(contentType: string | undefined): string {
  return (contentType ?? "").split(";")[0]!.trim().toLowerCase();
}

function extension(name: string | undefined): string {
  const m = /\.([A-Za-z0-9]+)$/.exec(name ?? "");
  return m ? m[1]!.toLowerCase() : "";
}

/**
 * What an attachment can be read as: by its declared type, or by its file name
 * only when the type says nothing (mail clients often send
 * `application/octet-stream`). Null when it can't be read.
 */
export function attachmentKind(contentType?: string, name?: string): AttachmentKind | null {
  const type = mediaType(contentType);
  if (type === "" || type === "application/octet-stream") {
    const ext = extension(name);
    if (ext === "ics") return "calendar";
    if (ext === "pdf") return "pdf";
    if (ext === "html" || ext === "htm") return "html";
    return TEXT_EXTENSIONS.has(ext) ? "text" : null;
  }
  if (type === "text/calendar" || type === "application/ics") return "calendar";
  if (type === "application/pdf") return "pdf";
  if (type === "text/html") return "html";
  return type.startsWith("text/") || TEXT_TYPES.has(type) ? "text" : null;
}

// Refusals are fixed wording: an attachment's name and type are the sender's
// text, and an error isn't marked untrusted the way a result is (T12).
const OPEN_IN_OUTLOOK = "Open it from the message in Outlook (the message's webLink).";
const NOT_A_FILE = `This attachment is an attached item or a link, not a file. ${OPEN_IN_OUTLOOK}`;
const tooLarge = (limit: number) =>
  `This attachment is too large to read here (over ${limit / 1_048_576} MB). ${OPEN_IN_OUTLOOK}`;
const NOT_TEXT =
  "This attachment can't be read as text: only text, HTML, calendar (.ics) and PDF attachments " +
  `can. ${OPEN_IN_OUTLOOK}`;
const PDF_REFUSALS: Record<Exclude<PdfResult["status"], "ok">, string> = {
  encrypted: `This PDF is password-protected, so it can't be read here. ${OPEN_IN_OUTLOOK}`,
  "no-text": `This PDF has no text to read; it may be a scanned image. ${OPEN_IN_OUTLOOK}`,
  unreadable: `This PDF couldn't be read; it may be damaged. ${OPEN_IN_OUTLOOK}`,
  "too-slow": `This PDF is too complex to read here in reasonable time. ${OPEN_IN_OUTLOOK}`,
  "too-complex": `This PDF is too complex to read here safely. ${OPEN_IN_OUTLOOK}`,
};

/** Remove iCalendar line folds (CRLF or LF, then a space or tab) at the byte level. */
export function unfoldBytes(bytes: Buffer): Buffer {
  const out = Buffer.alloc(bytes.length);
  let n = 0;
  for (let i = 0; i < bytes.length; i++) {
    const crlf = bytes[i] === 0x0d && bytes[i + 1] === 0x0a ? 2 : 0;
    const eol = crlf || (bytes[i] === 0x0a ? 1 : 0);
    const next = bytes[i + eol];
    if (eol && (next === 0x20 || next === 0x09)) {
      i += eol; // skip the line end and the one whitespace after it
      continue;
    }
    out[n++] = bytes[i]!;
  }
  return out.subarray(0, n);
}

/**
 * Decode text bytes: by the declared charset when there is one, else UTF-8,
 * falling back to Windows-1252 (which every byte decodes as) for older files.
 * A NUL byte means it isn't text after all.
 */
function decodeText(bytes: Buffer, contentType: string | undefined): string {
  if (bytes.includes(0)) throw new ValidationError(NOT_TEXT);
  const charset = /;\s*charset="?([^";\s]+)"?/i.exec(contentType ?? "")?.[1];
  if (charset) {
    try {
      return new TextDecoder(charset).decode(bytes);
    } catch {
      // An unknown label: read it as below.
    }
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

/**
 * Read one attachment's content as text or events. Its details are read first,
 * so a file that's too large or can't be read as text is refused before any of
 * its content is downloaded.
 */
export async function getAttachment(
  graph: GraphClient,
  input: GetAttachmentInput,
  // The PDF reader, replaceable in tests.
  deps: { readPdf?: (data: Uint8Array) => Promise<PdfResult> } = {},
): Promise<AttachmentContent> {
  const path =
    `/me/messages/${encodeURIComponent(input.messageId)}` +
    `/attachments/${encodeURIComponent(input.attachmentId)}`;
  const meta = await graph.get<GraphAttachment>(path, {
    query: { $select: "id,name,contentType,size,isInline" },
  });
  const assertFile = (a: GraphAttachment) => {
    if (a["@odata.type"] && a["@odata.type"] !== FILE_ATTACHMENT) {
      throw new ValidationError(NOT_A_FILE);
    }
  };
  assertFile(meta);
  const kind = attachmentKind(meta.contentType, meta.name);
  if (!kind) throw new ValidationError(NOT_TEXT);
  const limit = kind === "pdf" ? MAX_PDF_BYTES : MAX_ATTACHMENT_BYTES;
  if ((meta.size ?? 0) > limit) throw new ValidationError(tooLarge(limit));

  const full = await graph.get<GraphAttachment>(path);
  assertFile(full);
  const bytes = Buffer.from(full.contentBytes ?? "", "base64");
  // The size above is Graph's figure; check what actually arrived too.
  if (bytes.length > limit) throw new ValidationError(tooLarge(limit));

  const base = {
    id: meta.id ?? input.attachmentId,
    name: meta.name ?? null,
    contentType: meta.contentType ?? null,
    size: meta.size ?? null,
    kind,
  };
  if (kind === "pdf") {
    const read = await (deps.readPdf ?? readPdfText)(bytes);
    if (read.status !== "ok") throw new ValidationError(PDF_REFUSALS[read.status]);
    return {
      ...base,
      text: sliceBody(read.text, "text", input.maxChars, input.startChar),
      pdf: { pagesRead: read.pagesRead, totalPages: read.totalPages },
    };
  }
  // A calendar line can be folded between the bytes of one character, so
  // unfold before decoding (the size was checked on the bytes as they came).
  const text = decodeText(
    kind === "calendar" ? unfoldBytes(bytes) : bytes,
    full.contentType ?? meta.contentType,
  );
  if (kind === "calendar") return { ...base, calendar: parseIcs(text) };
  const content = kind === "html" ? htmlToText(text) : text;
  return { ...base, text: sliceBody(content, "text", input.maxChars, input.startChar) };
}
