import { describe, expect, it } from "vitest";
import { makePdf } from "../test-support/pdf.js";
import { extractPdfText } from "./pdf-text.js";

/** A trailer making the file password-protected (the empty password doesn't open it). */
const ENCRYPTED =
  `/Encrypt << /Filter /Standard /V 1 /R 2 /P -4 /O <${"1".repeat(64)}> /U <${"2".repeat(64)}> >> ` +
  `/ID [<${"3".repeat(32)}> <${"3".repeat(32)}>] `;

describe("extractPdfText (#156)", () => {
  it("reads the text page by page, heading pages after the first", async () => {
    const res = await extractPdfText(makePdf(["Booking ref ABC123", "Check-in from 3pm"]), 30);
    expect(res).toEqual({
      status: "ok",
      text: "Booking ref ABC123\n\n[Page 2]\nCheck-in from 3pm",
      pagesRead: 2,
      totalPages: 2,
    });
  });

  it("reads only the first maxPages pages, and says how many there were", async () => {
    const pages = Array.from({ length: 5 }, (_, i) => `Page text ${i + 1}`);
    const res = await extractPdfText(makePdf(pages), 2);
    expect(res).toMatchObject({ status: "ok", pagesRead: 2, totalPages: 5 });
    expect(res.status === "ok" && res.text).not.toContain("Page text 3");
  });

  it("says when there's no text to read (e.g. a scan)", async () => {
    expect(await extractPdfText(makePdf(["", ""]), 30)).toEqual({ status: "no-text" });
  });

  it("says when it's password-protected", async () => {
    expect(await extractPdfText(makePdf(["Secret"], ENCRYPTED), 30)).toEqual({
      status: "encrypted",
    });
  });

  it("says when it isn't a readable PDF, without throwing", async () => {
    expect(await extractPdfText(Buffer.from("not a pdf at all"), 30)).toEqual({
      status: "unreadable",
    });
    expect(await extractPdfText(new Uint8Array(), 30)).toEqual({ status: "unreadable" });
  });
});
