/**
 * Build small, valid PDFs for tests (#156): one page per string, Helvetica
 * text, with a correct cross-reference table.
 */
export function makePdf(pages: string[], extraTrailer = ""): Buffer {
  const objects: string[] = [];
  const pageIds = pages.map((_, i) => 4 + i * 2);
  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pages.length} >>`;
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>";
  pages.forEach((text, i) => {
    const page = 4 + i * 2;
    const escaped = text.replace(/[\\()]/g, (c) => String.fromCharCode(92) + c);
    const stream = `BT /F1 12 Tf 72 720 Td (${escaped}) Tj ET`;
    objects[page] =
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] " +
      `/Resources << /Font << /F1 3 0 R >> >> /Contents ${page + 1} 0 R >>`;
    objects[page + 1] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`;
  });
  let out = "%PDF-1.4\n";
  const offsets: number[] = [];
  for (let id = 1; id < objects.length; id++) {
    offsets[id] = Buffer.byteLength(out, "latin1");
    out += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(out, "latin1");
  out += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) {
    out += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
  }
  out += `trailer\n<< /Size ${objects.length} /Root 1 0 R ${extraTrailer}>>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}
