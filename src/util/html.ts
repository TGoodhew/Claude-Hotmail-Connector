/**
 * Minimal, dependency-free HTML -> plain-text conversion.
 *
 * Email bodies arrive as HTML; converting to text server-side keeps tool
 * output small (token-efficient) and readable. This is intentionally simple —
 * it drops markup, comments and non-content elements, turns block and table
 * elements into line/column breaks, and decodes entities. It is not a full HTML
 * renderer, and it never throws on malformed input (#32).
 */

// Removed with their content, in one left-to-right scan (replaceSpans), so
// whichever starts first wins: a `<!--` inside a <style> can't swallow the text
// after it, and a <style> inside a comment goes with the comment. Outlook's
// conditional comments (`<!--[if mso]>…<![endif]-->`) are comments.
const NON_CONTENT_OPEN_RE = /<!--|<(script|style|head|title)\b[^<>]*>/gi;
// Preformatted blocks keep their own spacing (see htmlToText).
const PRE_OPEN_RE = /<(pre)\b[^<>]*>/gi;
// Tag patterns stop at the next `<` (`[^<>]`), so a stray `<` costs only the
// text up to the next one, and malformed mail can't make the scan quadratic.
const LINE_BREAK_RE = /<(br|hr)\b[^<>]*>/gi;
const BLOCK_BREAK_RE =
  /<\/(p|div|li|tr|table|thead|tbody|tfoot|caption|h[1-6]|blockquote|section|article|header|footer|ul|ol|dl|dt|dd|pre)\s*>/gi;
const CELL_BREAK_RE = /<\/(td|th)\s*>/gi;
// Breaks that end a cell's content (`<td><p>Total</p></td>`) would split a row
// across lines, so they're dropped; breaks between blocks inside a cell stay.
// Each run of break tags is matched once, with the cell close (if any) captured
// rather than looked ahead for, so a long run outside a table stays linear.
const CELL_END_BREAKS_RE =
  /(?:\s*(?:<\/(?:p|div|li|h[1-6]|blockquote|section|article|header|footer|ul|ol|dl|dt|dd)\s*>|<br\b[^<>]*>))+\s*(<\/t[dh]\s*>)?/gi;
const TAG_RE = /<[^<>]+>/g;

const ENTITY_RE = /&(?:#(\d+)|#x([0-9a-f]+)|([a-z][a-z0-9]*));/gi;

// The HTML 4 Latin-1 entities, in code-point order from U+00A0 to U+00FF.
const LATIN1_NAMES =
  "nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr " +
  "deg plusmn sup2 sup3 acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 " +
  "iquest Agrave Aacute Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml " +
  "Igrave Iacute Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave " +
  "Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde auml aring aelig ccedil " +
  "egrave eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde " +
  "ouml divide oslash ugrave uacute ucirc uuml yacute thorn yuml";

/** Named entities -> code points: Latin-1 plus the punctuation and symbols mail uses. */
export const NAMED_ENTITIES: ReadonlyMap<string, number> = new Map([
  ...LATIN1_NAMES.split(" ").map((name, i) => [name, 0xa0 + i] as const),
  // prettier-ignore
  ...Object.entries({
    quot: 34, amp: 38, apos: 39, lt: 60, gt: 62,
    OElig: 338, oelig: 339, Scaron: 352, scaron: 353, Yuml: 376, fnof: 402, circ: 710, tilde: 732,
    ensp: 8194, emsp: 8195, thinsp: 8201, zwnj: 8204, zwj: 8205, lrm: 8206, rlm: 8207,
    ndash: 8211, mdash: 8212, lsquo: 8216, rsquo: 8217, sbquo: 8218, ldquo: 8220, rdquo: 8221,
    bdquo: 8222, dagger: 8224, Dagger: 8225, bull: 8226, hellip: 8230, permil: 8240,
    prime: 8242, Prime: 8243, lsaquo: 8249, rsaquo: 8250, euro: 8364, trade: 8482,
    larr: 8592, uarr: 8593, rarr: 8594, darr: 8595, harr: 8596, minus: 8722,
  }),
]);

// Numeric references in 0x80-0x9F mean Windows-1252 characters, as browsers
// decode them (Outlook writes `&#150;` for an en dash).
// prettier-ignore
const WINDOWS_1252: Readonly<Record<number, number>> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020,
  0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152,
  0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022,
  0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a,
  0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};

/**
 * A numeric reference's character. NUL, surrogates, out-of-range values and the
 * converter's own placeholder markers (see PRE_OPEN) become U+FFFD.
 */
function fromReference(code: number): string {
  if (code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return "\ufffd";
  if (code === PRE_OPEN_CODE || code === PRE_CLOSE_CODE) return "\ufffd";
  return String.fromCodePoint(WINDOWS_1252[code] ?? code);
}

/** Decode every entity in a single pass, so `&amp;#60;` yields `&#60;`, not `<`. */
function decodeEntities(text: string): string {
  return text.replace(
    ENTITY_RE,
    (match, dec: string | undefined, hex: string | undefined, name: string | undefined) => {
      if (dec !== undefined) return fromReference(Number(dec));
      if (hex !== undefined) return fromReference(parseInt(hex, 16));
      const code = NAMED_ENTITIES.get(name!);
      return code === undefined ? match : String.fromCodePoint(code);
    },
  );
}

const ch = (...codes: number[]): string => String.fromCharCode(...codes);

// Always dropped: soft hyphens and byte-order marks render as nothing.
const ALWAYS_INVISIBLE_RE = new RegExp(`[${ch(0xad, 0xfeff)}]`, "g");
// Joiners and zero-width spaces carry meaning between letters (Persian and
// Indic shaping, joined emoji), so they're dropped only next to whitespace or a
// line end: the padding marketing mail puts in its preheader.
const JOINERS = ch(0x34f, 0x200b, 0x200c, 0x200d, 0x2060);
const JOINER_RUN_RE = new RegExp(`[${JOINERS}]+`, "g");
const SPACE_RE = /\s/;

/** Drop joiner runs that touch whitespace or an end; checked per run, so linear. */
function dropPadding(text: string): string {
  return text.replace(JOINER_RUN_RE, (run, at: number) => {
    const before = at === 0 || SPACE_RE.test(text[at - 1]!);
    const after = at + run.length === text.length || SPACE_RE.test(text[at + run.length]!);
    return before || after ? "" : run;
  });
}
const NBSP_RE = new RegExp(ch(0xa0), "g");

// Each <pre> block waits as a placeholder while the rest is normalised, so its
// spacing and blank lines survive. The markers can't come from the message:
// literal ones are removed first, and references to them decode to U+FFFD.
const PRE_OPEN_CODE = 0xe000;
const PRE_CLOSE_CODE = 0xe001;
const PRE_OPEN = ch(PRE_OPEN_CODE);
const PRE_CLOSE = ch(PRE_CLOSE_CODE);
const MARKERS_RE = new RegExp(`[${PRE_OPEN}${PRE_CLOSE}]`, "g");
const PLACEHOLDER_RE = new RegExp(`${PRE_OPEN}(\\d+)${PRE_CLOSE}`, "g");

/** Markup to text: line breaks kept, tags stripped, entities decoded, invisibles dropped. */
function inlineText(html: string): string {
  return decodeEntities(html.replace(LINE_BREAK_RE, "\n").replace(TAG_RE, ""))
    .replace(/\r\n?/g, "\n")
    .replace(ALWAYS_INVISIBLE_RE, "")
    .replace(NBSP_RE, " ");
}

/**
 * Replace each span from an opener matched by `openRe` (group 1 names the
 * element; no group means a comment) to its closer, scanning once from the
 * left. An opener without a closer is left in place as text, and a closer found
 * missing isn't searched for again, so malformed mail stays linear time.
 */
function replaceSpans(html: string, openRe: RegExp, replace: (span: string) => string): string {
  const closers = new Map<string, RegExp | null>(); // null: none from here on
  let out = "";
  let pos = 0;
  openRe.lastIndex = 0;
  for (let m = openRe.exec(html); m; m = openRe.exec(html)) {
    const kind = m[1]?.toLowerCase() ?? "--";
    let close = closers.get(kind);
    if (close === undefined) {
      close = kind === "--" ? /-->/g : new RegExp(`</${kind}\\s*>`, "gi");
      closers.set(kind, close);
    }
    if (close === null) continue;
    close.lastIndex = openRe.lastIndex;
    const found = close.exec(html);
    if (!found) {
      closers.set(kind, null);
      continue;
    }
    const end = found.index + found[0].length;
    out += html.slice(pos, m.index) + replace(html.slice(m.index, end));
    pos = end;
    openRe.lastIndex = end;
  }
  return out + html.slice(pos);
}

/** A <pre> block's text, with its spacing and line breaks as written. */
function preText(block: string): string {
  const inner = block.replace(/^<pre\b[^>]*>/i, "").replace(/<\/pre\s*>$/i, "");
  return inlineText(inner)
    .replace(/^\n/, "") // HTML ignores a newline straight after <pre>
    .trimEnd();
}

/** Convert an HTML string to readable plain text. */
export function htmlToText(html: string): string {
  if (!html) return "";
  const pres: string[] = [];
  const content = replaceSpans(html.replace(MARKERS_RE, ""), NON_CONTENT_OPEN_RE, () => "");
  const marked = replaceSpans(
    content,
    PRE_OPEN_RE,
    (block) => `${PRE_OPEN}${pres.push(preText(block)) - 1}${PRE_CLOSE}`,
  )
    .replace(/\s+/g, " ") // source whitespace renders as one space
    .replace(PLACEHOLDER_RE, "\n$&\n")
    .replace(CELL_END_BREAKS_RE, (run, cellClose?: string) => cellClose ?? run)
    .replace(LINE_BREAK_RE, "\n")
    .replace(BLOCK_BREAK_RE, "\n")
    .replace(CELL_BREAK_RE, "\t"); // one tab per cell, so empty cells keep their column
  const text = dropPadding(inlineText(marked))
    .replace(/ {2,}/g, " ") // collapse runs of spaces (first, so the next step can't backtrack)
    .replace(/ ?\t ?/g, "\t")
    .replace(/\t$/gm, "") // the separator after a row's last cell
    .replace(/^[\t ]+$/gm, "") // lines holding only empty (layout) cells
    .replace(/^ +| +$/gm, "") // trim spaces; leading tabs are empty cells
    .replace(/\n{3,}/g, "\n\n") // collapse excessive blank lines
    .replace(/^[\n ]+|[\n ]+$/g, ""); // not tabs: they can be empty cells
  return text.replace(PLACEHOLDER_RE, (_, i: string) => pres[Number(i)] ?? "");
}
