/**
 * Minimal, dependency-free HTML -> plain-text conversion.
 *
 * Email bodies arrive as HTML; converting to text server-side keeps tool
 * output small (token-efficient) and readable. This is intentionally simple —
 * it drops markup, comments and non-content elements, turns block and table
 * elements into line/column breaks, and decodes entities. It is not a full HTML
 * renderer, and it never throws on malformed input (#32).
 */

// Removed with their content. Comments go first: they can hide tags, and
// Outlook's conditional comments (`<!--[if mso]>…<![endif]-->`) are comments.
const COMMENT_RE = /<!--[\s\S]*?-->/g;
const NON_CONTENT_RE = /<(script|style|head|title)\b[^>]*>[\s\S]*?<\/\1\s*>/gi;
// Preformatted blocks keep their own spacing (see htmlToText).
const PRE_RE = /(<pre\b[^>]*>[\s\S]*?<\/pre\s*>)/gi;
const LINE_BREAK_RE = /<(br|hr)\b[^>]*>/gi;
const BLOCK_BREAK_RE =
  /<\/(p|div|li|tr|table|thead|tbody|tfoot|caption|h[1-6]|blockquote|section|article|header|footer|ul|ol|dl|dt|dd|pre)\s*>/gi;
const CELL_BREAK_RE = /<\/(td|th)\s*>/gi;
const TAG_RE = /<[^>]+>/g;

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

/** A numeric reference's character; NUL, surrogates and out-of-range values become U+FFFD. */
function fromReference(code: number): string {
  if (code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return "\ufffd";
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
const PADDING_RE = new RegExp(`(?<=^|\\s)[${JOINERS}]+|[${JOINERS}]+(?=\\s|$)`, "gm");
const NBSP_RE = new RegExp(ch(0xa0), "g");

// Each <pre> block waits as a placeholder while the rest is normalised, so its
// spacing and blank lines survive. The input can't contain the marker characters.
const PRE_OPEN = ch(0xe000);
const PRE_CLOSE = ch(0xe001);
const MARKERS_RE = new RegExp(`[${PRE_OPEN}${PRE_CLOSE}]`, "g");
const PLACEHOLDER_RE = new RegExp(`${PRE_OPEN}(\\d+)${PRE_CLOSE}`, "g");

/** Markup to text: line breaks kept, tags stripped, entities decoded, invisibles dropped. */
function inlineText(html: string): string {
  return decodeEntities(html.replace(LINE_BREAK_RE, "\n").replace(TAG_RE, ""))
    .replace(/\r\n?/g, "\n")
    .replace(ALWAYS_INVISIBLE_RE, "")
    .replace(NBSP_RE, " ");
}

/** A <pre> block's text, with its spacing and line breaks as written. */
function preText(block: string): string {
  const inner = block.replace(/^<pre\b[^>]*>/i, "").replace(/<\/pre\s*>$/i, "");
  return inlineText(inner)
    .replace(/^\n/, "") // HTML ignores a newline straight after <pre>
    .replace(/\s+$/, "");
}

/** Convert an HTML string to readable plain text. */
export function htmlToText(html: string): string {
  if (!html) return "";
  const pres: string[] = [];
  const marked = html
    .replace(MARKERS_RE, "")
    .replace(COMMENT_RE, "")
    .replace(NON_CONTENT_RE, "")
    .replace(PRE_RE, (block) => `${PRE_OPEN}${pres.push(preText(block)) - 1}${PRE_CLOSE}`)
    .replace(/\s+/g, " ") // source whitespace renders as one space
    .replace(PLACEHOLDER_RE, "\n$&\n")
    .replace(LINE_BREAK_RE, "\n")
    .replace(BLOCK_BREAK_RE, "\n")
    .replace(CELL_BREAK_RE, "\t"); // one tab per cell, so empty cells keep their column
  const text = inlineText(marked)
    .replace(PADDING_RE, "")
    .replace(/ *\t */g, "\t")
    .replace(/ {2,}/g, " ") // collapse runs of spaces
    .replace(/\t$/gm, "") // the separator after a row's last cell
    .replace(/^[\t ]+$/gm, "") // lines holding only empty (layout) cells
    .replace(/^ +| +$/gm, "") // trim spaces; leading tabs are empty cells
    .replace(/\n{3,}/g, "\n\n") // collapse excessive blank lines
    .replace(/^[\n ]+|[\n ]+$/g, ""); // not tabs: they can be empty cells
  return text.replace(PLACEHOLDER_RE, (_, i: string) => pres[Number(i)] ?? "");
}
