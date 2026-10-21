/**
 * Minimal, dependency-free HTML -> plain-text conversion.
 *
 * Email bodies arrive as HTML; converting to text server-side keeps tool
 * output small (token-efficient) and readable. This is intentionally simple —
 * it drops markup, comments and non-content elements, turns block and table
 * elements into line/column breaks, and decodes entities. It is not a full HTML
 * renderer, and it never throws on malformed input (#32).
 *
 * The markup is read by one left-to-right scan (see `render`): each complete
 * tag, quoted attribute values included, or comment is consumed before anything
 * is decided, so text inside an attribute can't start a comment, end a table
 * cell or break a tag. Mail is untrusted and conversion is synchronous, so every
 * search runs forward and a closer found missing isn't looked for again: the
 * scan stays linear on hostile input.
 */

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

// ── Scanner ──────────────────────────────────────────────────────────────────

const TAG_NAME_RE = /[A-Za-z][A-Za-z0-9:-]*/y;
const WHITESPACE_RUN_RE = /\s+/g;
// JS `\s` includes the no-break space and other Unicode spaces.
const ONLY_SPACE_RE = /^\s*$/;
// Characters that render as nothing (see ALWAYS_INVISIBLE_RE and JOINERS).
const INVISIBLE_PADDING_RE = new RegExp(`[${ch(0xad, 0xfeff)}${JOINERS}]`, "g");

/** Raw text, removed with their content (<head> is handled by the scan itself). */
const NON_CONTENT = new Set(["script", "style", "title"]);
/** Their end starts a new line. */
const BLOCKS = new Set([
  "p", "div", "li", "caption", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote",
  "section", "article", "header", "footer", "ul", "ol", "dl", "dt", "dd", "pre",
]); // prettier-ignore
/** Cells end in a tab and rows in a line break (see `render`). */
const TABLE_PARTS = new Set(["td", "th", "tr", "table"]);

interface Tag {
  name: string;
  close: boolean;
  /** Index just after the tag's `>`. */
  end: number;
}

/**
 * What one scan has learnt is missing from the current position on (a quote
 * character, `-->`, `>`, or an element's closing tag). Searches only run
 * forward, so a miss stays a miss and isn't searched for again.
 */
type Missing = Set<string>;

/**
 * Read the tag that starts at `i` (where `html[i]` is `<`), or null if it isn't
 * one: no name, or no `>` before the next unquoted `<` or the end. Quoted
 * attribute values may contain `<` and `>`.
 */
function readTag(html: string, i: number, missing: Missing): Tag | null {
  let j = i + 1;
  const close = html[j] === "/";
  if (close) j += 1;
  TAG_NAME_RE.lastIndex = j;
  const name = TAG_NAME_RE.exec(html)?.[0];
  if (!name) return null;
  for (j += name.length; j < html.length; j += 1) {
    const c = html[j];
    if (c === ">") return { name: name.toLowerCase(), close, end: j + 1 };
    if (c === "<") return null;
    if (c === '"' || c === "'") {
      const q = missing.has(c) ? -1 : html.indexOf(c, j + 1);
      if (q < 0) {
        missing.add(c);
        return null;
      }
      j = q;
    }
  }
  return null;
}

/** The index just after the closing tag of `name` from `from`, or -1 (remembered). */
function findClose(html: string, name: string, from: number, missing: Missing): number {
  const key = `</${name}`;
  if (missing.has(key)) return -1;
  const re = new RegExp(`</${name}\\s*>`, "gi");
  re.lastIndex = from;
  const m = re.exec(html);
  if (!m) {
    missing.add(key);
    return -1;
  }
  return m.index + m[0].length;
}

interface TableState {
  cell: boolean;
  row: boolean;
}

/**
 * Render markup as text with `\n` line breaks, `\t` after each table cell and
 * `<pre>` blocks as placeholders (their text goes into `pres`). Inside <pre>,
 * spacing is kept and only <br> breaks lines; <head> content is dropped. Both
 * hold ordinary markup, so they end at a real `</pre>` / `</head>` tag found by
 * this same scan, never at one inside an attribute or comment. Only <script>,
 * <style> and <title> are raw text, ended (as browsers do) at the first closer.
 *
 * Block ends and <br> owe a line break that is written only when more text
 * follows, so breaks that end a table cell's content (`<td><p>Total</p></td>`,
 * `<td><span>Total<br></span><td>`) are dropped instead of splitting the row.
 *
 * HTML lets `</td>`, `</th>` and `</tr>` be left out (the next cell or row ends
 * them), so the open cell and row are tracked per table, nested tables
 * included, rather than relying on closing tags.
 */
function render(html: string, pres: string[], ignoreHead: boolean): string | null {
  const missing: Missing = new Set();
  const out: string[] = [];
  let owed = 0; // line breaks owed by block ends and <br>
  const outer: TableState[] = [];
  let table: TableState = { cell: false, row: false };
  let preText: string[] | null = null; // inside <pre>: its text so far
  let preStart = false; // nothing has followed the <pre> tag yet
  let inHead = false;
  let cellEmpty = false; // an open cell has had no content yet

  const write = (s: string) => {
    if (owed > 0) out.push(owed > 1 ? "\n\n" : "\n");
    owed = 0;
    // Indentation or padding before a cell's first content doesn't count as content.
    if (!ONLY_SPACE_RE.test(s.replace(INVISIBLE_PADDING_RE, ""))) cellEmpty = false;
    out.push(s);
  };
  const endPre = (parts: string[]) => {
    preText = null;
    const block = tidy(parts.join("")).trimEnd();
    // A <pre> is a block: a line break before it (none at the start of a cell)
    // and one owed after it, which ending the cell cancels like any other.
    if (!(table.cell && cellEmpty)) owed += 1;
    write(`${PRE_OPEN}${pres.push(block) - 1}${PRE_CLOSE}`);
    owed = 1;
  };
  const text = (raw: string) => {
    if (inHead) return;
    if (preText) {
      // HTML ignores one newline straight after the <pre> tag in the source,
      // but keeps one from markup (`<pre><br>`) or after anything else (#125).
      const t = preStart ? raw.replace(/^(?:\r\n|\r|\n)/, "") : raw; // one line ending, any style
      preStart = false;
      return void preText.push(decodeEntities(t));
    }
    // Decoded first, then invisibles dropped, then whitespace collapsed: an
    // encoded tab or newline (`&#9;`, `&#10;`) is spacing, as in a browser, not a
    // column or line break, and a BOM vanishes whether literal or encoded.
    // Spacing and invisible padding (soft hyphens, joiners) don't settle a
    // pending break; the padding itself is kept for the final cleanup.
    const t = decodeEntities(raw).replace(ALWAYS_INVISIBLE_RE, "").replace(WHITESPACE_RUN_RE, " ");
    if (t === "") return;
    if (owed > 0 && ONLY_SPACE_RE.test(t.replace(INVISIBLE_PADDING_RE, ""))) return;
    write(t);
  };
  const endCell = () => {
    if (table.cell) out.push("\t");
    table.cell = false;
  };
  const endRow = () => {
    endCell();
    if (table.row) out.push("\n");
    table.row = false;
  };
  /** Breaks owed when an open cell ends belong to its content: drop them. Others stay. */
  const dropCellEndBreaks = () => {
    if (table.cell) owed = 0;
  };
  const tablePart = (name: string, close: boolean) => {
    if (name === "table") {
      if (!close) {
        write(""); // breaks before the table are kept
        outer.push(table);
        table = { cell: false, row: false };
        return;
      }
      dropCellEndBreaks();
      endRow();
      table = outer.pop() ?? { cell: false, row: false };
      out.push("\n");
      return;
    }
    dropCellEndBreaks();
    if (name === "tr") {
      endRow();
      table.row = !close;
      return;
    }
    endCell(); // td/th: a new cell ends the open one
    if (close) return;
    // Breaks still owed here come from before the cell (a caption, say), not
    // from its content: write them now, so an empty cell can't cancel them.
    write("");
    table.cell = table.row = true;
    cellEmpty = true;
  };

  let i = 0;
  while (i < html.length) {
    const lt = html.indexOf("<", i);
    if (lt < 0) {
      text(html.slice(i));
      break;
    }
    if (lt > i) text(html.slice(i, lt));
    i = lt;
    preStart = false; // a tag or comment comes first, so a later newline counts

    if (html.startsWith("<!--", i)) {
      const end = missing.has("-->") ? -1 : html.indexOf("-->", i + 4);
      if (end < 0) {
        missing.add("-->");
        text("<!--"); // unclosed: kept as text
        i += 4;
      } else {
        i = end + 3;
      }
      continue;
    }
    if (html[i + 1] === "!" || html[i + 1] === "?") {
      // <!DOCTYPE>, <![if !mso]>, <![endif]>, <?xml?>: dropped.
      const gt = missing.has(">") ? -1 : html.indexOf(">", i + 2);
      if (gt < 0) {
        missing.add(">");
        text("<");
        i += 1;
      } else {
        i = gt + 1;
      }
      continue;
    }

    const tag = readTag(html, i, missing);
    if (!tag) {
      text("<"); // not a tag (`a < b`, an unclosed quote): it's text
      i += 1;
      continue;
    }
    i = tag.end;
    const { name, close } = tag;

    if (!close && NON_CONTENT.has(name)) {
      const end = findClose(html, name, i, missing);
      if (end >= 0) i = end; // unclosed: only the opening tag is dropped
      continue;
    }
    if (inHead) {
      if ((close && name === "head") || (!close && name === "body")) inHead = false;
      continue;
    }
    if (name === "br" || name === "hr") {
      if (preText) preText.push("\n");
      else owed += 1;
      continue;
    }
    if (preText) {
      // Inside <pre>, other markup only affects style; a real </pre> ends it.
      if (close && name === "pre") endPre(preText);
      continue;
    }
    if (!close && name === "head") inHead = !ignoreHead;
    else if (!close && name === "pre") {
      preText = [];
      preStart = true;
    } else if (TABLE_PARTS.has(name)) tablePart(name, close);
    else if (close && BLOCKS.has(name)) owed += 1;
  }
  // A <head> that never really ended: the caller renders again ignoring it.
  if (inHead) return null;
  if (preText) endPre(preText); // an unclosed <pre> runs to the end, as in a browser
  // The end of input ends open cells, rows and tables as their closing tags
  // would, so a trailing empty cell keeps its column (#125).
  for (let enclosing = outer.pop(); enclosing; enclosing = outer.pop()) {
    dropCellEndBreaks();
    endRow();
    table = enclosing;
    out.push("\n"); // as an explicit </table> does
  }
  dropCellEndBreaks(); // cells and rows outside any table (stray <td>s)
  endRow();
  return out.join("");
}

/** Line endings, invisible characters and no-break spaces, after decoding. */
function tidy(text: string): string {
  return text.replace(/\r\n?/g, "\n").replace(ALWAYS_INVISIBLE_RE, "").replace(NBSP_RE, " ");
}

/** Convert an HTML string to readable plain text. */
export function htmlToText(html: string): string {
  if (!html) return "";
  const pres: string[] = [];
  const source = html.replace(MARKERS_RE, "");
  // A <head> ends at a real </head> or <body> tag found by the scan. If it never
  // does, rendering once more without it keeps a stray <head> from hiding the
  // message (at most two passes, so still linear).
  let rendered = render(source, pres, false);
  if (rendered === null) {
    pres.length = 0;
    rendered = render(source, pres, true) ?? "";
  }
  const text = dropPadding(tidy(rendered))
    .replace(/ {2,}/g, " ") // collapse runs of spaces (first, so the next step can't backtrack)
    .replace(/ ?\t ?/g, "\t")
    .replace(/\t$/gm, "") // the separator after a row's last cell
    .replace(/^[\t ]+$/gm, "") // lines holding only empty (layout) cells
    .replace(/^ +| +$/gm, "") // trim spaces; leading tabs are empty cells
    .replace(/\n{3,}/g, "\n\n") // collapse excessive blank lines
    .replace(/^[\n ]+|[\n ]+$/g, ""); // not tabs: they can be empty cells
  return text.replace(PLACEHOLDER_RE, (_, i: string) => pres[Number(i)] ?? "");
}
