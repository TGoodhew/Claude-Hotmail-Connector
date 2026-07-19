/**
 * Minimal, dependency-free HTML -> plain-text conversion.
 *
 * Email bodies arrive as HTML; converting to text server-side keeps tool
 * output small (token-efficient) and readable. This is intentionally simple —
 * it drops markup and script/style, turns block elements into line breaks, and
 * decodes the common entities. It is not a full HTML renderer.
 */

const BLOCK_BREAK_RE = /<\/(p|div|li|tr|h[1-6]|blockquote|section|article)>/gi;
const LINE_BREAK_RE = /<br\s*\/?>/gi;
const SCRIPT_STYLE_RE = /<(script|style)[^>]*>[\s\S]*?<\/\1>/gi;
const TAG_RE = /<[^>]+>/g;

const ENTITIES: Record<string, string> = {
  "&nbsp;": " ",
  "&amp;": "&",
  "&lt;": "<",
  "&gt;": ">",
  "&quot;": '"',
  "&#39;": "'",
  "&apos;": "'",
};

function decodeEntities(text: string): string {
  return text
    .replace(/&nbsp;|&amp;|&lt;|&gt;|&quot;|&#39;|&apos;/g, (m) => ENTITIES[m] ?? m)
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)));
}

/** Convert an HTML string to readable plain text. */
export function htmlToText(html: string): string {
  if (!html) return "";
  const withBreaks = html
    .replace(SCRIPT_STYLE_RE, "")
    .replace(LINE_BREAK_RE, "\n")
    .replace(BLOCK_BREAK_RE, "\n");
  const stripped = withBreaks.replace(TAG_RE, "");
  const decoded = decodeEntities(stripped);
  return decoded
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n") // trailing spaces before newlines
    .replace(/\n{3,}/g, "\n\n") // collapse excessive blank lines
    .replace(/[ \t]{2,}/g, " ") // collapse runs of spaces
    .trim();
}
