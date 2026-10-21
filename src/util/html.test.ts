import { describe, expect, it } from "vitest";
import { htmlToText, NAMED_ENTITIES } from "./html.js";

describe("htmlToText", () => {
  it("returns empty string for empty input", () => {
    expect(htmlToText("")).toBe("");
  });

  it("strips tags and keeps text", () => {
    expect(htmlToText("<p>Hello <b>world</b></p>")).toBe("Hello world");
  });

  it("turns <br> and block ends into newlines", () => {
    expect(htmlToText("Line1<br>Line2")).toBe("Line1\nLine2");
    expect(htmlToText("<p>A</p><p>B</p>")).toBe("A\nB");
  });

  it("removes script and style content", () => {
    const html = "<style>.x{color:red}</style><p>Visible</p><script>alert(1)</script>";
    const out = htmlToText(html);
    expect(out).toBe("Visible");
    expect(out).not.toContain("alert");
    expect(out).not.toContain("color");
  });

  it("decodes common and numeric entities", () => {
    expect(htmlToText("<p>Tom &amp; Jerry &lt;3 &#39;quote&#39; &#x2764;</p>")).toBe(
      "Tom & Jerry <3 'quote' ❤",
    );
    expect(htmlToText("a&nbsp;b")).toBe("a b");
  });

  it("collapses excessive blank lines and spaces", () => {
    expect(htmlToText("<p>A</p><p></p><p></p><p>B</p>")).toBe("A\n\nB");
  });
});

describe("htmlToText entity decoding (#32)", () => {
  it("never throws on invalid numeric references; they become U+FFFD", () => {
    for (const ref of ["&#99999999;", "&#x110000;", "&#xD800;", "&#0;", `&#${"9".repeat(40)};`]) {
      expect(() => htmlToText(`a${ref}b`)).not.toThrow();
      expect(htmlToText(`a${ref}b`)).toBe("a�b");
    }
  });

  it("decodes the named entities booking emails use", () => {
    expect(
      htmlToText("It&rsquo;s confirmed &mdash; Caf&eacute; &amp; Bar&hellip; &euro;120 &copy;"),
    ).toBe("It’s confirmed — Café & Bar… €120 ©");
    expect(htmlToText("&Uuml;ber &times; &divide; &yuml; &frac12;")).toBe("Über × ÷ ÿ ½");
  });

  it("covers every Latin-1 entity, in code-point order", () => {
    expect([...NAMED_ENTITIES].filter(([, c]) => c >= 0xa0 && c <= 0xff)).toHaveLength(96);
    expect(NAMED_ENTITIES.get("nbsp")).toBe(0xa0);
    expect(NAMED_ENTITIES.get("times")).toBe(0xd7);
    expect(NAMED_ENTITIES.get("szlig")).toBe(0xdf);
    expect(NAMED_ENTITIES.get("yuml")).toBe(0xff);
  });

  it("decodes once, so double-encoded text keeps its outer layer", () => {
    expect(htmlToText("&amp;#60; &amp;lt; &amp;amp;")).toBe("&#60; &lt; &amp;");
  });

  it("maps 0x80-0x9F references to Windows-1252, as browsers do", () => {
    expect(htmlToText("9&#150;5 &#147;hi&#148; &#128;")).toBe("9–5 “hi” €");
  });

  it("leaves unknown entities as written", () => {
    expect(htmlToText("&notanentity; &;")).toBe("&notanentity; &;");
  });
});

describe("htmlToText structure (#32)", () => {
  it("drops comments (including Outlook conditionals), <head> and <title>", () => {
    const html =
      "<html><head><title>Secret title</title><style>p{}</style></head><body>" +
      "<!--[if mso]><table><tr><td>MSO only</td></tr></table><![endif]-->" +
      "<!-- tracking <img src=x> --><p>Body</p></body></html>";
    expect(htmlToText(html)).toBe("Body");
  });

  it("doesn't mistake <header> for <head>", () => {
    expect(htmlToText("<header>Top</header><p>Body</p>")).toBe("Top\nBody");
  });

  it("strips invisible preheader padding", () => {
    expect(htmlToText("Hi&zwnj;&nbsp;&#847;&zwnj;&shy; there")).toBe("Hi there");
  });

  it("keeps joiners inside words and emoji (Copilot review on #109)", () => {
    const ZWJ = String.fromCharCode(0x200d);
    const ZWNJ = String.fromCharCode(0x200c);
    const coder = String.fromCodePoint(0x1f469) + ZWJ + String.fromCodePoint(0x1f4bb);
    const persian = "می" + ZWNJ + "خواهم"; // "I want": the joiner stops the letters joining
    expect(htmlToText(`<p>${coder} dev</p>`)).toBe(`${coder} dev`);
    expect(htmlToText(`<p>${persian}</p>`)).toBe(persian);
    // Devanagari ka + virama + ZWJ + ssa: the joiner selects the half form.
    expect(htmlToText("<p>&#2325;&#2381;&zwj;&#2359;</p>")).toBe(
      String.fromCharCode(2325, 2381) + ZWJ + String.fromCharCode(2359),
    );
  });

  it("keeps a column for every cell, empty ones included (Copilot review on #109)", () => {
    const html =
      "<table><tr><th>A</th><th>B</th><th>C</th></tr>" +
      "<tr><td>1</td><td></td><td>3</td></tr>" +
      "<tr><td></td><td>2</td><td></td></tr></table>";
    expect(htmlToText(html)).toBe("A\tB\tC\n1\t\t3\n\t2\t");
  });

  it("keeps <pre> spacing, indentation and blank lines (Copilot review on #109)", () => {
    expect(htmlToText("<pre>First\n  second\nthird    fourth</pre>")).toBe(
      "First\n  second\nthird    fourth",
    );
    expect(
      htmlToText("<p>Intro</p><pre>\n  if (a &lt; b) {\n\n\n\n    go();\n  }</pre><p>After</p>"),
    ).toBe("Intro\n\n  if (a < b) {\n\n\n\n    go();\n  }\nAfter");
  });

  it("lays out a table-based confirmation email row by row", () => {
    const html = `
      <table>
        <tr><th>Booking</th><th>Details</th></tr>
        <tr><td>Check-in</td><td>Mon 20 July 2026, 3:00&nbsp;pm</td></tr>
        <tr><td>Check-out</td><td>Wed 22 July 2026, 10:00&nbsp;am</td></tr>
        <tr><td>Total</td><td>&euro;240.00</td></tr>
      </table>
      <hr>
      <p>Questions? Reply to this email.</p>`;
    expect(htmlToText(html)).toBe(
      [
        "Booking\tDetails",
        "Check-in\tMon 20 July 2026, 3:00 pm",
        "Check-out\tWed 22 July 2026, 10:00 am",
        "Total\t€240.00",
        "",
        "Questions? Reply to this email.",
      ].join("\n"),
    );
  });
});
