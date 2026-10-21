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

  it("can't be made to drop or repeat text with encoded placeholder markers (Copilot review on #109)", () => {
    const FFFD = String.fromCharCode(0xfffd);
    // Without a <pre> block: the text is kept, not taken for a missing placeholder.
    expect(htmlToText("A&#57344;0&#57345;B")).toBe(`A${FFFD}0${FFFD}B`);
    expect(htmlToText("A&#xE000;0&#xE001;B")).toBe(`A${FFFD}0${FFFD}B`);
    // With one: the block appears once, however often the marker is repeated.
    const out = htmlToText(`<pre>SAVED</pre>${"&#57344;0&#57345;".repeat(50)}`);
    expect(out.split("SAVED")).toHaveLength(2);
    // Literal markers in the input are removed before any placeholder exists.
    const literal = String.fromCharCode(0xe000) + "0" + String.fromCharCode(0xe001);
    expect(htmlToText(`<pre>SAVED</pre>${literal}`).split("SAVED")).toHaveLength(2);
  });

  it("keeps a row on one line when cells wrap their content in blocks (Copilot review on #109)", () => {
    expect(
      htmlToText(
        "<table><tr><td><p>Total</p></td><td><div><p>&euro;240</p></div></td></tr>" +
          "<tr><td>Ref<br></td><td><p>A</p><p>B</p></td></tr>" +
          "<tr><td><p></p></td><td><p>x</p></td></tr></table>",
      ),
    ).toBe("Total\t€240\nRef\tA\nB\n\tx");
  });

  it("doesn't let a `<!--` inside <style> swallow the message (Copilot review on #109)", () => {
    expect(htmlToText("<style>p::before{content:'<!--'}</style><p>Body</p><!-- footer -->")).toBe(
      "Body",
    );
    // And the other way round: a <style> inside a comment goes with the comment.
    expect(htmlToText("<!-- <style> --><p>Body</p><style>x{}</style>")).toBe("Body");
  });

  it("keeps the text after an unclosed <head>, <style> or comment", () => {
    expect(htmlToText("<head><meta charset=utf-8><body><p>Hi</p>")).toBe("Hi");
    expect(htmlToText("<p>Hi</p><style>x")).toBe("Hi\nx");
    expect(htmlToText("<p>Hi</p><!--x")).toBe("Hi\n<!--x");
  });

  it("reads tags whose quoted attributes contain < or > (Copilot review on #109)", () => {
    expect(
      htmlToText('<style media="(width < 600px)">.hidden { display:none }</style>Visible'),
    ).toBe("Visible");
    expect(htmlToText('<a title="A < B" href="https://example.test/">Link</a>')).toBe("Link");
    expect(htmlToText("<img alt='x > y'>After")).toBe("After");
    expect(htmlToText('<p class="a">One<br data-x="<">Two</p>')).toBe("One\nTwo");
  });

  it("ends cells and rows whose closing tags are left out (Copilot review on #109)", () => {
    expect(htmlToText("<table><tr><td>Check-in<td>Monday<tr><td>Total<td>&euro;240</table>")).toBe(
      "Check-in\tMonday\nTotal\t€240",
    );
    // Mixed explicit and implied closes don't double up.
    expect(htmlToText("<table><tr><td>a</td><td>b<tr><td>c</td></tr></table>")).toBe("a\tb\nc");
    // Block-wrapped cells with implied closes stay on their row.
    expect(htmlToText("<table><tr><td><p>x</p><td><p>y</p><tr><th>z</table>")).toBe("x\ty\nz");
    // A nested table's rows sit inside the outer cell; the outer row then goes on.
    expect(htmlToText("<table><tr><td>L<td><table><tr><td>in</table></td><td>R</tr></table>")).toBe(
      "L\tin\n\n\tR",
    );
  });

  it("keeps a row together when a cell's last break sits inside inline tags (Copilot review on #109)", () => {
    expect(htmlToText("<table><tr><td><span>Total<br></span></td><td>240</td></tr></table>")).toBe(
      "Total\t240",
    );
    expect(htmlToText("<table><tr><td><b><p>A</p></b><td><font>B<br></font></table>")).toBe("A\tB");
    // A break between content inside a cell still counts.
    expect(htmlToText("<table><tr><td><span>A<br></span>B</td></tr></table>")).toBe("A\nB");
  });

  it("treats encoded spacing like spacing before a cell's end (Copilot review on #109)", () => {
    const padding = ["&shy;", "&zwnj;&nbsp;", "&zwnj;&nbsp;&#847;&zwnj;&shy;", "&#xFEFF;"];
    for (const space of ["&nbsp;", "&#160;", "&#32;", "&#xA0;", ...padding]) {
      expect(
        htmlToText(`<table><tr><td>Total<br>${space}</td><td>240</td></tr></table>`),
        space,
      ).toBe("Total\t240");
      expect(htmlToText(`<table><tr><td><p>Total</p>${space}<td>240</table>`), space).toBe(
        "Total\t240",
      );
    }
  });

  it("keeps a <pre> inside a table cell on its row (Copilot review on #109)", () => {
    const row = (cells: string) => htmlToText(`<table><tr>${cells}</tr></table>`);
    expect(row("<td><pre>Total</pre></td><td>240</td>")).toBe("Total\t240");
    expect(row("<td>Total</td><td><pre>240</pre></td>")).toBe("Total\t240");
    expect(row("<td></td><td><pre>x</pre></td>")).toBe("\tx");
    // Its own spacing and line breaks, and content around it in the cell, stay.
    expect(row("<td>A<pre>  b\n  c</pre>D</td><td>E</td>")).toBe("A\n  b\n  c\nD\tE");
  });

  it("treats encoded tabs and newlines as spacing, not table structure (Copilot review on #109)", () => {
    expect(htmlToText("<table><tr><td>Check&#9;in</td><td>Mon</td></tr></table>")).toBe(
      "Check in\tMon",
    );
    expect(htmlToText("<table><tr><td>Check&#10;in</td><td>Mon</td></tr></table>")).toBe(
      "Check in\tMon",
    );
    // A BOM vanishes the same way whether literal or encoded.
    const bom = String.fromCharCode(0xfeff);
    expect(htmlToText(`a${bom}b`)).toBe("ab");
    expect(htmlToText("a&#xFEFF;b")).toBe("ab");
  });

  it("keeps breaks owed outside a cell, such as after a caption (Copilot review on #109)", () => {
    expect(htmlToText("<table><caption>Booking</caption><tr><td>Date<td>Mon</table>")).toBe(
      "Booking\nDate\tMon",
    );
    // An empty first cell keeps the caption's break (Copilot review on #109).
    expect(
      htmlToText("<table><caption>Booking</caption><tr><td></td><td>Mon</td></tr></table>"),
    ).toBe("Booking\n\tMon");
    // A joiner that sits between inline tags inside a word is still kept.
    const zwj = String.fromCharCode(0x200d);
    expect(htmlToText(`<p><b>a</b>&zwj;<b>b</b></p>`)).toBe(`a${zwj}b`);
    expect(htmlToText("<table><tr><td>a</td></tr><p>Note</p></table>After")).toBe(
      "a\nNote\n\nAfter",
    );
  });

  it("isn't fooled by tag-like text inside another tag's attribute (Copilot review on #109)", () => {
    expect(htmlToText('<a title="<!--">Keep me</a><!-- footer -->')).toBe("Keep me");
    expect(htmlToText('<table><tr><td><a title="</td>">A</a></td><td>B</td></tr></table>')).toBe(
      "A\tB",
    );
    expect(htmlToText('<img alt="<style>">Visible')).toBe("Visible");
    expect(htmlToText('<a title="<pre>  x</pre>">Link</a> text')).toBe("Link text");
    expect(htmlToText("<p data-x='</p><p>'>One</p>Two")).toBe("One\nTwo");
  });

  it("shows a `<` that doesn't start a tag as text", () => {
    expect(htmlToText("<p>1 < 2 and 3 > 2</p>")).toBe("1 < 2 and 3 > 2");
    expect(htmlToText("I <3 mail")).toBe("I <3 mail");
  });

  it("ends <pre> and <head> only at a real closing tag (Copilot review on #109)", () => {
    // A </pre> inside an attribute or a comment doesn't end the block.
    expect(htmlToText('<pre><a title="</pre>">  A    B</a></pre>')).toBe("  A    B");
    expect(htmlToText("<pre>  x<!-- </pre> -->  y</pre>After")).toBe("  x  y\nAfter");
    // Nor does a </head> inside an attribute or a comment end the head early.
    expect(htmlToText('<head><meta content="</head>"></head><p>Body</p>')).toBe("Body");
    expect(htmlToText("<head><!-- </head> --><style>x{}</style></head><p>Body</p>")).toBe("Body");
    // The head also ends where the body starts, and an unended <head> hides nothing.
    expect(htmlToText("<head><title>T</title><body><p>Body</p>")).toBe("Body");
    expect(htmlToText("<head><p>Visible</p>")).toBe("Visible");
    // Fake terminators (in a comment or attribute) with no real one hide nothing
    // either (Copilot review on #109).
    expect(htmlToText("<head><!-- </head> --><p>Visible</p>")).toBe("Visible");
    expect(htmlToText('<head><meta content="</head>"><p>Visible</p>')).toBe("Visible");
    expect(htmlToText('<head><meta content="<body>"><title>T</title><p>Visible</p>')).toBe(
      "Visible",
    );
    // An unclosed <pre> keeps its spacing to the end, as in a browser.
    expect(htmlToText("<p>Hi</p><pre>  a\n    b")).toBe("Hi\n\n  a\n    b");
  });
});

// Mail is untrusted and conversion is synchronous: none of these may be
// quadratic (Copilot review on #109). Quadratic versions take tens of seconds.
describe("htmlToText stays fast on hostile input", () => {
  const ZWNJ = String.fromCharCode(0x200c);
  const cases: Record<string, string> = {
    "<br> run with no cell close": "<br>".repeat(100_000),
    "</p> run with no cell close": "</p>".repeat(100_000),
    "unclosed comments": "<!--x".repeat(100_000),
    "unclosed <style>": "<style>x".repeat(100_000),
    "unclosed <pre>": "<pre>x".repeat(100_000),
    "stray <": "<x".repeat(100_000),
    "joiner padding": `${ZWNJ} `.repeat(100_000),
    "joiner run": ZWNJ.repeat(400_000) + "x",
    "long space run in <pre>": `<pre>${" ".repeat(400_000)}x</pre>`,
    "unclosed attribute quotes": '<a title="x'.repeat(100_000),
    "mixed quotes": `<a '"<b "'`.repeat(100_000),
    "implied table cells": "<tr><td>x<td>".repeat(100_000),
    "nested table opens": "<table><td>".repeat(100_000),
    "unclosed declarations": "<!x".repeat(100_000),
    "comment openers in attributes": '<a title="<!--">k</a>'.repeat(100_000),
    "nested <pre> openers": "<pre>".repeat(100_000) + "</pre>",
    "repeated <head> openers": "<head>".repeat(100_000) + "</head>x",
    "unended <head> openers": "<head>x".repeat(100_000),
    "inline-wrapped cell breaks": "<td><span>x<br></span>".repeat(100_000),
  };
  for (const [name, html] of Object.entries(cases)) {
    it(name, () => {
      const t0 = performance.now();
      htmlToText(html);
      expect(performance.now() - t0).toBeLessThan(1_500);
    });
  }
});

describe("htmlToText tables and blocks (#32, continued)", () => {
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
