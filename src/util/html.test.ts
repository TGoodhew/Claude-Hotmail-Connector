import { describe, expect, it } from "vitest";
import { htmlToText } from "./html.js";

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
