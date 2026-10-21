/**
 * Fail if any tracked text file contains a raw control character (other than
 * tab, LF and CR).
 *
 * A literal NUL in source makes git classify the file as binary, which hides
 * its diffs from review (see #36). Write such characters as escape sequences
 * instead, e.g. "\u0000".
 *
 * Run:  npm run check:text
 */
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { extname } from "node:path";

// Tracked files that are legitimately binary.
const BINARY_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".gif", ".ico", ".exe", ".blob"]);

// Control characters except \t (0x09), \n (0x0A) and \r (0x0D).
const CONTROL_RE = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g; // eslint-disable-line no-control-regex

const files = execFileSync("git", ["ls-files", "-z"], { encoding: "utf8" })
  .split("\0")
  .filter((f) => f && !BINARY_EXTENSIONS.has(extname(f).toLowerCase()));

const problems = [];
for (const file of files) {
  const lines = readFileSync(file, "utf8").split("\n");
  lines.forEach((line, i) => {
    for (const m of line.matchAll(CONTROL_RE)) {
      const code = m[0].charCodeAt(0).toString(16).padStart(4, "0");
      problems.push(`${file}:${i + 1}:${m.index + 1}  U+${code}`);
    }
  });
}

if (problems.length > 0) {
  console.error('Raw control characters found (use an escape sequence such as "\\u0000"):');
  for (const p of problems) console.error(`  ${p}`);
  process.exit(1);
}
console.log(`check:text: ${files.length} tracked text files clean.`);
