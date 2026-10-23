import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createLogFileSink } from "./log-file.js";
import { createLogger, REDACTED, setLogFileSink } from "./log.js";

let dir: string;
const at = () => new Date("2026-10-23T12:00:00.000Z");

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "cht-logfile-"));
});
afterEach(() => {
  setLogFileSink(undefined);
  rmSync(dir, { recursive: true, force: true });
});

const lines = (path: string) => readFileSync(path, "utf8").split("\n").filter(Boolean);

describe("createLogFileSink (#178)", () => {
  it("creates the folder and appends each line with a timestamp and the process id", () => {
    const path = join(dir, "nested", "connector.log");
    const sink = createLogFileSink(path, { now: at });
    sink('{"msg":"one"}');
    sink('{"msg":"two"}');
    expect(lines(path)).toEqual([
      `2026-10-23T12:00:00.000Z [${process.pid}] {"msg":"one"}`,
      `2026-10-23T12:00:00.000Z [${process.pid}] {"msg":"two"}`,
    ]);
  });

  it("never passes the cap: the older half is cut at a line boundary", () => {
    const path = join(dir, "connector.log");
    const maxBytes = 1000;
    const sink = createLogFileSink(path, { maxBytes, now: at });
    for (let i = 0; i < 200; i++) {
      sink(`line ${String(i).padStart(3, "0")}`);
      expect(statSync(path).size).toBeLessThanOrEqual(maxBytes);
    }
    const kept = lines(path);
    // Whole lines only, ending with the newest, and the oldest gone.
    for (const l of kept) expect(l).toMatch(/^\S+ \[\d+\] line \d{3}$/);
    expect(kept.at(-1)).toMatch(/line 199$/);
    expect(kept[0]).not.toMatch(/line 000$/);
    // Numbers stay in order with no gaps: only the oldest lines were dropped.
    const nums = kept.map((l) => Number(l.slice(-3)));
    expect(nums).toEqual(nums.map((_, i) => nums[0]! + i));
  });

  it("cuts a file that was already over the cap when it started", () => {
    const path = join(dir, "connector.log");
    writeFileSync(path, "old line\n".repeat(300)); // 2700 bytes
    createLogFileSink(path, { maxBytes: 1000, now: at })("new");
    expect(statSync(path).size).toBeLessThanOrEqual(1000);
    expect(lines(path).at(-1)).toMatch(/ new$/);
  });

  it("drops a single line too big to fit, rather than pass the cap", () => {
    const path = join(dir, "connector.log");
    const sink = createLogFileSink(path, { maxBytes: 100, now: at });
    sink("x".repeat(200));
    sink("small");
    expect(lines(path)).toHaveLength(1);
    expect(lines(path)[0]).toMatch(/ small$/);
  });

  it("warns once and stops if the file can't be written", () => {
    const warn = vi.fn();
    // A directory where the file should be: appends fail.
    const sink = createLogFileSink(dir, { warn });
    sink("a");
    sink("b");
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]![0]).toMatch(/can't write the log file/);
  });
});

describe("the logger with a log file (#178)", () => {
  it("sends default-sink lines to the file too, redacted", () => {
    const path = join(dir, "connector.log");
    setLogFileSink(createLogFileSink(path, { now: at }));
    const stderr = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    try {
      createLogger({ name: "test" }).warn("hello", { access_token: "leak-me" });
      expect(stderr).toHaveBeenCalledTimes(1); // stderr still gets it
    } finally {
      stderr.mockRestore();
    }
    const [line] = lines(path);
    expect(line).toContain('"msg":"hello"');
    expect(line).toContain(REDACTED);
    expect(line).not.toContain("leak-me");
  });

  it("writes nothing when the log file is off", () => {
    const path = join(dir, "connector.log");
    const stderr = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    try {
      createLogger({ name: "test" }).warn("hello");
    } finally {
      stderr.mockRestore();
    }
    expect(existsSync(path)).toBe(false);
  });
});
