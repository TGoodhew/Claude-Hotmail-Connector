import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
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
    // A directory where the file should be: appends fail (on Windows possibly
    // as a "transient" EPERM, which is given up on after a run of them).
    const sink = createLogFileSink(dir, { warn });
    for (let i = 0; i < 60; i++) sink(`line ${i}`);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]![0]).toMatch(/can't write the log file/);
  });

  it("warns once and stops if a cut fails, without passing the cap", () => {
    const path = join(dir, "connector.log");
    writeFileSync(path, "old line\n".repeat(20)); // 180 bytes, over the cap below
    // A directory where the cut writes its temporary file: the cut fails.
    mkdirSync(`${path}.${process.pid}.tmp`);
    const warn = vi.fn();
    const sink = createLogFileSink(path, { maxBytes: 100, warn });
    for (let i = 0; i < 60; i++) sink(`line ${i}`);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(readFileSync(path, "utf8")).toBe("old line\n".repeat(20)); // nothing appended
  });

  it("drops a line while another process holds the lock, and leaves its lock alone", () => {
    const path = join(dir, "connector.log");
    writeFileSync(`${path}.lock`, "");
    createLogFileSink(path, { now: at })("dropped");
    expect(existsSync(path)).toBe(false);
    expect(existsSync(`${path}.lock`)).toBe(true);
  });

  it("breaks a stale lock left by a process that stopped mid-write", () => {
    const path = join(dir, "connector.log");
    writeFileSync(`${path}.lock`, "");
    const old = new Date(Date.now() - 60_000);
    utimesSync(`${path}.lock`, old, old);
    createLogFileSink(path, { now: at })("written");
    expect(lines(path)[0]).toMatch(/ written$/);
    expect(existsSync(`${path}.lock`)).toBe(false);
  });

  it("keeps the cap with two writers sharing the file", () => {
    const path = join(dir, "connector.log");
    const a = createLogFileSink(path, { maxBytes: 1000, now: at });
    const b = createLogFileSink(path, { maxBytes: 1000, now: at });
    for (let i = 0; i < 100; i++) {
      (i % 2 ? a : b)(`writer ${i % 2 ? "a" : "b"} ${i}`);
      expect(statSync(path).size).toBeLessThanOrEqual(1000);
    }
    expect(lines(path).at(-1)).toMatch(/writer a 99$/);
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

  it("keeps email addresses out of the file, even inside error text", () => {
    const path = join(dir, "connector.log");
    setLogFileSink(createLogFileSink(path, { now: at }));
    const stderr = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    try {
      createLogger({ name: "mcp" }).error("Tool handler failed for tony@hotmail.com", {
        tool: "get_message",
        error: "Mailbox tony@hotmail.com item could not be found",
        account: "tony@hotmail.com",
      });
    } finally {
      stderr.mockRestore();
    }
    const text = readFileSync(path, "utf8");
    expect(text).toContain("Tool handler failed");
    expect(text).not.toContain("tony@hotmail.com");
    expect(text).toContain("[REDACTED_EMAIL]");
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
