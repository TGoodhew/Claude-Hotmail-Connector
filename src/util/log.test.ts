import { describe, expect, it, vi } from "vitest";
import { createLogger, redact, redactString, REDACTED } from "./log.js";

describe("redact", () => {
  it("replaces sensitive object keys with the redaction placeholder", () => {
    const input = {
      authorization: "Bearer abc.def.ghi",
      access_token: "secret-token",
      refresh_token: "refresh-me",
      client_secret: "shhh",
      apiKey: "k",
      subject: "Booking confirmation",
      nested: { password: "p", note: "keep me" },
    };
    const out = redact(input) as Record<string, unknown>;

    expect(out.authorization).toBe(REDACTED);
    expect(out.access_token).toBe(REDACTED);
    expect(out.refresh_token).toBe(REDACTED);
    expect(out.client_secret).toBe(REDACTED);
    expect(out.apiKey).toBe(REDACTED);
    // Non-sensitive keys are preserved.
    expect(out.subject).toBe("Booking confirmation");
    expect((out.nested as Record<string, unknown>).password).toBe(REDACTED);
    expect((out.nested as Record<string, unknown>).note).toBe("keep me");
  });

  it("scrubs bearer tokens and JWTs embedded in strings", () => {
    expect(redactString("Authorization: Bearer abc123.def456.ghi789")).not.toContain("abc123");
    expect(redactString("Bearer xyz")).toContain(REDACTED);
    const jwt = "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxIn0.s1gnatureVALUE";
    expect(redactString(`token=${jwt}`)).not.toContain("s1gnatureVALUE");
  });

  it("scrubs secrets embedded in URL query strings", () => {
    const url = "https://host/callback?code=SECRETCODE&state=STATEVAL";
    const out = redactString(url);
    expect(out).not.toContain("SECRETCODE");
    expect(out).not.toContain("STATEVAL");
  });

  it("handles circular references without throwing", () => {
    const a: Record<string, unknown> = { name: "x" };
    a.self = a;
    expect(() => redact(a)).not.toThrow();
    const out = redact(a) as Record<string, unknown>;
    expect(out.self).toBe("[Circular]");
  });

  it("serialises Errors to a safe shape", () => {
    const out = redact(new Error("boom Bearer tok")) as Record<string, unknown>;
    expect(out.name).toBe("Error");
    expect(out.message).not.toContain("tok");
  });
});

describe("createLogger", () => {
  it("writes to the sink and never to stdout", () => {
    const stdoutSpy = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    const lines: string[] = [];
    const log = createLogger({ level: "debug", write: (l) => lines.push(l) });

    log.info("hello", { access_token: "leak-me", ok: 1 });

    expect(stdoutSpy).not.toHaveBeenCalled();
    expect(lines).toHaveLength(1);
    const record = JSON.parse(lines[0]!);
    expect(record.level).toBe("info");
    expect(record.msg).toBe("hello");
    expect(record.ctx.access_token).toBe(REDACTED);
    expect(record.ctx.ok).toBe(1);

    stdoutSpy.mockRestore();
  });

  it("respects the level threshold", () => {
    const lines: string[] = [];
    const log = createLogger({ level: "warn", write: (l) => lines.push(l) });
    log.debug("nope");
    log.info("nope");
    log.warn("yes");
    log.error("yes");
    expect(lines).toHaveLength(2);
  });

  it("the default sink targets stderr, not stdout", () => {
    const stdoutSpy = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    const stderrSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    const log = createLogger({ level: "info" });

    log.error("to stderr");

    expect(stdoutSpy).not.toHaveBeenCalled();
    expect(stderrSpy).toHaveBeenCalledTimes(1);

    stdoutSpy.mockRestore();
    stderrSpy.mockRestore();
  });

  it("merges child bindings into every record", () => {
    const lines: string[] = [];
    const log = createLogger({ level: "info", write: (l) => lines.push(l) }).child({
      correlationId: "abc-123",
    });
    log.info("with binding");
    expect(JSON.parse(lines[0]!).correlationId).toBe("abc-123");
  });
});
