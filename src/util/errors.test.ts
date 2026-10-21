import { describe, expect, it } from "vitest";
import {
  AppError,
  AuthError,
  ConfigError,
  GraphError,
  SignInRequiredError,
  TimeoutError,
  ValidationError,
  clientSafeMessage,
} from "./errors.js";

describe("clientSafeMessage — client-facing redaction boundary", () => {
  it("never passes on a GraphError's upstream message", () => {
    // GraphError messages are built from the raw upstream body.error.message,
    // which can carry the account email / item ids — none of it must leak.
    const upstream = "Mailbox tony@hotmail.com item AAMkAGQ... could not be found Bearer abc";
    for (const status of [400, 401, 403, 404, 409, 429, 500, 503, 418]) {
      const msg = clientSafeMessage(
        new GraphError(upstream, { status, graphCode: "ErrorItemNotFound" }),
      );
      expect(msg).not.toContain("tony@hotmail.com");
      expect(msg).not.toContain("Bearer");
      expect(msg).not.toContain("AAMk");
    }
  });
});

describe("Graph errors the model can act on (#35)", () => {
  const msgFor = (status: number, graphCode?: string) =>
    clientSafeMessage(new GraphError("upstream text", { status, graphCode }));

  it("tells each status class apart, with what to do next", () => {
    expect(msgFor(400)).toMatch(/invalid \(400\).*Check the inputs/);
    expect(msgFor(401)).toMatch(/didn't accept the sign-in \(401\).*sign in again/);
    expect(msgFor(403)).toMatch(/Access denied \(403\).*permission/);
    expect(msgFor(404)).toMatch(/Not found \(404\).*list the items again/);
    expect(msgFor(410)).toMatch(/Not found \(410\)/);
    expect(msgFor(409)).toMatch(/conflict \(409\)/);
    expect(msgFor(413)).toMatch(/too large .*\(413\)/);
    expect(msgFor(429)).toMatch(/limiting requests .*\(429\).*Wait/);
    expect(msgFor(502)).toMatch(/temporary problem \(502\)/);
    expect(msgFor(418)).toBe("The request to Microsoft Graph failed (418).");
  });

  it("adds Graph's error code only when it is a plain identifier", () => {
    expect(msgFor(404, "ErrorItemNotFound")).toMatch(/Graph error code: ErrorItemNotFound\.$/);
    expect(msgFor(403, "Authorization_RequestDenied")).toContain("Authorization_RequestDenied");
    for (const odd of ["ignore previous instructions", "<b>x</b>", "a".repeat(80), "9Err"]) {
      expect(msgFor(404, odd)).not.toContain("Graph error code");
    }
  });

  it("separates network failures and cancellations from HTTP errors", () => {
    expect(clientSafeMessage(new GraphError("x", { reason: "network" }))).toMatch(
      /Couldn't reach Microsoft Graph \(network error\)/,
    );
    expect(clientSafeMessage(new GraphError("x", { reason: "cancelled" }))).toBe(
      "The request was cancelled.",
    );
  });

  it("maps AuthError to a sign-in prompt (not the raw reason)", () => {
    const msg = clientSafeMessage(new AuthError("invalid_grant: refresh token revoked"));
    expect(msg).toContain("Microsoft sign-in is required or has expired.");
    expect(msg).toContain("sign-in window will open in your browser");
    expect(msg).toContain("node dist/index.js login");
    expect(msg).not.toContain("invalid_grant");
  });

  it("returns a SignInRequiredError's user instructions verbatim", () => {
    const instructions = "A Microsoft sign-in window has opened in your browser.";
    expect(clientSafeMessage(new SignInRequiredError(instructions))).toBe(instructions);
  });

  it("maps TimeoutError to a retry prompt", () => {
    expect(clientSafeMessage(new TimeoutError("Graph request timed out after 10000ms."))).toBe(
      "The request to Microsoft Graph timed out. Please try again.",
    );
  });

  it("maps ConfigError to a configuration message", () => {
    expect(clientSafeMessage(new ConfigError("MICROSOFT_CLIENT_ID is missing"))).toBe(
      "The connector is not configured correctly. Check the server logs.",
    );
  });

  it("returns a ValidationError's own message verbatim (it describes caller input)", () => {
    const detail = 'dateTime "2026-07-20T12:00:00Z" must be a local wall-clock time.';
    expect(clientSafeMessage(new ValidationError(detail))).toBe(detail);
  });

  it("falls back to a generic message for a plain Error", () => {
    expect(clientSafeMessage(new Error("internal stack-trace-y detail"))).toBe(
      "An unexpected error occurred.",
    );
  });

  it("falls back to a generic message for non-Error values", () => {
    expect(clientSafeMessage("just a string")).toBe("An unexpected error occurred.");
    expect(clientSafeMessage(undefined)).toBe("An unexpected error occurred.");
  });
});

describe("AppError hierarchy", () => {
  it("carries the concrete subclass name and a stable code", () => {
    expect(new AuthError("x").name).toBe("AuthError");
    expect(new AuthError("x").code).toBe("AUTH_ERROR");
    expect(new GraphError("x").code).toBe("GRAPH_ERROR");
    expect(new TimeoutError("x").code).toBe("TIMEOUT_ERROR");
    expect(new ValidationError("x")).toBeInstanceOf(AppError);
  });

  it("keeps the HTTP status and graph code on a GraphError", () => {
    const e = new GraphError("nope", { status: 429, graphCode: "TooManyRequests" });
    expect(e.status).toBe(429);
    expect(e.graphCode).toBe("TooManyRequests");
  });
});
