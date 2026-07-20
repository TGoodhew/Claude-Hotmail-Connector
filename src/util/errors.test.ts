import { describe, expect, it } from "vitest";
import {
  AppError,
  AuthError,
  ConfigError,
  GraphError,
  TimeoutError,
  ValidationError,
  clientSafeMessage,
} from "./errors.js";

describe("clientSafeMessage — client-facing redaction boundary", () => {
  it("redacts a GraphError's upstream message to a generic string", () => {
    // GraphError messages are built from the raw upstream body.error.message,
    // which can carry the account email / item ids — none of it must leak.
    const upstream = "Mailbox tony@hotmail.com item AAMkAGQ... could not be found";
    const msg = clientSafeMessage(
      new GraphError(upstream, { status: 404, graphCode: "ErrorItemNotFound" }),
    );
    expect(msg).toBe("The request to Microsoft Graph could not be completed.");
    expect(msg).not.toContain("tony@hotmail.com");
    expect(msg).not.toContain(upstream);
  });

  it("maps AuthError to a re-authenticate prompt (not the raw reason)", () => {
    expect(clientSafeMessage(new AuthError("invalid_grant: refresh token revoked"))).toBe(
      "Microsoft sign-in is required or has expired. Please re-authenticate.",
    );
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
