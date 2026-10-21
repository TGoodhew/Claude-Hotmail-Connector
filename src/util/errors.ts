/**
 * Typed error hierarchy for the connector.
 *
 * Errors carry a stable `code` for internal handling. Client-facing surfaces
 * (MCP tool results) should use {@link clientSafeMessage} so that detailed
 * reasons stay in the (redacted) internal logs and are not leaked to the model.
 */

export type ErrorCode =
  "CONFIG_ERROR" | "AUTH_ERROR" | "GRAPH_ERROR" | "VALIDATION_ERROR" | "TIMEOUT_ERROR";

/** Base class for all errors raised inside the connector. */
export class AppError extends Error {
  readonly code: ErrorCode;
  override readonly cause?: unknown;

  constructor(message: string, code: ErrorCode, cause?: unknown) {
    super(message);
    // Preserve the concrete subclass name (ConfigError, AuthError, ...).
    this.name = new.target.name;
    this.code = code;
    this.cause = cause;
    // Maintain a proper stack across the transpiled class hierarchy.
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, new.target);
    }
  }
}

/** Configuration is missing or invalid (e.g. no MICROSOFT_CLIENT_ID). */
export class ConfigError extends AppError {
  constructor(message: string, cause?: unknown) {
    super(message, "CONFIG_ERROR", cause);
  }
}

/** Upstream Microsoft sign-in / token acquisition failed. */
export class AuthError extends AppError {
  constructor(message: string, cause?: unknown) {
    super(message, "AUTH_ERROR", cause);
  }
}

/**
 * The user must act in their browser before the call can succeed: an
 * on-demand sign-in has been started, is still pending, or just failed. Its
 * message is written for the user, contains nothing sensitive, and is returned
 * to the client verbatim.
 */
export class SignInRequiredError extends AuthError {}

/**
 * The user signed out while this request (a sign-in, refresh or tool call) was
 * in progress. It must stop: no fallback sign-in, nothing saved.
 */
export class SignedOutError extends SignInRequiredError {
  constructor() {
    super(
      "You were signed out while this request was in progress, so it wasn't completed. " +
        "Ask me again if you want to sign in.",
    );
  }
}

/**
 * The explicit sign-in path, appended to user-facing auth messages. Extension
 * installs sign in on demand only; source checkouts also have `login`.
 */
export const LOGIN_HINT = "(from a source checkout you can also run `node dist/index.js login`)";

/**
 * A Microsoft Graph request failed. Carries the HTTP status when there was a
 * response, or the `reason` there wasn't one.
 */
export class GraphError extends AppError {
  readonly status?: number;
  readonly graphCode?: string;
  readonly reason?: "network" | "cancelled";

  constructor(
    message: string,
    options: {
      status?: number;
      graphCode?: string;
      reason?: "network" | "cancelled";
      cause?: unknown;
    } = {},
  ) {
    super(message, "GRAPH_ERROR", options.cause);
    this.status = options.status;
    this.graphCode = options.graphCode;
    this.reason = options.reason;
  }
}

// Graph error codes are identifiers (`ErrorItemNotFound`, `Request_ResourceNotFound`).
// Anything else in that field is dropped rather than echoed.
const GRAPH_CODE_RE = /^[A-Za-z][A-Za-z0-9_.]{0,63}$/;

function graphStatusMessage(s: number): string {
  if (s === 400) {
    return (
      "Microsoft Graph rejected the request as invalid (400). Check the inputs, such as ids, " +
      "dates and time zones, and try again."
    );
  }
  if (s === 401) {
    return (
      "Microsoft Graph didn't accept the sign-in (401). Ask me to try again; if it keeps " +
      `happening, sign out and sign in again ${LOGIN_HINT}.`
    );
  }
  if (s === 403) {
    return (
      "Access denied (403). The item isn't accessible to this account, or the connector wasn't " +
      "granted the permission this needs (for example, if MICROSOFT_SCOPES was narrowed). If " +
      "permissions changed, sign out and sign in again to re-consent."
    );
  }
  if (s === 404 || s === 410) {
    return (
      `Not found (${s}). The id may be stale or wrong: list the items again ` +
      "(list_messages or list_events) to get a current id."
    );
  }
  if (s === 409) {
    return "Microsoft Graph reported a conflict (409): the item changed at the same time. Try again.";
  }
  if (s === 413) return "The request was too large for Microsoft Graph (413).";
  if (s === 429) {
    return "Microsoft Graph is limiting requests right now (429). Wait a minute, then try again.";
  }
  if (s >= 500) return `Microsoft Graph had a temporary problem (${s}). Try again shortly.`;
  return `The request to Microsoft Graph failed (${s}).`;
}

/**
 * What to tell the model about a failed Graph call, by status, so it can
 * correct itself (a stale id, a missing permission, bad input) instead of
 * seeing one generic failure (#35). Graph's own message is never included: it
 * can quote request content. Its error code is, when it's a plain identifier.
 */
export function graphErrorMessage(err: GraphError): string {
  if (err.reason === "cancelled") return "The request was cancelled.";
  if (err.status === undefined) {
    return "Couldn't reach Microsoft Graph (network error). Check the connection and try again.";
  }
  const base = graphStatusMessage(err.status);
  const code = err.graphCode && GRAPH_CODE_RE.test(err.graphCode) ? err.graphCode : undefined;
  return code ? `${base} Graph error code: ${code}.` : base;
}

/** Tool input failed schema validation. */
export class ValidationError extends AppError {
  constructor(message: string, cause?: unknown) {
    super(message, "VALIDATION_ERROR", cause);
  }
}

/** A network / Graph call exceeded its timeout budget. */
export class TimeoutError extends AppError {
  constructor(message: string, cause?: unknown) {
    super(message, "TIMEOUT_ERROR", cause);
  }
}

/**
 * Produce a non-sensitive message suitable for returning to the MCP client.
 * Never includes tokens, upstream error bodies, or stack traces.
 */
export function clientSafeMessage(err: unknown): string {
  if (err instanceof ValidationError || err instanceof SignInRequiredError) {
    // Validation messages describe the caller's own input, and sign-in
    // messages are user instructions; both are safe to return as-is.
    return err.message;
  }
  if (err instanceof GraphError) return graphErrorMessage(err);
  if (err instanceof AppError) {
    switch (err.code) {
      case "CONFIG_ERROR":
        return "The connector is not configured correctly. Check the server logs.";
      case "AUTH_ERROR":
        return (
          "Microsoft sign-in is required or has expired. Ask me to try again and a Microsoft " +
          `sign-in window will open in your browser ${LOGIN_HINT}.`
        );
      case "TIMEOUT_ERROR":
        return "The request to Microsoft Graph timed out. Please try again.";
      default:
        return "An unexpected error occurred.";
    }
  }
  return "An unexpected error occurred.";
}
