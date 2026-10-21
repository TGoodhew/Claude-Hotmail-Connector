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
 * The explicit sign-in path, appended to user-facing auth messages. Extension
 * installs sign in on demand only; source checkouts also have `login`.
 */
export const LOGIN_HINT = "(from a source checkout you can also run `node dist/index.js login`)";

/** A Microsoft Graph request failed. Carries the HTTP status when known. */
export class GraphError extends AppError {
  readonly status?: number;
  readonly graphCode?: string;

  constructor(
    message: string,
    options: { status?: number; graphCode?: string; cause?: unknown } = {},
  ) {
    super(message, "GRAPH_ERROR", options.cause);
    this.status = options.status;
    this.graphCode = options.graphCode;
  }
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
 * Produce a generic, non-sensitive message suitable for returning to the MCP
 * client. Never includes tokens, upstream error bodies, or stack traces.
 */
export function clientSafeMessage(err: unknown): string {
  if (err instanceof ValidationError || err instanceof SignInRequiredError) {
    // Validation messages describe the caller's own input, and sign-in
    // messages are user instructions; both are safe to return as-is.
    return err.message;
  }
  if (err instanceof AppError) {
    switch (err.code) {
      case "CONFIG_ERROR":
        return "The connector is not configured correctly. Check the server logs.";
      case "AUTH_ERROR":
        return (
          "Microsoft sign-in is required or has expired. Ask me to try again and a Microsoft " +
          `sign-in window will open in your browser ${LOGIN_HINT}.`
        );
      case "GRAPH_ERROR":
        return "The request to Microsoft Graph could not be completed.";
      case "TIMEOUT_ERROR":
        return "The request to Microsoft Graph timed out. Please try again.";
      default:
        return "An unexpected error occurred.";
    }
  }
  return "An unexpected error occurred.";
}
