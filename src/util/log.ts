/**
 * Structured logging that writes **only to stderr**.
 *
 * stdout is reserved exclusively for MCP JSON-RPC frames (see the local-hosting
 * spec, Mode A). Writing anything else to stdout would corrupt the protocol, so
 * this logger never touches it.
 *
 * All context objects and messages pass through {@link redact} before being
 * serialised, so tokens, `Authorization` headers, OAuth codes, cookies,
 * secrets, and account identifiers (email/UPN) can never reach the log output.
 */

export type LogLevel = "debug" | "info" | "warn" | "error";

const LEVEL_ORDER: Record<LogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40,
};

/** Placeholder written in place of any redacted value. */
export const REDACTED = "[REDACTED]";

// Key names whose values are always secret (compared case-insensitively).
const SENSITIVE_KEY_EXACT = new Set([
  "authorization",
  "auth",
  "code",
  "code_verifier",
  "code_challenge",
  "client_secret",
  "clientsecret",
  "refresh_token",
  "access_token",
  "id_token",
  "password",
  "passwd",
  "pwd",
  "cookie",
  "set-cookie",
  "state",
  "session",
  "mcp-session-id",
  // Account identifiers (PII) — redacted from logs even though they aren't secrets.
  "account",
  "email",
  "upn",
  "userprincipalname",
  "username",
  "preferred_username",
]);

// Substrings that mark a key as secret regardless of surrounding text.
const SENSITIVE_KEY_SUBSTR = /(token|secret|password|passwd|bearer|api[_-]?key|credential)/i;

function isSensitiveKey(key: string): boolean {
  const k = key.toLowerCase();
  return SENSITIVE_KEY_EXACT.has(k) || SENSITIVE_KEY_SUBSTR.test(k);
}

// Patterns that scrub secret-looking substrings out of free text.
const BEARER_RE = /(Bearer\s+)[A-Za-z0-9._~+/=-]+/gi;
const JWT_RE = /eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g;
const QUERY_SECRET_RE =
  /([?&](?:access_token|refresh_token|id_token|code|client_secret|code_verifier|state)=)[^&\s"']+/gi;

/** Redact secret-looking substrings from a free-text string. */
export function redactString(input: string): string {
  return input
    .replace(BEARER_RE, `$1${REDACTED}`)
    .replace(QUERY_SECRET_RE, `$1${REDACTED}`)
    .replace(JWT_RE, "[REDACTED_JWT]");
}

/**
 * Deep-redact a value: sensitive object keys become {@link REDACTED} and all
 * strings are scrubbed for embedded secrets. Cycles are handled safely.
 */
export function redact(value: unknown, seen = new WeakSet<object>()): unknown {
  if (typeof value === "string") {
    return redactString(value);
  }
  if (value === null || typeof value !== "object") {
    return value;
  }
  if (seen.has(value as object)) {
    return "[Circular]";
  }
  seen.add(value as object);

  if (Array.isArray(value)) {
    return value.map((v) => redact(v, seen));
  }

  // Error -> serialisable, redacted shape.
  if (value instanceof Error) {
    return { name: value.name, message: redactString(value.message) };
  }

  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    out[k] = isSensitiveKey(k) ? REDACTED : redact(v, seen);
  }
  return out;
}

function parseLevel(raw: string | undefined): LogLevel {
  const v = (raw ?? "info").toLowerCase();
  return v === "debug" || v === "info" || v === "warn" || v === "error" ? v : "info";
}

export interface Logger {
  debug(message: string, context?: Record<string, unknown>): void;
  info(message: string, context?: Record<string, unknown>): void;
  warn(message: string, context?: Record<string, unknown>): void;
  error(message: string, context?: Record<string, unknown>): void;
  /** Derive a child logger that adds fixed context (e.g. a correlation id). */
  child(bindings: Record<string, unknown>): Logger;
}

export interface LoggerOptions {
  name?: string;
  level?: LogLevel;
  /** Injectable sink for testing; defaults to stderr. */
  write?: (line: string) => void;
  bindings?: Record<string, unknown>;
}

function defaultWrite(line: string): void {
  // Intentionally stderr — never stdout (reserved for MCP JSON-RPC).
  process.stderr.write(line + "\n");
}

/** Create a structured stderr logger. */
export function createLogger(options: LoggerOptions = {}): Logger {
  const level = options.level ?? parseLevel(process.env.LOG_LEVEL);
  const write = options.write ?? defaultWrite;
  const name = options.name;
  const bindings = options.bindings ?? {};
  const threshold = LEVEL_ORDER[level];

  function emit(lvl: LogLevel, message: string, context?: Record<string, unknown>): void {
    if (LEVEL_ORDER[lvl] < threshold) return;
    const record: Record<string, unknown> = {
      level: lvl,
      msg: redactString(message),
      ...(name ? { logger: name } : {}),
      ...(redact(bindings) as Record<string, unknown>),
    };
    if (context && Object.keys(context).length > 0) {
      record.ctx = redact(context);
    }
    write(JSON.stringify(record));
  }

  return {
    debug: (m, c) => emit("debug", m, c),
    info: (m, c) => emit("info", m, c),
    warn: (m, c) => emit("warn", m, c),
    error: (m, c) => emit("error", m, c),
    child: (extra) => createLogger({ ...options, level, bindings: { ...bindings, ...extra } }),
  };
}

/** A process-wide default logger. */
export const logger = createLogger({ name: "claude-hotmail-connector" });
