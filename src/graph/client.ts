/**
 * Hardened Microsoft Graph HTTP client.
 *
 * Every Graph module goes through this wrapper, which centralises the common
 * tool rules from the spec (§7): inject the Bearer access token; refresh once on
 * 401 and retry; back off on 429 honouring `Retry-After`, and on 5xx, timeouts
 * and network errors when the request is safe to repeat; follow
 * `@odata.nextLink` pagination up to a cap; enforce a per-request timeout
 * (covering the body) and an optional overall deadline; stop at once when the
 * caller cancels; and never log tokens.
 */

import type { MicrosoftAuth } from "../auth/microsoft.js";
import { GraphError, TimeoutError, ValidationError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";

const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_RETRIES = 3;
const MAX_RETRY_DELAY_MS = 60_000;
const DEFAULT_PAGE_CAP = 100;
/** Extra page-request budget beyond the item cap, bounding pathological pagination. */
const MAX_EXTRA_PAGES = 10;

export interface GraphRequestOptions {
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
  /** JSON body (object). Serialised with application/json. */
  body?: unknown;
  timeoutMs?: number;
  /** External cancellation (e.g. the MCP client cancelled the tool call). */
  signal?: AbortSignal;
  /**
   * Time (ms since the epoch) by which the whole request, retries and backoff
   * included, must finish. Each attempt's timeout and each wait is cut to fit,
   * and a retry that can't fit isn't made (#37).
   */
  deadline?: number;
  /**
   * Sign-out epoch of the operation this request belongs to (see
   * `MicrosoftAuth.signOutEpoch`). Set by {@link scopeToOperation} so every
   * request of one tool call (pages, follow-ups) stops if the user signed out.
   */
  signOutEpoch?: number;
}

export interface GraphResponse<T> {
  status: number;
  data: T;
  headers: Headers;
}

export interface PageOptions {
  /** Maximum items to accumulate across pages (default 100). */
  maxItems?: number;
}

export interface ODataPage<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

export interface GraphClient {
  request<T = unknown>(
    method: string,
    pathOrUrl: string,
    opts?: GraphRequestOptions,
  ): Promise<GraphResponse<T>>;
  get<T = unknown>(pathOrUrl: string, opts?: GraphRequestOptions): Promise<T>;
  post<T = unknown>(path: string, body: unknown, opts?: GraphRequestOptions): Promise<T>;
  patch<T = unknown>(path: string, body: unknown, opts?: GraphRequestOptions): Promise<T>;
  delete(path: string, opts?: GraphRequestOptions): Promise<void>;
  /** Follow @odata.nextLink, aggregating `value` arrays up to a cap. */
  getAllPages<T = unknown>(
    path: string,
    opts?: GraphRequestOptions,
    pageOpts?: PageOptions,
  ): Promise<T[]>;
}

export interface GraphClientDeps {
  auth: Pick<MicrosoftAuth, "getAccessToken"> &
    Partial<Pick<MicrosoftAuth, "signOutEpoch" | "syncSignOuts">>;
  baseUrl: string;
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  logger?: Logger;
  maxRetries?: number;
  defaultTimeoutMs?: number;
  /** Injectable clock for deadlines (defaults to Date.now). */
  now?: () => number;
  /**
   * Start the browser sign-in on demand when no valid token exists, waiting up
   * to `waitMs` for it (see `GetAccessTokenOptions.interactiveWaitMs`). Omit for
   * non-interactive use (CLI diagnostics): a missing sign-in is then an error.
   */
  interactiveSignIn?: { waitMs: number };
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Resolve a request target against the Graph base URL.
 *
 * A full URL (e.g. an `@odata.nextLink`) is only accepted if it is https, on
 * the Graph origin, and under the base path (e.g. `/v1.0/`). The access token is
 * attached to every request, so anything else could hand it to a third party
 * (spec v2 §9.2, T1). This is enforced here so no caller can bypass it.
 */
export function buildUrl(
  baseUrl: string,
  pathOrUrl: string,
  query?: GraphRequestOptions["query"],
): string {
  const base = new URL(baseUrl.replace(/\/?$/, "/"));
  const absolute = /^[a-z][a-z0-9+.-]*:/i.test(pathOrUrl) || pathOrUrl.startsWith("//");
  let url: URL;
  try {
    url = absolute ? new URL(pathOrUrl) : new URL(base.href + pathOrUrl.replace(/^\/+/, ""));
  } catch {
    throw new ValidationError("Refusing a malformed request URL.");
  }
  // Checked on the *parsed* URL for both forms: parsing normalises dot
  // segments, including percent-encoded ones, so a relative `me/../../beta/…`
  // can't climb out of the base path either.
  if (
    url.protocol !== "https:" ||
    url.origin !== base.origin ||
    !url.pathname.startsWith(base.pathname)
  ) {
    throw new ValidationError(
      "Refusing to send a request outside Microsoft Graph (the access token would be exposed).",
    );
  }
  if (query) {
    for (const [k, v] of Object.entries(query)) {
      if (v !== undefined) url.searchParams.set(k, String(v));
    }
  }
  return url.toString();
}

/** Parse a Retry-After header (seconds or HTTP-date) into milliseconds. */
export function retryAfterMs(headerValue: string | null, now = Date.now()): number | undefined {
  if (!headerValue) return undefined;
  const secs = Number(headerValue);
  if (Number.isFinite(secs)) return Math.max(0, Math.min(secs * 1000, MAX_RETRY_DELAY_MS));
  const date = Date.parse(headerValue);
  if (!Number.isNaN(date)) return Math.max(0, Math.min(date - now, MAX_RETRY_DELAY_MS));
  return undefined;
}

function backoffMs(attempt: number): number {
  const base = 500 * 2 ** attempt;
  const jitter = base * 0.2 * Math.random();
  return Math.min(base + jitter, MAX_RETRY_DELAY_MS);
}

/**
 * Whether sending the request again can't do its work twice. GET is safe;
 * PATCH and DELETE are idempotent (and ours carry `If-Match`, so a repeat after
 * a lost success is refused as a version mismatch); a POST only when it carries
 * a Graph `transactionId`, which Graph de-duplicates. The id must be a
 * non-empty string the JSON body will actually contain (an own, enumerable
 * property), and any other method is treated as unsafe.
 */
const REPEATABLE_METHODS = new Set(["GET", "HEAD", "OPTIONS", "PUT", "PATCH", "DELETE"]);

function safeToRepeat(method: string, body: unknown): boolean {
  const m = method.toUpperCase(); // fetch normalises "post" to POST, so must we
  if (REPEATABLE_METHODS.has(m)) return true;
  if (m !== "POST" || typeof body !== "object" || body === null) return false;
  const id = Object.getOwnPropertyDescriptor(body, "transactionId");
  return id?.enumerable === true && typeof id.value === "string" && id.value.length > 0;
}

/**
 * 429 means the request was refused unprocessed, so any request may retry it.
 * A 5xx may come after the work was done, so only a request safe to repeat may.
 */
function isRetriableStatus(status: number, repeatable: boolean): boolean {
  if (status === 429) return true;
  return repeatable && (status === 500 || status === 502 || status === 503 || status === 504);
}

function cancelled(cause?: unknown): GraphError {
  return new GraphError("Graph request was cancelled.", { reason: "cancelled", cause });
}

function outOfTime(): TimeoutError {
  return new TimeoutError("The request ran out of time before Microsoft Graph answered.");
}

/**
 * Wait for `work`, but give up (rejecting) when `signal` aborts or `remainingMs`
 * runs out. `work` itself isn't cancelled: it may be shared with other callers.
 */
async function withinLimits<T>(
  work: Promise<T>,
  signal: AbortSignal | undefined,
  remainingMs: number,
): Promise<T> {
  const timed = Number.isFinite(remainingMs);
  if (!signal && !timed) return work;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort = (): void => undefined;
  const limit = new Promise<never>((_, reject) => {
    if (timed) timer = setTimeout(() => reject(outOfTime()), Math.max(0, remainingMs));
    if (signal) {
      onAbort = () => reject(cancelled());
      if (signal.aborted) onAbort();
      else signal.addEventListener("abort", onAbort, { once: true });
    }
  });
  try {
    return await Promise.race([work, limit]);
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", onAbort);
  }
}

export function createGraphClient(deps: GraphClientDeps): GraphClient {
  const doFetch = deps.fetchImpl ?? fetch;
  const sleep = deps.sleep ?? defaultSleep;
  const log = deps.logger ?? createLogger({ name: "graph" });
  const maxRetries = deps.maxRetries ?? DEFAULT_MAX_RETRIES;
  const defaultTimeoutMs = deps.defaultTimeoutMs ?? DEFAULT_TIMEOUT_MS;
  const now = deps.now ?? Date.now;

  /** The sign-out epoch for an operation starting now, sign-outs elsewhere included (#95). */
  async function currentEpoch(): Promise<number | undefined> {
    if (deps.auth.syncSignOuts) return deps.auth.syncSignOuts();
    return deps.auth.signOutEpoch?.();
  }

  /** Wait `ms`, but stop at once (as a cancellation) if `signal` aborts. */
  async function backOff(ms: number, signal?: AbortSignal): Promise<void> {
    if (!signal) return sleep(ms);
    if (signal.aborted) throw cancelled();
    let onAbort = (): void => undefined;
    const aborted = new Promise<never>((_, reject) => {
      onAbort = () => reject(cancelled());
      signal.addEventListener("abort", onAbort, { once: true });
    });
    try {
      await Promise.race([sleep(ms), aborted]);
    } finally {
      signal.removeEventListener("abort", onAbort);
    }
  }

  interface RawResponse {
    status: number;
    headers: Headers;
    text: string;
  }

  async function fetchOnce(
    method: string,
    url: string,
    token: string,
    opts: GraphRequestOptions,
  ): Promise<RawResponse> {
    const controller = new AbortController();
    const timeoutMs = opts.timeoutMs ?? defaultTimeoutMs;
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const onExternalAbort = (): void => controller.abort();
    if (opts.signal) {
      if (opts.signal.aborted) controller.abort();
      else opts.signal.addEventListener("abort", onExternalAbort, { once: true });
    }

    const headers: Record<string, string> = {
      authorization: `Bearer ${token}`,
      accept: "application/json",
      ...opts.headers,
    };
    let body: string | undefined;
    if (opts.body !== undefined) {
      headers["content-type"] = "application/json";
      body = JSON.stringify(opts.body);
    }

    try {
      const res = await doFetch(url, { method, headers, body, signal: controller.signal });
      // Read the body under the SAME timeout/abort signal so a stalled or
      // trickled response body is bounded too — not just the wait for headers.
      const text = res.status === 204 ? "" : await res.text();
      return { status: res.status, headers: res.headers, text };
    } catch (e) {
      // A caller-initiated cancellation must not be retried or reported as a timeout.
      if (opts.signal?.aborted) throw cancelled(e);
      if (controller.signal.aborted) {
        throw new TimeoutError(`Graph request timed out after ${timeoutMs}ms.`, e);
      }
      throw new GraphError("Network error calling Microsoft Graph.", {
        reason: "network",
        cause: e,
      });
    } finally {
      clearTimeout(timer);
      opts.signal?.removeEventListener("abort", onExternalAbort);
    }
  }

  function toGraphError(res: RawResponse): GraphError {
    let code: string | undefined;
    let message = `Graph request failed with status ${res.status}.`;
    if (res.text) {
      try {
        const body = JSON.parse(res.text) as { error?: { code?: string; message?: string } };
        if (body.error) {
          code = body.error.code;
          if (body.error.message) message = body.error.message;
        }
      } catch {
        // non-JSON body; keep the generic message.
      }
    }
    return new GraphError(message, { status: res.status, graphCode: code });
  }

  function parseBody<T>(res: RawResponse): T {
    if (res.status === 204 || !res.text) return undefined as T;
    return JSON.parse(res.text) as T;
  }

  /**
   * A pre-response failure (timeout or transient network error) worth retrying.
   * A GraphError that carries an HTTP `status` came from a real response and is
   * handled by the status paths instead; a caller cancellation is thrown earlier.
   */
  function isRetriableThrow(e: unknown): boolean {
    return e instanceof TimeoutError || (e instanceof GraphError && e.status === undefined);
  }

  async function request<T>(
    method: string,
    pathOrUrl: string,
    opts: GraphRequestOptions = {},
  ): Promise<GraphResponse<T>> {
    const url = buildUrl(deps.baseUrl, pathOrUrl, opts.query);
    let attempt = 0;
    let refreshed = false;
    // One epoch for the whole request, retries included: a retry or 401
    // recovery that outlives a sign-out must stop, not sign in again.
    const signOutEpoch = opts.signOutEpoch ?? (await currentEpoch());
    const repeatable = safeToRepeat(method, opts.body);
    const remainingMs = (): number =>
      opts.deadline === undefined ? Number.POSITIVE_INFINITY : opts.deadline - now();
    /** Back off before a retry, or report false if the wait wouldn't fit the deadline. */
    const waitToRetry = async (delayMs: number): Promise<boolean> => {
      if (delayMs >= remainingMs()) return false;
      await backOff(delayMs, opts.signal);
      return true;
    };

    /** Stop if the caller cancelled or the deadline passed: before any wait or send. */
    const ensureStillWanted = (): void => {
      if (opts.signal?.aborted) throw cancelled();
      if (remainingMs() <= 0) throw outOfTime();
    };

    for (;;) {
      ensureStillWanted();
      // The token lookup can refresh or wait for a browser sign-in. This call
      // stops waiting for it on cancellation or the deadline, but doesn't abort
      // it: other calls may be sharing the same refresh or sign-in.
      const token = await withinLimits(
        deps.auth.getAccessToken({
          interactive: deps.interactiveSignIn !== undefined,
          ...(deps.interactiveSignIn ? { interactiveWaitMs: deps.interactiveSignIn.waitMs } : {}),
          // Force a refresh after a 401 (token may be revoked despite our clock).
          ...(refreshed ? { minTtlMs: Number.MAX_SAFE_INTEGER } : {}),
          ...(signOutEpoch !== undefined ? { signOutEpoch } : {}),
        }),
        opts.signal,
        remainingMs(),
      );
      ensureStillWanted(); // nothing is sent once the call is no longer wanted

      // Each attempt gets its usual timeout, cut to what's left of the deadline.
      const timeoutMs = Math.max(1, Math.min(opts.timeoutMs ?? defaultTimeoutMs, remainingMs()));
      let res: RawResponse;
      try {
        res = await fetchOnce(method, url, token, { ...opts, timeoutMs });
      } catch (e) {
        // Never retry a caller-initiated cancellation. Timeouts and transient
        // network errors may have reached Graph, so retry them (with the same
        // backoff budget as throttling) only when repeating is safe.
        if (opts.signal?.aborted) throw e;
        if (isRetriableThrow(e) && repeatable && attempt < maxRetries) {
          const delay = backoffMs(attempt);
          if (await waitToRetry(delay)) {
            attempt += 1;
            log.warn("Graph request failed transiently; backed off", {
              attempt,
              delayMs: delay,
              error: e instanceof Error ? e.name : "unknown",
            });
            continue;
          }
        }
        throw e;
      }

      if (res.status === 401 && !refreshed) {
        refreshed = true;
        log.warn("Graph returned 401; refreshing token and retrying once", { url: redactUrl(url) });
        continue;
      }

      if (isRetriableStatus(res.status, repeatable) && attempt < maxRetries) {
        const delay = retryAfterMs(res.headers.get("retry-after")) ?? backoffMs(attempt);
        if (await waitToRetry(delay)) {
          attempt += 1;
          log.warn("Graph busy or failing; backed off", {
            status: res.status,
            attempt,
            delayMs: delay,
          });
          continue;
        }
      }

      if (res.status < 200 || res.status >= 300) throw toGraphError(res);

      return { status: res.status, data: parseBody<T>(res), headers: res.headers };
    }
  }

  return {
    request,
    async get<T>(pathOrUrl: string, opts?: GraphRequestOptions): Promise<T> {
      return (await request<T>("GET", pathOrUrl, opts)).data;
    },
    async post<T>(path: string, body: unknown, opts?: GraphRequestOptions): Promise<T> {
      return (await request<T>("POST", path, { ...opts, body })).data;
    },
    async patch<T>(path: string, body: unknown, opts?: GraphRequestOptions): Promise<T> {
      return (await request<T>("PATCH", path, { ...opts, body })).data;
    },
    async delete(path: string, opts?: GraphRequestOptions): Promise<void> {
      await request<void>("DELETE", path, opts);
    },
    async getAllPages<T>(
      path: string,
      opts: GraphRequestOptions = {},
      pageOpts: PageOptions = {},
    ): Promise<T[]> {
      const cap = pageOpts.maxItems ?? DEFAULT_PAGE_CAP;
      const items: T[] = [];
      let target = path;
      let first = true;
      // Backstop: bound the number of page requests so an empty page that still
      // carries an @odata.nextLink (or a self-referential link) cannot loop forever.
      let pages = 0;
      const maxPages = cap + MAX_EXTRA_PAGES;
      // Every page belongs to the same operation: page 2 after a sign-out stops.
      const signOutEpoch = opts.signOutEpoch ?? (await currentEpoch());
      const epochOpt = signOutEpoch !== undefined ? { signOutEpoch } : {};
      do {
        // On the first request send the full opts; for nextLink pages keep the
        // headers (e.g. Prefer: outlook.timezone), cancellation and time limits
        // but not the baked-in query.
        const pageOptsToSend: GraphRequestOptions = first
          ? { ...opts, ...epochOpt }
          : {
              headers: opts.headers,
              signal: opts.signal,
              timeoutMs: opts.timeoutMs,
              deadline: opts.deadline,
              ...epochOpt,
            };
        const { data } = await request<ODataPage<T>>("GET", target, pageOptsToSend);
        if (data.value) items.push(...data.value);
        const next = data["@odata.nextLink"] ?? "";
        // Stop if Graph hands back the same link it just gave us (would spin).
        if (next === target) break;
        target = next;
        first = false;
        pages += 1;
      } while (target && items.length < cap && pages < maxPages);
      return items.slice(0, cap);
    },
  };
}

/** What every request of one operation (e.g. one tool call) shares. */
export interface OperationScope {
  /** Sign-out epoch at the operation's start (see `GraphRequestOptions.signOutEpoch`). */
  signOutEpoch?: number;
  /** Cancellation of the whole operation. */
  signal?: AbortSignal;
  /** Time (ms since the epoch) by which the whole operation must finish. */
  deadline?: number;
}

/**
 * A view of `graph` whose every request belongs to one operation (e.g. one
 * tool call): pages, follow-up requests and retries all stop if the user signs
 * out mid-operation (rather than starting a new sign-in for a call made before
 * it), if the operation is cancelled, or when its deadline passes.
 */
export function scopeToOperation(graph: GraphClient, scope: OperationScope): GraphClient {
  const defined = Object.fromEntries(
    Object.entries(scope).filter(([, v]) => v !== undefined),
  ) as OperationScope;
  const scoped = (opts?: GraphRequestOptions): GraphRequestOptions => ({ ...defined, ...opts });
  return {
    request: (method, pathOrUrl, opts) => graph.request(method, pathOrUrl, scoped(opts)),
    get: (pathOrUrl, opts) => graph.get(pathOrUrl, scoped(opts)),
    post: (path, body, opts) => graph.post(path, body, scoped(opts)),
    patch: (path, body, opts) => graph.patch(path, body, scoped(opts)),
    delete: (path, opts) => graph.delete(path, scoped(opts)),
    getAllPages: (path, opts, pageOpts) => graph.getAllPages(path, scoped(opts), pageOpts),
  };
}

function redactUrl(url: string): string {
  try {
    const u = new URL(url);
    return u.origin + u.pathname;
  } catch {
    return "[url]";
  }
}
