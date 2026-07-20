/**
 * Hardened Microsoft Graph HTTP client.
 *
 * Every Graph module goes through this wrapper, which centralises the common
 * tool rules from the spec (§7): inject the Bearer access token; refresh once on
 * 401 and retry; back off on 429/503/504 honouring `Retry-After`; retry timeouts
 * and transient network errors with the same budget; follow `@odata.nextLink`
 * pagination up to a cap; enforce a per-request timeout (covering the body); and
 * never log tokens.
 */

import type { MicrosoftAuth } from "../auth/microsoft.js";
import { GraphError, TimeoutError } from "../util/errors.js";
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
  /** External cancellation. */
  signal?: AbortSignal;
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
  auth: Pick<MicrosoftAuth, "getAccessToken">;
  baseUrl: string;
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  logger?: Logger;
  maxRetries?: number;
  defaultTimeoutMs?: number;
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function buildUrl(
  baseUrl: string,
  pathOrUrl: string,
  query?: GraphRequestOptions["query"],
): string {
  // A full URL (e.g. an @odata.nextLink) is used verbatim.
  const url = /^https?:\/\//i.test(pathOrUrl)
    ? new URL(pathOrUrl)
    : new URL(baseUrl.replace(/\/$/, "") + "/" + pathOrUrl.replace(/^\//, ""));
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

function isRetriableStatus(status: number): boolean {
  return status === 429 || status === 503 || status === 504;
}

export function createGraphClient(deps: GraphClientDeps): GraphClient {
  const doFetch = deps.fetchImpl ?? fetch;
  const sleep = deps.sleep ?? defaultSleep;
  const log = deps.logger ?? createLogger({ name: "graph" });
  const maxRetries = deps.maxRetries ?? DEFAULT_MAX_RETRIES;
  const defaultTimeoutMs = deps.defaultTimeoutMs ?? DEFAULT_TIMEOUT_MS;

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
      if (opts.signal?.aborted) {
        throw new GraphError("Graph request was cancelled.", { cause: e });
      }
      if (controller.signal.aborted) {
        throw new TimeoutError(`Graph request timed out after ${timeoutMs}ms.`, e);
      }
      throw new GraphError("Network error calling Microsoft Graph.", { cause: e });
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

    for (;;) {
      const token = await deps.auth.getAccessToken({
        interactive: false,
        // Force a refresh after a 401 (token may be revoked despite our clock).
        ...(refreshed ? { minTtlMs: Number.MAX_SAFE_INTEGER } : {}),
      });

      let res: RawResponse;
      try {
        res = await fetchOnce(method, url, token, opts);
      } catch (e) {
        // Never retry a caller-initiated cancellation; retry timeouts and
        // transient network errors with the same backoff budget as throttling.
        if (opts.signal?.aborted) throw e;
        if (isRetriableThrow(e) && attempt < maxRetries) {
          const delay = backoffMs(attempt);
          attempt += 1;
          log.warn("Graph request failed transiently; backing off", {
            attempt,
            delayMs: delay,
            error: e instanceof Error ? e.name : "unknown",
          });
          await sleep(delay);
          continue;
        }
        throw e;
      }

      if (res.status === 401 && !refreshed) {
        refreshed = true;
        log.warn("Graph returned 401; refreshing token and retrying once", { url: redactUrl(url) });
        continue;
      }

      if (isRetriableStatus(res.status) && attempt < maxRetries) {
        const delay = retryAfterMs(res.headers.get("retry-after")) ?? backoffMs(attempt);
        attempt += 1;
        log.warn("Graph throttled; backing off", { status: res.status, attempt, delayMs: delay });
        await sleep(delay);
        continue;
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
      do {
        // On the first request send the full opts; for nextLink pages keep the
        // headers (e.g. Prefer: outlook.timezone) but not the baked-in query.
        const pageOptsToSend: GraphRequestOptions = first ? opts : { headers: opts.headers };
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

function redactUrl(url: string): string {
  try {
    const u = new URL(url);
    return u.origin + u.pathname;
  } catch {
    return "[url]";
  }
}
