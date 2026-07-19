/**
 * Hardened Microsoft Graph HTTP client.
 *
 * Every Graph module goes through this wrapper, which centralises the common
 * tool rules from the spec (§7): inject the Bearer access token; refresh once on
 * 401 and retry; back off on 429/503 honouring `Retry-After`; follow
 * `@odata.nextLink` pagination up to a cap; enforce a per-request timeout; and
 * never log tokens.
 */

import type { MicrosoftAuth } from "../auth/microsoft.js";
import { GraphError, TimeoutError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";

const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_RETRIES = 3;
const MAX_RETRY_DELAY_MS = 60_000;
const DEFAULT_PAGE_CAP = 100;

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

interface ODataPage<T> {
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

  async function fetchOnce(
    method: string,
    url: string,
    token: string,
    opts: GraphRequestOptions,
  ): Promise<Response> {
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
      return await doFetch(url, { method, headers, body, signal: controller.signal });
    } catch (e) {
      if (controller.signal.aborted) {
        throw new TimeoutError(`Graph request timed out after ${timeoutMs}ms.`, e);
      }
      throw new GraphError("Network error calling Microsoft Graph.", { cause: e });
    } finally {
      clearTimeout(timer);
      opts.signal?.removeEventListener("abort", onExternalAbort);
    }
  }

  async function toGraphError(res: Response): Promise<GraphError> {
    let code: string | undefined;
    let message = `Graph request failed with status ${res.status}.`;
    try {
      const body = (await res.json()) as { error?: { code?: string; message?: string } };
      if (body.error) {
        code = body.error.code;
        if (body.error.message) message = body.error.message;
      }
    } catch {
      // non-JSON body; keep the generic message.
    }
    return new GraphError(message, { status: res.status, graphCode: code });
  }

  async function parseBody<T>(res: Response): Promise<T> {
    if (res.status === 204) return undefined as T;
    const text = await res.text();
    if (!text) return undefined as T;
    return JSON.parse(text) as T;
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

      const res = await fetchOnce(method, url, token, opts);

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

      if (!res.ok) throw await toGraphError(res);

      const data = await parseBody<T>(res);
      return { status: res.status, data, headers: res.headers };
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
      do {
        // On the first request send the full opts; for nextLink pages keep the
        // headers (e.g. Prefer: outlook.timezone) but not the baked-in query.
        const pageOptsToSend: GraphRequestOptions = first ? opts : { headers: opts.headers };
        const { data } = await request<ODataPage<T>>("GET", target, pageOptsToSend);
        if (data.value) items.push(...data.value);
        target = data["@odata.nextLink"] ?? "";
        first = false;
      } while (target && items.length < cap);
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
