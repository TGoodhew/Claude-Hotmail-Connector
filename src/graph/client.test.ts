import { describe, expect, it, vi } from "vitest";
import { GraphError, TimeoutError } from "../util/errors.js";
import { createGraphClient, retryAfterMs } from "./client.js";

const BASE = "https://graph.microsoft.com/v1.0";

function json(body: unknown, init: { status?: number; headers?: Record<string, string> } = {}) {
  return new Response(body === undefined ? "" : JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { "content-type": "application/json", ...init.headers },
  });
}

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;
type AuthFn = (opts?: { interactive?: boolean; minTtlMs?: number }) => Promise<string>;

function fetchMock(impl?: FetchFn) {
  return vi.fn<FetchFn>(impl);
}
function makeAuth(token = "AT") {
  return { getAccessToken: vi.fn<AuthFn>(async () => token) };
}

describe("retryAfterMs", () => {
  it("parses seconds", () => {
    expect(retryAfterMs("2")).toBe(2000);
  });
  it("caps very large values", () => {
    expect(retryAfterMs("100000")).toBe(60_000);
  });
  it("parses an HTTP date relative to now", () => {
    const now = 1_000_000;
    const when = new Date(now + 3000).toUTCString();
    expect(retryAfterMs(when, now)).toBeGreaterThanOrEqual(0);
  });
  it("returns undefined for junk / null", () => {
    expect(retryAfterMs(null)).toBeUndefined();
    expect(retryAfterMs("not-a-date")).toBeUndefined();
  });
});

describe("createGraphClient", () => {
  it("injects the Bearer token and returns parsed JSON", async () => {
    const auth = makeAuth("MY-TOKEN");
    const fetchImpl = fetchMock(async () => json({ displayName: "Tony" }));
    const client = createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never });

    const data = await client.get<{ displayName: string }>("/me");
    expect(data.displayName).toBe("Tony");
    const [url, init] = fetchImpl.mock.calls[0]!;
    expect(url).toBe(`${BASE}/me`);
    expect(init?.headers).toMatchObject({ authorization: "Bearer MY-TOKEN" });
  });

  it("appends query params, skipping undefined", async () => {
    const fetchImpl = fetchMock(async () => json({ value: [] }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    await client.get("/me/messages", { query: { $top: 5, $search: undefined, q: "a b" } });
    const url = new URL(fetchImpl.mock.calls[0]![0]);
    expect(url.searchParams.get("$top")).toBe("5");
    expect(url.searchParams.has("$search")).toBe(false);
    expect(url.searchParams.get("q")).toBe("a b");
  });

  it("refreshes once on 401 then retries successfully", async () => {
    const auth = makeAuth();
    const fetchImpl = fetchMock()
      .mockResolvedValueOnce(
        json({ error: { code: "InvalidAuthenticationToken" } }, { status: 401 }),
      )
      .mockResolvedValueOnce(json({ ok: true }));
    const client = createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never });

    const data = await client.get<{ ok: boolean }>("/me");
    expect(data.ok).toBe(true);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    // Second token acquisition forces a refresh via a huge minTtlMs.
    expect(auth.getAccessToken).toHaveBeenCalledTimes(2);
    expect(auth.getAccessToken.mock.calls[1]![0]).toMatchObject({
      minTtlMs: Number.MAX_SAFE_INTEGER,
    });
  });

  it("does not loop forever on repeated 401s", async () => {
    const fetchImpl = fetchMock(async () => json({ error: { code: "x" } }, { status: 401 }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    await expect(client.get("/me")).rejects.toBeInstanceOf(GraphError);
    expect(fetchImpl).toHaveBeenCalledTimes(2); // original + one post-refresh retry
  });

  it("backs off on 429 honouring Retry-After, then succeeds", async () => {
    const sleep = vi.fn(async () => {});
    const fetchImpl = fetchMock()
      .mockResolvedValueOnce(json(undefined, { status: 429, headers: { "retry-after": "1" } }))
      .mockResolvedValueOnce(json({ ok: true }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep,
    });

    const data = await client.get<{ ok: boolean }>("/me");
    expect(data.ok).toBe(true);
    expect(sleep).toHaveBeenCalledWith(1000);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it("gives up after maxRetries on persistent 429", async () => {
    const sleep = vi.fn(async () => {});
    const fetchImpl = fetchMock(async () =>
      json({ error: { code: "TooManyRequests" } }, { status: 429 }),
    );
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep,
      maxRetries: 2,
    });
    await expect(client.get("/me")).rejects.toBeInstanceOf(GraphError);
    // original + 2 retries
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });

  it("maps a non-ok response to a GraphError with status and code", async () => {
    const fetchImpl = fetchMock(async () =>
      json({ error: { code: "ErrorItemNotFound", message: "Not found." } }, { status: 404 }),
    );
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    await expect(client.get("/me/messages/xyz")).rejects.toMatchObject({
      name: "GraphError",
      status: 404,
      graphCode: "ErrorItemNotFound",
    });
  });

  it("follows @odata.nextLink and aggregates pages up to the cap", async () => {
    const fetchImpl = fetchMock()
      .mockResolvedValueOnce(
        json({ value: [1, 2], "@odata.nextLink": `${BASE}/me/messages?$skip=2` }),
      )
      .mockResolvedValueOnce(json({ value: [3] }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });

    const items = await client.getAllPages<number>("/me/messages", {}, { maxItems: 10 });
    expect(items).toEqual([1, 2, 3]);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(fetchImpl.mock.calls[1]![0]).toBe(`${BASE}/me/messages?$skip=2`);
  });

  it("respects the page cap and stops early", async () => {
    const fetchImpl = fetchMock(async () =>
      json({ value: [1, 2, 3], "@odata.nextLink": `${BASE}/next` }),
    );
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    const items = await client.getAllPages<number>("/x", {}, { maxItems: 2 });
    expect(items).toEqual([1, 2]);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it("raises a TimeoutError when the request aborts", async () => {
    const fetchImpl = fetchMock(
      (_url, init) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () =>
            reject(Object.assign(new Error("aborted"), { name: "AbortError" })),
          );
        }),
    );
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    await expect(client.get("/me", { timeoutMs: 5 })).rejects.toBeInstanceOf(TimeoutError);
  });
});
