import { describe, expect, it, vi } from "vitest";
import { GraphError, TimeoutError, ValidationError } from "../util/errors.js";
import {
  buildUrl,
  createGraphClient,
  retryAfterMs,
  scopeToOperation,
  type GraphClient,
} from "./client.js";

const BASE = "https://graph.microsoft.com/v1.0";

describe("buildUrl origin guard (#30, spec v2 §9.2 T1)", () => {
  it("resolves relative paths under the Graph base", () => {
    expect(buildUrl(BASE, "/me/messages")).toBe(`${BASE}/me/messages`);
    expect(buildUrl(BASE, "me/messages?$skip=10")).toBe(`${BASE}/me/messages?$skip=10`);
  });

  it("accepts Graph's own absolute nextLinks", () => {
    expect(buildUrl(BASE, `${BASE}/me/messages?$skiptoken=x`)).toBe(
      `${BASE}/me/messages?$skiptoken=x`,
    );
  });

  it("refuses anything that would send the token elsewhere", () => {
    for (const bad of [
      "https://evil.example/steal",
      "http://graph.microsoft.com/v1.0/me", // not https
      "https://graph.microsoft.com.evil.example/v1.0/me", // lookalike host
      "https://graph.microsoft.com@evil.example/v1.0/me", // userinfo trick
      "https://graph.microsoft.com/beta/me", // outside the base path
      "//evil.example/v1.0/me", // scheme-relative
      "javascript:alert(1)",
      "file:///etc/passwd",
    ]) {
      expect(() => buildUrl(BASE, bad), bad).toThrow(ValidationError);
    }
  });

  it("refuses relative paths that climb out of the base path (literal or percent-encoded)", () => {
    for (const bad of [
      "me/../../beta/me/messages",
      "../beta/me",
      "me/%2e%2e/%2e%2e/beta/me/messages",
      "me/.%2E/.%2e/beta/me",
    ]) {
      expect(() => buildUrl(BASE, bad), bad).toThrow(ValidationError);
    }
    // Staying inside the base is fine.
    expect(buildUrl(BASE, "me/x/../messages")).toBe(`${BASE}/me/messages`);
  });

  it("is enforced on the request path (no fetch for a foreign URL)", async () => {
    const fetchImpl = vi.fn();
    const client = createGraphClient({
      auth: { getAccessToken: async () => "AT" },
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    await expect(client.get("https://evil.example/x")).rejects.toBeInstanceOf(ValidationError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

function json(body: unknown, init: { status?: number; headers?: Record<string, string> } = {}) {
  return new Response(body === undefined ? "" : JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { "content-type": "application/json", ...init.headers },
  });
}

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;
type AuthFn = (opts?: {
  interactive?: boolean;
  interactiveWaitMs?: number;
  minTtlMs?: number;
  signOutEpoch?: number;
}) => Promise<string>;

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
    expect(retryAfterMs(when, now)).toBe(3000);
  });
  it("floors a past HTTP date at 0", () => {
    const now = 2_000_000;
    const past = new Date(now - 5000).toUTCString();
    expect(retryAfterMs(past, now)).toBe(0);
  });
  it("caps a far-future HTTP date at the max delay", () => {
    const now = 1_000_000;
    const far = new Date(now + 10 * 60_000).toUTCString();
    expect(retryAfterMs(far, now)).toBe(60_000);
  });
  it("returns undefined for junk / null", () => {
    expect(retryAfterMs(null)).toBeUndefined();
    expect(retryAfterMs("not-a-date")).toBeUndefined();
  });
});

describe("createGraphClient token acquisition", () => {
  it("passes one sign-out epoch to every lookup of a request, retries included", async () => {
    const auth = { ...makeAuth(), signOutEpoch: vi.fn(() => 7) };
    const fetchImpl = fetchMock()
      .mockResolvedValueOnce(
        json({ error: { code: "InvalidAuthenticationToken" } }, { status: 401 }),
      )
      .mockResolvedValueOnce(json({ ok: true }));
    const client = createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never });
    await client.get("/me");
    expect(auth.signOutEpoch).toHaveBeenCalledOnce(); // taken once, at the start
    expect(auth.getAccessToken.mock.calls.map(([o]) => o?.signOutEpoch)).toEqual([7, 7]);
  });

  it("scopeToOperation pins every request (and page) to the given epoch", async () => {
    let epoch = 1;
    const auth = { ...makeAuth(), signOutEpoch: vi.fn(() => epoch) };
    const fetchImpl = fetchMock()
      .mockResolvedValueOnce(json({ value: [1], "@odata.nextLink": `${BASE}/me/messages?p=2` }))
      .mockResolvedValueOnce(json({ value: [2] }))
      .mockResolvedValueOnce(json({ ok: true }));
    const scoped = scopeToOperation(
      createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never }),
      { signOutEpoch: 1 },
    );
    await scoped.getAllPages("/me/messages");
    epoch = 2; // signed out between this operation's requests
    await scoped.patch("/me/events/e", {});
    expect(auth.getAccessToken.mock.calls.map(([o]) => o?.signOutEpoch)).toEqual([1, 1, 1]);
  });

  it("is non-interactive by default (CLI diagnostics never open a browser)", async () => {
    const auth = makeAuth();
    const client = createGraphClient({
      auth,
      baseUrl: BASE,
      fetchImpl: fetchMock(async () => json({ ok: true })) as never,
    });
    await client.get("/me");
    expect(auth.getAccessToken.mock.calls[0]![0]).toMatchObject({ interactive: false });
  });

  it("requests on-demand sign-in with the configured wait for tool calls", async () => {
    const auth = makeAuth();
    const client = createGraphClient({
      auth,
      baseUrl: BASE,
      fetchImpl: fetchMock(async () => json({ ok: true })) as never,
      interactiveSignIn: { waitMs: 45_000 },
    });
    await client.get("/me");
    expect(auth.getAccessToken.mock.calls[0]![0]).toMatchObject({
      interactive: true,
      interactiveWaitMs: 45_000,
    });
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

  it("retries then raises a TimeoutError when every attempt times out", async () => {
    const sleep = vi.fn(async () => {});
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
      sleep,
      maxRetries: 2,
    });
    await expect(client.get("/me", { timeoutMs: 5 })).rejects.toBeInstanceOf(TimeoutError);
    expect(fetchImpl).toHaveBeenCalledTimes(3); // original + 2 retries
  });

  it("does not retry a caller-cancelled request", async () => {
    const sleep = vi.fn(async () => {});
    const external = new AbortController();
    const fetchImpl = fetchMock(
      (_url, init) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () =>
            reject(Object.assign(new Error("aborted"), { name: "AbortError" })),
          );
          external.abort();
        }),
    );
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep,
    });
    await expect(client.get("/me", { signal: external.signal })).rejects.toBeInstanceOf(GraphError);
    expect(fetchImpl).toHaveBeenCalledTimes(1); // cancelled → no retry
    expect(sleep).not.toHaveBeenCalled();
  });

  it("retries a transient network error then succeeds", async () => {
    const sleep = vi.fn(async () => {});
    const fetchImpl = fetchMock()
      .mockRejectedValueOnce(new Error("ECONNRESET"))
      .mockResolvedValueOnce(json({ ok: true }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep,
    });
    const data = await client.get<{ ok: boolean }>("/me");
    expect(data.ok).toBe(true);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(sleep).toHaveBeenCalledTimes(1);
  });

  it("retries 503 and 504 like 429", async () => {
    const sleep = vi.fn(async () => {});
    for (const status of [503, 504]) {
      const fetchImpl = fetchMock()
        .mockResolvedValueOnce(json(undefined, { status }))
        .mockResolvedValueOnce(json({ ok: true }));
      const client = createGraphClient({
        auth: makeAuth(),
        baseUrl: BASE,
        fetchImpl: fetchImpl as never,
        sleep,
      });
      const data = await client.get<{ ok: boolean }>("/me");
      expect(data.ok, `status ${status}`).toBe(true);
      expect(fetchImpl, `status ${status}`).toHaveBeenCalledTimes(2);
    }
  });

  it("uses exponential backoff (bounded) when no Retry-After is present", async () => {
    const delays: number[] = [];
    const sleep = vi.fn(async (ms: number) => {
      delays.push(ms);
    });
    const fetchImpl = fetchMock(async () => json({ error: { code: "x" } }, { status: 503 }));
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep,
      maxRetries: 3,
    });
    await expect(client.get("/me")).rejects.toBeInstanceOf(GraphError);
    expect(delays).toHaveLength(3);
    // base grows as 500 * 2^attempt (with jitter) and never exceeds the 60s cap
    expect(delays[0]).toBeLessThan(delays[1]!);
    expect(delays[1]).toBeLessThan(delays[2]!);
    for (const d of delays) expect(d).toBeLessThanOrEqual(60_000);
  });

  it("bounds pagination when Graph returns empty pages carrying a nextLink", async () => {
    let calls = 0;
    const fetchImpl = fetchMock(async () => {
      calls += 1;
      // Always-empty page with an ever-changing nextLink would loop forever
      // without the page-count backstop.
      return json({ value: [], "@odata.nextLink": `${BASE}/next?skip=${calls}` });
    });
    const client = createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
    });
    const items = await client.getAllPages<number>("/x", {}, { maxItems: 3 });
    expect(items).toEqual([]);
    // cap (3) + MAX_EXTRA_PAGES (10) = 13 page requests max.
    expect(calls).toBeLessThanOrEqual(13);
    expect(calls).toBeGreaterThan(3);
  });
});

describe("retry policy, deadlines and cancellation (#37)", () => {
  const client = (fetchImpl: ReturnType<typeof fetchMock>, extra: Record<string, unknown> = {}) =>
    createGraphClient({
      auth: makeAuth(),
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      sleep: async () => {},
      ...extra,
    });

  it("retries 500 and 502 for requests that are safe to repeat", async () => {
    for (const status of [500, 502]) {
      const fetchImpl = fetchMock()
        .mockResolvedValueOnce(json({ error: { code: "x" } }, { status }))
        .mockResolvedValueOnce(json({ ok: true }));
      await expect(client(fetchImpl).get("/me")).resolves.toEqual({ ok: true });
      expect(fetchImpl).toHaveBeenCalledTimes(2);
    }
  });

  it("never repeats a POST without a transactionId after a 5xx or timeout", async () => {
    const failing = fetchMock(async () => json({ error: { code: "x" } }, { status: 503 }));
    await expect(client(failing).post("/me/sendMail", { a: 1 })).rejects.toMatchObject({
      status: 503,
    });
    expect(failing).toHaveBeenCalledOnce();

    const timingOut = fetchMock(
      (_url, init) =>
        new Promise((_, reject) =>
          init?.signal?.addEventListener("abort", () => reject(new Error("aborted"))),
        ),
    );
    await expect(
      client(timingOut, { defaultTimeoutMs: 5 }).post("/me/x", { a: 1 }),
    ).rejects.toBeInstanceOf(TimeoutError);
    expect(timingOut).toHaveBeenCalledOnce();
  });

  it("doesn't repeat a POST whose dedup key won't be sent, however it's spelt (Copilot review on #120)", async () => {
    const inherited = Object.create({ transactionId: "t" }) as object;
    const hidden = Object.defineProperty({}, "transactionId", { value: "t", enumerable: false });
    const cases: [string, unknown][] = [
      ["post", { a: 1 }], // fetch sends lowercase "post" as POST
      ["Post", { a: 1 }],
      ["POST", { subject: "x", transactionId: undefined }], // dropped by JSON.stringify
      ["POST", { transactionId: "" }],
      ["POST", { transactionId: 42 }],
      ["POST", inherited],
      ["POST", hidden],
      ["PROPFIND", undefined], // an unknown method isn't assumed safe
    ];
    for (const [method, body] of cases) {
      const fetchImpl = fetchMock(async () => json({}, { status: 503 }));
      await expect(client(fetchImpl).request(method, "/me/x", { body })).rejects.toMatchObject({
        status: 503,
      });
      expect(fetchImpl, `${method} ${JSON.stringify(body)}`).toHaveBeenCalledOnce();
    }
    // Lowercase is still recognised as safe for idempotent methods.
    const get = fetchMock()
      .mockResolvedValueOnce(json({}, { status: 503 }))
      .mockResolvedValueOnce(json({ ok: true }));
    await client(get).request("get", "/me");
    expect(get).toHaveBeenCalledTimes(2);
  });

  it("does repeat a POST that Graph de-duplicates (transactionId), and any 429", async () => {
    const deduped = fetchMock()
      .mockResolvedValueOnce(json({}, { status: 503 }))
      .mockResolvedValueOnce(json({ id: "e1" }));
    await client(deduped).post("/me/events", { transactionId: "t" });
    expect(deduped).toHaveBeenCalledTimes(2);

    const throttled = fetchMock()
      .mockResolvedValueOnce(json({}, { status: 429, headers: { "retry-after": "1" } }))
      .mockResolvedValueOnce(json({}));
    await client(throttled).post("/me/x", { a: 1 });
    expect(throttled).toHaveBeenCalledTimes(2);
  });

  it("doesn't wait past the deadline: a Retry-After that won't fit ends the request", async () => {
    let clock = 1_000_000;
    const sleep = vi.fn(async (ms: number) => {
      clock += ms;
    });
    const fetchImpl = fetchMock(async () =>
      json(
        { error: { code: "TooManyRequests" } },
        { status: 429, headers: { "retry-after": "30" } },
      ),
    );
    const graph = client(fetchImpl, { sleep, now: () => clock });
    await expect(graph.get("/me", { deadline: clock + 10_000 })).rejects.toMatchObject({
      status: 429,
    });
    expect(sleep).not.toHaveBeenCalled();
    expect(fetchImpl).toHaveBeenCalledOnce();
  });

  it("fails without calling Graph once the deadline has passed", async () => {
    const fetchImpl = fetchMock(async () => json({}));
    await expect(
      client(fetchImpl, { now: () => 5_000 }).get("/me", { deadline: 4_999 }),
    ).rejects.toBeInstanceOf(TimeoutError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("cuts an attempt's timeout to what's left of the deadline", async () => {
    const fetchImpl = fetchMock(
      (_url, init) =>
        new Promise((_, reject) =>
          init?.signal?.addEventListener("abort", () => reject(new Error("aborted"))),
        ),
    );
    const started = Date.now();
    await expect(
      client(fetchImpl, { defaultTimeoutMs: 10_000, maxRetries: 0 }).get("/me", {
        deadline: Date.now() + 50,
      }),
    ).rejects.toBeInstanceOf(TimeoutError);
    expect(Date.now() - started).toBeLessThan(2_000); // not the 10s default
  });

  it("stops backing off at once when the caller cancels", async () => {
    const controller = new AbortController();
    const fetchImpl = fetchMock(async () => {
      queueMicrotask(() => controller.abort());
      return json({}, { status: 503 });
    });
    const graph = client(fetchImpl, { sleep: () => new Promise(() => {}) }); // waits forever
    await expect(graph.get("/me", { signal: controller.signal })).rejects.toMatchObject({
      reason: "cancelled",
    });
    expect(fetchImpl).toHaveBeenCalledOnce();
  });

  it("keeps cancellation and the deadline on later pages", async () => {
    const controller = new AbortController();
    const fetchImpl = fetchMock(async () => {
      controller.abort(); // cancelled while page 1 was in flight
      return json({ value: [1], "@odata.nextLink": `${BASE}/me/messages?p=2` });
    });
    await expect(
      client(fetchImpl).getAllPages("/me/messages", { signal: controller.signal }),
    ).rejects.toMatchObject({ reason: "cancelled" });
    expect(fetchImpl).toHaveBeenCalledOnce(); // page 2 never requested

    let clock = 0;
    const late = fetchMock()
      .mockImplementationOnce(async () => {
        clock = 200; // page 1 took us past the deadline
        return json({ value: [1], "@odata.nextLink": `${BASE}/me/messages?p=2` });
      })
      .mockResolvedValueOnce(json({ value: [2] }));
    const graph = client(late, { now: () => clock });
    await expect(graph.getAllPages("/me/messages", { deadline: 100 })).rejects.toBeInstanceOf(
      TimeoutError,
    );
    expect(late).toHaveBeenCalledOnce();
  });

  it("stops waiting for a slow sign-in or refresh on cancellation, sending nothing (Copilot review on #120)", async () => {
    const controller = new AbortController();
    let finishSignIn = (_t: string): void => undefined;
    const auth = {
      getAccessToken: vi.fn(() => new Promise<string>((resolve) => (finishSignIn = resolve))),
    };
    const fetchImpl = fetchMock(async () => json({}));
    const graph = createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never });
    const call = graph.patch("/me/events/e", {}, { signal: controller.signal });
    controller.abort();
    await expect(call).rejects.toMatchObject({ reason: "cancelled" });
    finishSignIn("AT"); // the shared sign-in still finishes, for others
    await new Promise((r) => setTimeout(r, 10));
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("gives up on token acquisition at the deadline, sending nothing (Copilot review on #120)", async () => {
    let finishSignIn = (_t: string): void => undefined;
    const auth = {
      getAccessToken: vi.fn(() => new Promise<string>((resolve) => (finishSignIn = resolve))),
    };
    const fetchImpl = fetchMock(async () => json({}));
    const graph = createGraphClient({ auth, baseUrl: BASE, fetchImpl: fetchImpl as never });
    await expect(
      graph.post("/me/events", { transactionId: "t" }, { deadline: Date.now() + 30 }),
    ).rejects.toBeInstanceOf(TimeoutError);
    finishSignIn("AT");
    await new Promise((r) => setTimeout(r, 10));
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("sends nothing when the token arrives after the deadline", async () => {
    let clock = 0;
    const auth = {
      getAccessToken: vi.fn(async () => {
        clock = 500; // the lookup itself outlasted the deadline
        return "AT";
      }),
    };
    const fetchImpl = fetchMock(async () => json({}));
    const graph = createGraphClient({
      auth,
      baseUrl: BASE,
      fetchImpl: fetchImpl as never,
      now: () => clock,
    });
    await expect(graph.delete("/me/events/e", { deadline: 100 })).rejects.toBeInstanceOf(
      TimeoutError,
    );
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("scopeToOperation passes the signal and deadline to every request", async () => {
    const seen: unknown[] = [];
    const inner = {
      get: vi.fn(async (_p: string, o?: { signal?: AbortSignal; deadline?: number }) => {
        seen.push([o?.signal, o?.deadline]);
      }),
    } as unknown as GraphClient;
    const signal = new AbortController().signal;
    const scoped = scopeToOperation(inner, { signal, deadline: 42, signOutEpoch: undefined });
    await scoped.get("/a");
    await scoped.get("/b", { deadline: 7 }); // a request's own option still wins
    expect(seen).toEqual([
      [signal, 42],
      [signal, 7],
    ]);
  });
});
