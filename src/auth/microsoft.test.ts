import { createHash } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import { loadConfig } from "../config.js";
import { createGraphClient } from "../graph/client.js";
import {
  AuthError,
  NoCurrentTokenError,
  SignedOutError,
  SignInRequiredError,
} from "../util/errors.js";
import {
  buildAuthorizeUrl,
  buildCodeExchangeBody,
  buildRefreshBody,
  createMicrosoftAuth,
  decodeIdTokenUsername,
  generatePkce,
  launcherSucceeded,
  SIGN_IN_RETRY_COOLDOWN_MS,
  toTokenSet,
  type TokenResponse,
} from "./microsoft.js";
import type { TokenSet, TokenStore } from "./tokens.js";

describe("launcherSucceeded (browser launcher exit status)", () => {
  /** A fake child process whose exit we control. */
  function child(exitCode: number | null = null) {
    const handlers: Record<string, (arg: unknown) => void> = {};
    return {
      exitCode,
      once(event: string, fn: (arg: unknown) => void) {
        handlers[event] = fn;
        return this;
      },
      emit: (event: string, arg: unknown) => handlers[event]?.(arg),
    };
  }

  it("fails when the launcher exits non-zero (e.g. xdg-open with no browser)", async () => {
    const c = child();
    const p = launcherSucceeded(c as never, 1_000);
    c.emit("exit", 3);
    await expect(p).rejects.toThrow(/exited with code 3/);
  });

  it("fails when the launcher is killed by a signal (exit code null)", async () => {
    const c = child();
    const p = launcherSucceeded(c as never, 1_000);
    c.emit("exit", null);
    await expect(p).rejects.toThrow(/terminated/);
  });

  it("fails for a launcher that had already exited non-zero", async () => {
    await expect(launcherSucceeded(child(1) as never, 1_000)).rejects.toThrow(/code 1/);
  });

  it("succeeds on exit 0, or if the launcher is still running after the grace period", async () => {
    const ok = child();
    const p = launcherSucceeded(ok as never, 1_000);
    ok.emit("exit", 0);
    await expect(p).resolves.toBeUndefined();
    await expect(launcherSucceeded(child() as never, 10)).resolves.toBeUndefined();
  });
});

const config = loadConfig({ MICROSOFT_CLIENT_ID: "client-abc" });

function base64url(s: string): string {
  return Buffer.from(s)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

/** In-memory token store for testing. */
function memoryStore(initial: TokenSet | null = null): TokenStore & { current: TokenSet | null } {
  const state = { current: initial };
  return {
    current: state.current,
    async load() {
      return this.current;
    },
    async save(t) {
      this.current = t;
    },
    async clear() {
      this.current = null;
    },
  };
}

describe("generatePkce", () => {
  it("produces a valid verifier and matching S256 challenge", () => {
    const { verifier, challenge } = generatePkce();
    expect(verifier).toMatch(/^[A-Za-z0-9_-]{43,128}$/);
    const expected = createHash("sha256")
      .update(verifier)
      .digest("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");
    expect(challenge).toBe(expected);
  });

  it("is random each call", () => {
    expect(generatePkce().verifier).not.toBe(generatePkce().verifier);
  });
});

describe("buildAuthorizeUrl", () => {
  it("targets the default consumers authority with PKCE and required params", () => {
    const url = new URL(
      buildAuthorizeUrl({
        authority: config.authority,
        clientId: config.microsoftClientId,
        redirectUri: "http://localhost:12345",
        scopes: ["openid", "Mail.Read"],
        state: "st-1",
        codeChallenge: "chal-1",
      }),
    );
    expect(url.origin + url.pathname).toBe(
      "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize",
    );
    expect(url.searchParams.get("client_id")).toBe("client-abc");
    expect(url.searchParams.get("response_type")).toBe("code");
    expect(url.searchParams.get("redirect_uri")).toBe("http://localhost:12345");
    expect(url.searchParams.get("scope")).toBe("openid Mail.Read");
    expect(url.searchParams.get("state")).toBe("st-1");
    expect(url.searchParams.get("code_challenge")).toBe("chal-1");
    expect(url.searchParams.get("code_challenge_method")).toBe("S256");
  });
});

describe("token request bodies", () => {
  it("builds an authorization_code body with the verifier", () => {
    const body = buildCodeExchangeBody({
      clientId: "c",
      code: "the-code",
      redirectUri: "http://localhost:1",
      codeVerifier: "verif",
      scopes: ["a", "b"],
    });
    expect(body.get("grant_type")).toBe("authorization_code");
    expect(body.get("code")).toBe("the-code");
    expect(body.get("code_verifier")).toBe("verif");
    expect(body.get("scope")).toBe("a b");
    expect(body.has("client_secret")).toBe(false); // public client — no secret
  });

  it("builds a refresh_token body", () => {
    const body = buildRefreshBody({ clientId: "c", refreshToken: "r", scopes: ["a"] });
    expect(body.get("grant_type")).toBe("refresh_token");
    expect(body.get("refresh_token")).toBe("r");
    expect(body.has("client_secret")).toBe(false);
  });
});

describe("decodeIdTokenUsername", () => {
  it("extracts preferred_username from a JWT payload", () => {
    const jwt = `h.${base64url(JSON.stringify({ preferred_username: "tony@hotmail.com" }))}.sig`;
    expect(decodeIdTokenUsername(jwt)).toBe("tony@hotmail.com");
  });
  it("returns undefined for malformed input", () => {
    expect(decodeIdTokenUsername(undefined)).toBeUndefined();
    expect(decodeIdTokenUsername("garbage")).toBeUndefined();
  });
});

describe("toTokenSet", () => {
  const resp: TokenResponse = {
    access_token: "AT",
    refresh_token: "RT",
    expires_in: 3600,
    scope: "openid Mail.Read",
    token_type: "Bearer",
  };
  it("computes expiresAt from expires_in and now", () => {
    const set = toTokenSet(resp, 1_000_000);
    expect(set.accessToken).toBe("AT");
    expect(set.refreshToken).toBe("RT");
    expect(set.expiresAt).toBe(1_000_000 + 3600 * 1000);
    expect(set.scopes).toEqual(["openid", "Mail.Read"]);
  });
  it("keeps the previous refresh token when none is returned", () => {
    const { refresh_token: _omit, ...noRt } = resp;
    void _omit;
    const set = toTokenSet(noRt, 0, { refreshToken: "OLD-RT" });
    expect(set.refreshToken).toBe("OLD-RT");
  });
  it("throws when no refresh token is available at all", () => {
    const { refresh_token: _omit, ...noRt } = resp;
    void _omit;
    expect(() => toTokenSet(noRt, 0)).toThrow(AuthError);
  });
});

describe("createMicrosoftAuth.getAccessToken", () => {
  const future = 10_000_000_000_000;
  const validCached: TokenSet = {
    accessToken: "CACHED-AT",
    refreshToken: "RT",
    expiresAt: future,
    scopes: config.scopes,
  };

  it("cachedOnly: the saved token while it lasts, else an error, never a refresh or sign-in (#137)", async () => {
    const fetchImpl = vi.fn();
    const openBrowser = vi.fn(async () => undefined);
    const make = (expiresAt: number) =>
      createMicrosoftAuth({
        config,
        store: memoryStore({ ...validCached, expiresAt }),
        fetchImpl: fetchImpl as unknown as typeof fetch,
        openBrowser,
        now: () => 1_000,
      });
    // 30 s left: inside the usual 60 s refresh margin, but still current.
    expect(await make(31_000).getAccessToken({ cachedOnly: true })).toBe("CACHED-AT");
    // Expired: no refresh, no browser, and an error saying so.
    await expect(make(500).getAccessToken({ cachedOnly: true })).rejects.toBeInstanceOf(
      NoCurrentTokenError,
    );
    const empty = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    await expect(empty.getAccessToken({ cachedOnly: true })).rejects.toBeInstanceOf(
      NoCurrentTokenError,
    );
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("returns a cached access token that is still valid", async () => {
    const fetchImpl = vi.fn();
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(validCached),
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    expect(await auth.getAccessToken()).toBe("CACHED-AT");
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("refreshes silently when the cached token is expired", async () => {
    const store = memoryStore({ ...validCached, expiresAt: 1_500 });
    const fetchImpl = vi.fn(async () =>
      jsonResponse({ access_token: "NEW-AT", refresh_token: "NEW-RT", expires_in: 3600 }),
    );
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    expect(await auth.getAccessToken()).toBe("NEW-AT");
    expect(fetchImpl).toHaveBeenCalledOnce();
    expect(store.current?.accessToken).toBe("NEW-AT");
    expect(store.current?.refreshToken).toBe("NEW-RT");
  });

  it("gives up on a token endpoint that never answers, instead of hanging (#37)", async () => {
    const store = memoryStore({ ...validCached, expiresAt: 1_500 });
    const stalled = vi.fn(
      (_url: string, init?: RequestInit) =>
        new Promise<Response>((_, reject) =>
          init?.signal?.addEventListener("abort", () => reject(new Error("aborted"))),
        ),
    );
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: stalled as unknown as typeof fetch,
      now: () => 1_000,
      tokenTimeoutMs: 20,
      logger: { debug() {}, info() {}, warn() {}, error() {} } as never,
    });
    await expect(auth.getAccessToken({ interactive: false })).rejects.toBeInstanceOf(AuthError);
    expect(stalled).toHaveBeenCalledOnce();
    expect(stalled.mock.calls[0]![1]?.signal).toBeInstanceOf(AbortSignal);
  });

  it("throws (no browser) when refresh fails and interactive=false", async () => {
    const store = memoryStore({ ...validCached, expiresAt: 1_500 });
    const fetchImpl = vi.fn(async () => jsonResponse({ error: "invalid_grant" }, 400));
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    await expect(auth.getAccessToken({ interactive: false })).rejects.toBeInstanceOf(AuthError);
  });

  it("throws when not signed in and interactive=false", async () => {
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: vi.fn() as unknown as typeof fetch,
      now: () => 1_000,
    });
    await expect(auth.getAccessToken({ interactive: false })).rejects.toBeInstanceOf(AuthError);
  });

  it("forces a refresh of a still-valid token when minTtlMs is huge (the 401-recovery contract)", async () => {
    const store = memoryStore({ ...validCached, accessToken: "OLD-AT", expiresAt: future });
    const fetchImpl = vi.fn(async () =>
      jsonResponse({ access_token: "FORCED-AT", refresh_token: "RT2", expires_in: 3600 }),
    );
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    // Even though the cached token is valid, a huge minTtlMs must bypass it and refresh.
    const token = await auth.getAccessToken({ minTtlMs: Number.MAX_SAFE_INTEGER });
    expect(token).toBe("FORCED-AT");
    expect(fetchImpl).toHaveBeenCalledOnce();
    expect(store.current?.accessToken).toBe("FORCED-AT");
  });

  it("collapses concurrent refreshes into a single network call (single-flight)", async () => {
    const store = memoryStore({ ...validCached, expiresAt: 1_500 });
    let calls = 0;
    const fetchImpl = vi.fn(async () => {
      calls += 1;
      return jsonResponse({
        access_token: `AT-${calls}`,
        refresh_token: `RT-${calls}`,
        expires_in: 3600,
      });
    });
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    const [a, b] = await Promise.all([auth.getAccessToken(), auth.getAccessToken()]);
    // One refresh serves both callers — the rotating refresh token isn't spent twice.
    expect(fetchImpl).toHaveBeenCalledOnce();
    expect(a).toBe("AT-1");
    expect(b).toBe("AT-1");
  });

  it("recovers from a corrupt cache: clears it and reports not-signed-in", async () => {
    const clear = vi.fn(async () => {});
    const throwingStore: TokenStore = {
      load: async () => {
        throw new AuthError("Token cache is unreadable.");
      },
      save: async () => {},
      clear,
    };
    const auth = createMicrosoftAuth({
      config,
      store: throwingStore,
      fetchImpl: vi.fn() as unknown as typeof fetch,
      now: () => 1_000,
    });
    await expect(auth.getAccessToken({ interactive: false })).rejects.toBeInstanceOf(AuthError);
    expect(clear).toHaveBeenCalledOnce();
    // getAccount() swallows the same load error and reports "no account".
    expect(await auth.getAccount()).toBeUndefined();
  });
});

describe("on-demand interactive sign-in (real loopback listener)", () => {
  const tokenReply = () =>
    jsonResponse({ access_token: "INTERACTIVE-AT", refresh_token: "RT", expires_in: 3600 });

  /** Act as the browser: follow the authorize URL's redirect back to the loopback listener. */
  function redirectBack(authorizeUrl: string, params: Record<string, string>): Promise<Response> {
    const u = new URL(authorizeUrl);
    const redirect = new URL(u.searchParams.get("redirect_uri")!);
    redirect.hostname = "127.0.0.1"; // the listener binds IPv4 loopback
    const q = new URLSearchParams({ state: u.searchParams.get("state")!, ...params });
    return fetch(`${redirect.origin}/?${q}`);
  }

  async function until(cond: () => boolean): Promise<void> {
    for (let i = 0; i < 200 && !cond(); i++) await new Promise((r) => setTimeout(r, 10));
    expect(cond()).toBe(true);
  }

  it("signs in and returns the new token when the user finishes within the wait", async () => {
    const store = memoryStore(null);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser: async (url) => {
        void redirectBack(url, { code: "CODE" });
      },
      now: () => 1_000,
    });
    expect(await auth.getAccessToken({ interactiveWaitMs: 5_000 })).toBe("INTERACTIVE-AT");
    expect(store.current?.refreshToken).toBe("RT");
  });

  it("asks the user to finish in the browser, joins the same attempt, then succeeds", async () => {
    const store = memoryStore(null);
    let authorizeUrl = "";
    const openBrowser = vi.fn(async (url: string) => {
      authorizeUrl = url;
    });
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });

    const first = auth.getAccessToken({ interactiveWaitMs: 20 });
    await expect(first).rejects.toBeInstanceOf(SignInRequiredError);
    await expect(first).rejects.toThrow(/sign-in window has opened/);
    await expect(auth.getAccessToken({ interactiveWaitMs: 20 })).rejects.toThrow(/Still waiting/);
    expect(openBrowser).toHaveBeenCalledOnce();

    await redirectBack(authorizeUrl, { code: "CODE" });
    await until(() => store.current !== null);
    expect(await auth.getAccessToken({ interactiveWaitMs: 20 })).toBe("INTERACTIVE-AT");
    expect(openBrowser).toHaveBeenCalledOnce();
  });

  it("ignores stray loopback requests (wrong state) instead of aborting the sign-in (#41)", async () => {
    const store = memoryStore(null);
    let authorizeUrl = "";
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser: async (url) => {
        authorizeUrl = url;
      },
      now: () => 1_000,
    });
    const pending = auth.signIn();
    await until(() => authorizeUrl !== "");

    // A hostile/stray local request with a wrong state, and one trying to cancel.
    expect((await redirectBack(authorizeUrl, { code: "EVIL", state: "wrong" })).status).toBe(400);
    expect(
      (await redirectBack(authorizeUrl, { error: "access_denied", state: "wrong" })).status,
    ).toBe(400);
    expect(store.current).toBeNull(); // nothing completed or aborted

    // Microsoft's real redirect still completes the sign-in.
    await redirectBack(authorizeUrl, { code: "CODE" });
    await expect(pending).resolves.toMatchObject({ accessToken: "INTERACTIVE-AT" });
  });

  it("ignores a launcher failure that arrives after Microsoft's redirect was accepted", async () => {
    let releaseExchange!: () => void;
    const exchangeGate = new Promise<void>((r) => (releaseExchange = r));
    let failLauncher!: (e: Error) => void;
    let authorizeUrl = "";
    const store = memoryStore(null);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => {
        await exchangeGate; // token exchange in progress…
        return tokenReply();
      }) as unknown as typeof fetch,
      openBrowser: (url) => {
        authorizeUrl = url;
        return new Promise<void>((_, reject) => (failLauncher = reject));
      },
      now: () => 1_000,
    });
    const pending = auth.signIn();
    await until(() => authorizeUrl !== "");
    await redirectBack(authorizeUrl, { code: "CODE" }); // callback accepted
    failLauncher(new Error("launcher exited 1")); // …the launcher reports failure late…
    releaseExchange(); // …and the exchange then succeeds.
    await expect(pending).resolves.toMatchObject({ accessToken: "INTERACTIVE-AT" });
    expect(store.current?.accessToken).toBe("INTERACTIVE-AT");
  });

  it("ends the attempt at once if the browser can't be opened (server mode)", async () => {
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: vi.fn() as unknown as typeof fetch,
      openBrowser: async () => {
        throw new Error("no default browser");
      },
      now: () => 1_000,
    });
    const started = Date.now();
    await expect(auth.getAccessToken({ interactiveWaitMs: 5_000 })).rejects.toThrow(
      /Couldn't open a web browser/,
    );
    expect(Date.now() - started).toBeLessThan(2_000); // not stuck until the 5-minute timeout
  });

  it("with printUrlIfBrowserFails (login), keeps waiting so the printed URL still works", async () => {
    const store = memoryStore(null);
    let authorizeUrl = "";
    const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    try {
      const auth = createMicrosoftAuth({
        config,
        store,
        fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
        openBrowser: async (url) => {
          authorizeUrl = url;
          throw new Error("no default browser");
        },
        printUrlIfBrowserFails: true,
        now: () => 1_000,
      });
      const pending = auth.signIn();
      await until(() => stderr.mock.calls.some(([s]) => String(s).includes("Open this URL")));
      await redirectBack(authorizeUrl, { code: "CODE" }); // the user pastes the URL
      await expect(pending).resolves.toMatchObject({ accessToken: "INTERACTIVE-AT" });
    } finally {
      stderr.mockRestore();
    }
  });

  it("joins a pending sign-in instead of refreshing again (no second window)", async () => {
    const expired: TokenSet = {
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 0,
      scopes: config.scopes,
    };
    const store = memoryStore(expired);
    let authorizeUrl = "";
    const openBrowser = vi.fn(async (url: string) => {
      authorizeUrl = url;
    });
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) =>
      String(init?.body).includes("grant_type=refresh_token")
        ? jsonResponse({ error: "invalid_grant" }, 400)
        : tokenReply(),
    );
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    const refreshCalls = () =>
      fetchImpl.mock.calls.filter(([, i]) => String(i?.body).includes("refresh_token")).length;

    await expect(auth.getAccessToken({ interactiveWaitMs: 20 })).rejects.toThrow(/has opened/);
    await expect(auth.getAccessToken({ interactiveWaitMs: 20 })).rejects.toThrow(/Still waiting/);
    expect(refreshCalls()).toBe(1); // the retry joined the sign-in, no second refresh
    expect(openBrowser).toHaveBeenCalledOnce();

    await redirectBack(authorizeUrl, { code: "CODE" });
    await until(() => store.current?.accessToken === "INTERACTIVE-AT");
    expect(await auth.getAccessToken({ interactiveWaitMs: 20 })).toBe("INTERACTIVE-AT");
    expect(openBrowser).toHaveBeenCalledOnce();
  });

  it("uses tokens saved during a failed refresh instead of opening a browser", async () => {
    const store = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 0,
      scopes: config.scopes,
    });
    const openBrowser = vi.fn(async () => undefined);
    const fetchImpl = vi.fn(async () => {
      // Meanwhile another sign-in completes and saves fresh tokens…
      await store.save({
        accessToken: "FRESH",
        refreshToken: "RT2",
        expiresAt: 10_000_000,
        scopes: config.scopes,
      });
      return jsonResponse({ error: "invalid_grant" }, 400); // …and this refresh fails.
    });
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    expect(await auth.getAccessToken({ interactiveWaitMs: 20 })).toBe("FRESH");
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("a caller whose cache re-check was stale uses the sign-in that just completed", async () => {
    const expired: TokenSet = { accessToken: "OLD", refreshToken: "RT", expiresAt: 0, scopes: [] };
    let current: TokenSet | null = expired;
    let loads = 0;
    let releaseStale!: () => void;
    const staleGate = new Promise<void>((r) => (releaseStale = r));
    const store: TokenStore = {
      async load() {
        const snapshot = current;
        // The 4th read (caller B's re-check) is slow and returns what it saw at call time.
        if (++loads === 4) await staleGate;
        return snapshot;
      },
      async save(t) {
        current = t;
      },
      async clear() {
        current = null;
      },
    };
    let authorizeUrl = "";
    const openBrowser = vi.fn(async (url: string) => {
      authorizeUrl = url;
    });
    const fetchImpl = vi.fn(async (_u: string, init?: RequestInit) =>
      String(init?.body).includes("grant_type=refresh_token")
        ? jsonResponse({ error: "invalid_grant" }, 400)
        : tokenReply(),
    );
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });

    const a = auth.getAccessToken({ interactiveWaitMs: 5_000 });
    const b = auth.getAccessToken({ interactiveWaitMs: 5_000 });
    await until(() => authorizeUrl !== "");
    await redirectBack(authorizeUrl, { code: "CODE" }); // A's sign-in completes…
    await expect(a).resolves.toBe("INTERACTIVE-AT");
    releaseStale(); // …then B's stale re-check returns the old, expired snapshot.
    await expect(b).resolves.toBe("INTERACTIVE-AT");
    expect(openBrowser).toHaveBeenCalledOnce(); // no second window
  });

  it("also uses them on the 401-recovery path (forced refresh, huge minTtlMs)", async () => {
    const store = memoryStore({
      accessToken: "REVOKED",
      refreshToken: "RT",
      expiresAt: 10_000_000, // still "valid" by the clock, but Graph said 401
      scopes: config.scopes,
    });
    const openBrowser = vi.fn(async () => undefined);
    const fetchImpl = vi.fn(async () => {
      await store.save({
        accessToken: "FRESH",
        refreshToken: "RT2",
        expiresAt: 10_000_000,
        scopes: config.scopes,
      });
      return jsonResponse({ error: "invalid_grant" }, 400);
    });
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    expect(
      await auth.getAccessToken({ interactiveWaitMs: 20, minTtlMs: Number.MAX_SAFE_INTEGER }),
    ).toBe("FRESH");
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("a sign-in still pending when the user signs out is discarded, not saved (#43)", async () => {
    const store = memoryStore(null);
    let authorizeUrl = "";
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser: async (url) => {
        authorizeUrl = url;
      },
      now: () => 1_000,
    });
    // Capture the outcome now: the attempt rejects while we're still signing out.
    const outcome = auth.signIn().then(
      () => new Error("expected the sign-in to be discarded"),
      (e: unknown) => e as Error,
    );
    await until(() => authorizeUrl !== "");
    await auth.signOut(); // user signs out while the browser flow is open
    expect(await outcome).toBeInstanceOf(SignedOutError);
    // The pending listener was closed, so finishing the old tab can't complete anything.
    await expect(redirectBack(authorizeUrl, { code: "CODE" })).rejects.toThrow();
    expect(store.current).toBeNull(); // the sign-out stands
  });

  it("after sign-out, a new sign-in starts a fresh browser flow straight away", async () => {
    const store = memoryStore(null);
    const urls: string[] = [];
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser: async (url) => {
        urls.push(url);
      },
      now: () => 1_000,
    });
    const old = auth.signIn().catch((e: unknown) => e);
    await until(() => urls.length === 1);
    await auth.signOut();
    await old;

    const fresh = auth.signIn(); // must not join the old, cancelled attempt
    await until(() => urls.length === 2);
    await redirectBack(urls[1]!, { code: "CODE" });
    await expect(fresh).resolves.toMatchObject({ accessToken: "INTERACTIVE-AT" });
    // No cooldown was imposed by the cancelled attempt.
    expect(await auth.getAccessToken({ interactiveWaitMs: 20 })).toBe("INTERACTIVE-AT");
  });

  it("an interactive request whose refresh is cut off by sign-out doesn't open a browser", async () => {
    const store = memoryStore({ accessToken: "OLD", refreshToken: "RT", expiresAt: 0, scopes: [] });
    let releaseRefresh!: () => void;
    const gate = new Promise<void>((r) => (releaseRefresh = r));
    const fetchImpl = vi.fn(async () => {
      await gate;
      return tokenReply();
    });
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    const outcome = auth.getAccessToken({ interactiveWaitMs: 5_000 }).then(
      () => new Error("expected the request to stop"),
      (e: unknown) => e as Error,
    );
    await until(() => fetchImpl.mock.calls.length > 0);
    await auth.signOut();
    releaseRefresh();
    expect(await outcome).toBeInstanceOf(SignedOutError);
    expect(openBrowser).not.toHaveBeenCalled();
    expect(store.current).toBeNull();
  });

  it("a request whose cache read straddles a sign-out doesn't adopt the new epoch", async () => {
    let releaseLoad!: () => void;
    const loadGate = new Promise<void>((r) => (releaseLoad = r));
    let loads = 0;
    let current: TokenSet | null = {
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 0,
      scopes: [],
    };
    const store: TokenStore = {
      async load() {
        const snapshot = current;
        if (++loads === 1) await loadGate; // the request's first read is slow…
        return snapshot; // …and returns the pre-sign-out credentials
      },
      async save(t) {
        current = t;
      },
      async clear() {
        current = null;
      },
    };
    const fetchImpl = vi.fn(async () => tokenReply());
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    const outcome = auth.getAccessToken({ interactive: false }).then(
      () => new Error("expected the request to stop"),
      (e: unknown) => e as Error,
    );
    await auth.signOut();
    releaseLoad();
    expect(await outcome).toBeInstanceOf(SignedOutError);
    expect(fetchImpl).not.toHaveBeenCalled(); // no refresh with the stale credentials
    expect(current).toBeNull();
  });

  it("a request starting while sign-out is still cleaning up waits instead of reviving tokens", async () => {
    let releaseClear!: () => void;
    const clearGate = new Promise<void>((r) => (releaseClear = r));
    let current: TokenSet | null = {
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 0,
      scopes: [],
    };
    let clearStarted = false;
    const store: TokenStore = {
      async load() {
        return current;
      },
      async save(t) {
        current = t;
      },
      async clear() {
        clearStarted = true;
        await clearGate; // slow cleanup (e.g. file I/O)
        current = null;
      },
    };
    const fetchImpl = vi.fn(async () => tokenReply());
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    const signingOut = auth.signOut();
    await until(() => clearStarted);
    // A new request (already in the new epoch) arrives mid-cleanup…
    const request = auth.getAccessToken({ interactive: false }).catch((e: unknown) => e);
    releaseClear();
    await signingOut;
    // …and finds nothing to refresh, rather than refreshing the old token back.
    expect(await request).toBeInstanceOf(AuthError);
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(current).toBeNull();
  });

  it("a request waiting on one sign-out's cleanup is stopped by a second sign-out", async () => {
    let releaseFirstClear!: () => void;
    const firstClearGate = new Promise<void>((r) => (releaseFirstClear = r));
    let clears = 0;
    const store: TokenStore = {
      async load() {
        return null;
      },
      async save() {},
      async clear() {
        if (++clears === 1) await firstClearGate;
      },
    };
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn() as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    const first = auth.signOut();
    await until(() => clears === 1);
    // A request arrives mid-cleanup and waits…
    const request = auth.getAccessToken({ interactiveWaitMs: 20 }).then(
      () => new Error("expected the request to stop"),
      (e: unknown) => e as Error,
    );
    const second = auth.signOut(); // …then the user signs out again.
    releaseFirstClear();
    await Promise.all([first, second]);
    expect(await request).toBeInstanceOf(SignedOutError);
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("a Graph request whose 401 retry spans a sign-out stops instead of signing in", async () => {
    const store = memoryStore({
      accessToken: "AT",
      refreshToken: "RT",
      expiresAt: 10_000_000_000_000,
      scopes: config.scopes,
    });
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    let releaseResponse!: () => void;
    const responseGate = new Promise<void>((r) => (releaseResponse = r));
    let graphCalls = 0;
    const graph = createGraphClient({
      auth,
      baseUrl: "https://graph.microsoft.com/v1.0",
      fetchImpl: (async () => {
        graphCalls += 1;
        await responseGate; // the response is slow…
        return jsonResponse({ error: { code: "InvalidAuthenticationToken" } }, 401);
      }) as unknown as typeof fetch,
      interactiveSignIn: { waitMs: 5_000 },
    });
    const call = graph.get("/me").then(
      () => new Error("expected the request to stop"),
      (e: unknown) => e as Error,
    );
    await until(() => graphCalls === 1);
    await auth.signOut(); // …the user signs out meanwhile…
    releaseResponse(); // …then the 401 triggers the retry.
    expect(await call).toBeInstanceOf(SignedOutError);
    expect(openBrowser).not.toHaveBeenCalled();
    expect(store.current).toBeNull();
  });

  it("paging stops at the next page if the user signed out during the first", async () => {
    const store = memoryStore({
      accessToken: "AT",
      refreshToken: "RT",
      expiresAt: 10_000_000_000_000,
      scopes: config.scopes,
    });
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    let releaseFirstPage!: () => void;
    const firstPageGate = new Promise<void>((r) => (releaseFirstPage = r));
    let graphCalls = 0;
    const graph = createGraphClient({
      auth,
      baseUrl: "https://graph.microsoft.com/v1.0",
      fetchImpl: (async () => {
        graphCalls += 1;
        await firstPageGate; // the first page is slow, but succeeds…
        return jsonResponse({
          value: [1],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skip=1",
        });
      }) as unknown as typeof fetch,
      interactiveSignIn: { waitMs: 5_000 },
    });
    const listing = graph.getAllPages("/me/calendarView").then(
      () => new Error("expected paging to stop"),
      (e: unknown) => e as Error,
    );
    await until(() => graphCalls === 1);
    await auth.signOut(); // …while the user signs out.
    releaseFirstPage();
    expect(await listing).toBeInstanceOf(SignedOutError); // page 2 doesn't sign in again
    expect(graphCalls).toBe(1);
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("a save already writing when the user signs out is removed by the cleanup", async () => {
    let releaseSave!: () => void;
    const saveGate = new Promise<void>((r) => (releaseSave = r));
    const order: string[] = [];
    let current: TokenSet | null = {
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 0,
      scopes: [],
    };
    let saveStarted = false;
    const store: TokenStore = {
      async load() {
        return current;
      },
      async save(t) {
        saveStarted = true;
        order.push("save-start");
        await saveGate; // e.g. slow key creation / file I/O
        current = t;
        order.push("save-end");
      },
      async clear() {
        order.push("clear");
        current = null;
      },
    };
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => tokenReply()) as unknown as typeof fetch,
      now: () => 1_000,
    });
    const refreshing = auth.getAccessToken({ interactive: false }).catch((e: unknown) => e);
    await until(() => saveStarted);
    const signingOut = auth.signOut();
    releaseSave();
    await signingOut;
    await refreshing;
    expect(order).toEqual(["save-start", "save-end", "clear"]);
    expect(current).toBeNull(); // the sign-out stands
  });

  it("a refresh in flight when the user signs out doesn't re-save tokens", async () => {
    const store = memoryStore({ accessToken: "OLD", refreshToken: "RT", expiresAt: 0, scopes: [] });
    let releaseRefresh!: () => void;
    const gate = new Promise<void>((r) => (releaseRefresh = r));
    const fetchImpl = vi.fn(async () => {
      await gate;
      return tokenReply();
    });
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      now: () => 1_000,
    });
    const outcome = auth.getAccessToken({ interactive: false }).then(
      () => new Error("expected the refresh to be discarded"),
      (e: unknown) => e as Error,
    );
    await until(() => fetchImpl.mock.calls.length > 0); // the refresh request is in flight
    await auth.signOut();
    releaseRefresh();
    expect(await outcome).toBeInstanceOf(AuthError);
    expect(store.current).toBeNull();
  });

  it("signOut clears the cache including the key", async () => {
    const clear = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store: { load: async () => null, save: async () => undefined, clear },
      fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    await auth.signOut();
    expect(clear).toHaveBeenCalledWith({ includeKey: true });
  });

  it("cools down after a cancelled sign-in instead of reopening the browser", async () => {
    let clock = 1_000;
    const openBrowser = vi.fn(async (url: string) => {
      void redirectBack(url, { error: "access_denied" });
    });
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: vi.fn() as unknown as typeof fetch,
      openBrowser,
      now: () => clock,
    });

    await expect(auth.getAccessToken({ interactiveWaitMs: 5_000 })).rejects.toThrow(
      /Authorization failed: access_denied/,
    );
    await expect(auth.getAccessToken({ interactiveWaitMs: 5_000 })).rejects.toThrow(
      /didn't complete/,
    );
    expect(openBrowser).toHaveBeenCalledOnce();

    clock += SIGN_IN_RETRY_COOLDOWN_MS + 1;
    await expect(auth.getAccessToken({ interactiveWaitMs: 5_000 })).rejects.toThrow(
      /access_denied/,
    );
    expect(openBrowser).toHaveBeenCalledTimes(2);
  });
});

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
}
