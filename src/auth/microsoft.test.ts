import { createHash } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
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
import {
  createFileCoordinator,
  createMemoryCoordinator,
  type AuthCoordinator,
} from "./coordination.js";
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

  it("cachedOnly: the token last handed out, from memory: no store read, refresh or sign-in (Copilot review on #137)", async () => {
    let clock = 1_000;
    const base = memoryStore({ ...validCached, expiresAt: 100_000 });
    let loads = 0;
    let slowLoad = false;
    const store: TokenStore = {
      async load() {
        loads += 1;
        if (slowLoad) await new Promise(() => undefined); // a key unwrap that hangs
        return base.load();
      },
      save: (t) => base.save(t),
      clear: (o) => base.clear(o),
    };
    const fetchImpl = vi.fn();
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      openBrowser,
      now: () => clock,
    });
    // Nothing handed out yet: no token, and no store read to find one.
    await expect(auth.getAccessToken({ cachedOnly: true })).rejects.toBeInstanceOf(
      NoCurrentTokenError,
    );
    expect(loads).toBe(0);
    // A normal lookup hands one out; the store then stalls, and the cachedOnly
    // lookup doesn't touch it.
    expect(await auth.getAccessToken()).toBe("CACHED-AT");
    slowLoad = true;
    const before = loads;
    expect(await auth.getAccessToken({ cachedOnly: true })).toBe("CACHED-AT");
    expect(loads).toBe(before);
    // Expired: an error at once, with no refresh or browser.
    clock = 200_000;
    await expect(auth.getAccessToken({ cachedOnly: true })).rejects.toBeInstanceOf(
      NoCurrentTokenError,
    );
    expect(loads).toBe(before);
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("cachedOnly: a sign-out drops the held token, and a lookup from before it stops", async () => {
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(validCached),
      fetchImpl: vi.fn() as unknown as typeof fetch,
      now: () => 1_000,
    });
    await auth.getAccessToken();
    const epoch = auth.signOutEpoch();
    await auth.signOut();
    await expect(
      auth.getAccessToken({ cachedOnly: true, signOutEpoch: epoch }),
    ).rejects.toBeInstanceOf(SignedOutError);
    await expect(auth.getAccessToken({ cachedOnly: true })).rejects.toBeInstanceOf(
      NoCurrentTokenError,
    );
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
    // The sign-out's cleanup waits for the read, which holds the lock.
    const signedOut = auth.signOut();
    releaseLoad();
    await signedOut;
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

// Two auth instances sharing one token cache and one file coordinator stand in
// for two processes, e.g. Claude Desktop's server and the `logout` command (#95).
describe("sign-in and sign-out across processes (#95)", () => {
  const dirs: string[] = [];
  afterEach(() => {
    for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
  });
  function sharedDir(): string {
    const d = mkdtempSync(join(tmpdir(), "hotmail-95-"));
    dirs.push(d);
    return d;
  }
  const reply = (at: string) => () =>
    jsonResponse({ access_token: at, refresh_token: `RT-${at}`, expires_in: 3600 });

  /** One "process": its own auth, sharing the cache and the coordinator directory. */
  function processWith(
    store: TokenStore,
    dir: string,
    fetchImpl: () => Response | Promise<Response>,
  ): { auth: ReturnType<typeof createMicrosoftAuth>; authorizeUrl: () => string } {
    let url = "";
    const auth = createMicrosoftAuth({
      config,
      store,
      coordinator: createFileCoordinator(dir),
      fetchImpl: vi.fn(async () => fetchImpl()) as unknown as typeof fetch,
      openBrowser: async (u) => {
        url = u;
      },
      now: () => 1_000,
    });
    return { auth, authorizeUrl: () => url };
  }

  function redirectBack(authorizeUrl: string, params: Record<string, string>): Promise<Response> {
    const u = new URL(authorizeUrl);
    const redirect = new URL(u.searchParams.get("redirect_uri")!);
    redirect.hostname = "127.0.0.1";
    const q = new URLSearchParams({ state: u.searchParams.get("state")!, ...params });
    return fetch(`${redirect.origin}/?${q}`);
  }
  async function until(cond: () => boolean): Promise<void> {
    for (let i = 0; i < 200 && !cond(); i++) await new Promise((r) => setTimeout(r, 10));
    expect(cond()).toBe(true);
  }
  /** Settle a promise into an outcome now, so a late rejection is never unhandled. */
  const outcome = <T>(p: Promise<T>) =>
    p.then(
      (value) => ({ ok: true as const, value }),
      (error: unknown) => ({ ok: false as const, error }),
    );

  it("a sign-out in another process wins over a pending browser sign-in", async () => {
    const dir = sharedDir();
    const store = memoryStore(null);
    const server = processWith(store, dir, reply("SERVER-AT"));
    const cli = processWith(store, dir, reply("UNUSED"));

    const pending = outcome(server.auth.signIn());
    await until(() => server.authorizeUrl() !== "");
    await cli.auth.signOut(); // `logout` while the browser flow is still open

    await redirectBack(server.authorizeUrl(), { code: "CODE" }); // the user finishes it anyway
    const result = await pending;
    expect(result.ok).toBe(false);
    expect(!result.ok && result.error).toBeInstanceOf(SignedOutError);
    expect(store.current).toBeNull(); // the logout wasn't undone

    // A new sign-in after the logout saves normally.
    const firstUrl = server.authorizeUrl();
    const again = outcome(server.auth.signIn());
    await until(() => server.authorizeUrl() !== firstUrl);
    await redirectBack(server.authorizeUrl(), { code: "CODE2" });
    expect((await again).ok).toBe(true);
    expect(store.current?.accessToken).toBe("SERVER-AT");
  });

  it("a sign-out in another process wins over a refresh already in flight", async () => {
    const dir = sharedDir();
    const store = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 500, // expired at now = 1000
      scopes: config.scopes,
    });
    let sent = false;
    let answer = (_r: Response): void => undefined;
    const slowTokenEndpoint = () =>
      new Promise<Response>((r) => {
        sent = true;
        answer = r;
      });
    const server = processWith(store, dir, slowTokenEndpoint);
    const cli = processWith(store, dir, reply("UNUSED"));

    const lookup = outcome(server.auth.getAccessToken({ interactive: false }));
    await until(() => sent); // the refresh request is out
    await cli.auth.signOut();
    answer(reply("REFRESHED")());

    const result = await lookup;
    expect(result.ok).toBe(false);
    expect(!result.ok && result.error).toBeInstanceOf(SignedOutError);
    expect(store.current).toBeNull(); // the refresh didn't put the sign-in back
  });

  it("of two pending browser sign-ins, the first to finish wins", async () => {
    const dir = sharedDir();
    const store = memoryStore(null);
    const desktop = processWith(store, dir, reply("DESKTOP-AT"));
    const code = processWith(store, dir, reply("CODE-AT"));

    const desktopSignIn = outcome(desktop.auth.signIn());
    const codeSignIn = outcome(code.auth.signIn());
    await until(() => desktop.authorizeUrl() !== "" && code.authorizeUrl() !== "");

    await redirectBack(code.authorizeUrl(), { code: "C1" });
    const first = await codeSignIn;
    expect(first.ok && first.value.accessToken).toBe("CODE-AT");

    await redirectBack(desktop.authorizeUrl(), { code: "D1" }); // finished later
    const second = await desktopSignIn;
    // Its own tokens were dropped; it uses the sign-in that is saved.
    expect(second.ok && second.value.accessToken).toBe("CODE-AT");
    expect(store.current?.accessToken).toBe("CODE-AT");
  });

  it("a later sign-in doesn't adopt an older account when the winner's save failed (Copilot review on #126)", async () => {
    const dir = sharedDir();
    // Account X is already signed in.
    const base = memoryStore({
      accessToken: "ACCOUNT-X",
      refreshToken: "RT-X",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    });
    let saves = 0;
    const store: TokenStore = {
      load: () => base.load(),
      async save(t) {
        if (++saves === 1) throw new Error("disk full"); // the winner's save fails
        await base.save(t);
      },
      clear: (o) => base.clear(o),
    };
    const first = processWith(store, dir, reply("ACCOUNT-A"));
    const second = processWith(store, dir, reply("ACCOUNT-B"));

    const firstSignIn = outcome(first.auth.signIn());
    const secondSignIn = outcome(second.auth.signIn());
    await until(() => first.authorizeUrl() !== "" && second.authorizeUrl() !== "");

    await redirectBack(first.authorizeUrl(), { code: "A" }); // raises the count, then fails to save
    const a = await firstSignIn;
    expect(!a.ok && (a.error as Error).message).toMatch(/disk full/);

    await redirectBack(second.authorizeUrl(), { code: "B" });
    const b = await secondSignIn;
    // Not account X (an older sign-in no one chose now): told to try again.
    expect(b.ok).toBe(false);
    expect(!b.ok && b.error).toBeInstanceOf(SignInRequiredError);
    expect(!b.ok && (b.error as Error).message).toMatch(/try again/);
    expect(base.current?.accessToken).toBe("ACCOUNT-X"); // untouched
  });

  it("a refreshed sign-in keeps its generation, so it still counts as the winner", async () => {
    const dir = sharedDir();
    const store = memoryStore(null);
    const desktop = processWith(store, dir, reply("DESKTOP-AT"));
    const signIn = outcome(desktop.auth.signIn());
    await until(() => desktop.authorizeUrl() !== "");
    await redirectBack(desktop.authorizeUrl(), { code: "D" });
    expect((await signIn).ok).toBe(true);
    const generation = store.current?.signInGeneration;
    expect(generation).toBe(1);
    // Expire it, refresh it: the generation survives.
    store.current = { ...store.current!, expiresAt: 500 };
    await desktop.auth.getAccessToken({ interactive: false });
    expect(store.current?.signInGeneration).toBe(generation);
  });

  /** A token endpoint that answers only when told to. */
  function slowEndpoint() {
    const s = { sent: false, answer: (_r: Response): void => undefined };
    const respond = () =>
      new Promise<Response>((r) => {
        s.sent = true;
        s.answer = r;
      });
    return { s, respond };
  }

  it("a slow refresh of the old account can't replace a sign-in that finished elsewhere (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const store = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 500,
      scopes: config.scopes,
    });
    const { s, respond } = slowEndpoint();
    const server = processWith(store, dir, respond);
    const cli = processWith(store, dir, reply("NEW-ACCOUNT-AT"));

    const lookup = outcome(server.auth.getAccessToken({ interactive: false }));
    await until(() => s.sent); // refreshing the old account
    const login = outcome(cli.auth.signIn()); // meanwhile, `login` with another account
    await until(() => cli.authorizeUrl() !== "");
    await redirectBack(cli.authorizeUrl(), { code: "C" });
    expect((await login).ok).toBe(true);

    s.answer(reply("REFRESHED-OLD")());
    const result = await lookup;
    // The refresh's tokens were dropped, and the request uses the new sign-in.
    expect(result.ok && result.value).toBe("NEW-ACCOUNT-AT");
    expect(store.current?.accessToken).toBe("NEW-ACCOUNT-AT");
  });

  it("a request from before a sign-out elsewhere doesn't fall back to a browser sign-in (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const store = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 500,
      scopes: config.scopes,
    });
    const { s, respond } = slowEndpoint();
    const server = processWith(store, dir, respond);
    const cli = processWith(store, dir, reply("UNUSED"));

    const lookup = outcome(server.auth.getAccessToken({ interactiveWaitMs: 1_000 }));
    await until(() => s.sent);
    await cli.auth.signOut();
    // The refresh then fails, which would normally lead to a browser sign-in.
    s.answer(jsonResponse({ error: "invalid_grant" }, 400));

    const result = await lookup;
    expect(result.ok).toBe(false);
    expect(!result.ok && result.error).toBeInstanceOf(SignedOutError);
    expect(server.authorizeUrl()).toBe(""); // no browser was opened
    expect(store.current).toBeNull();
  });

  it("a sign-in started while a local sign-out is still cleaning up saves normally (Copilot review on #126)", async () => {
    let finishClear = (): void => undefined;
    let clearing = false;
    const base = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    });
    const store: TokenStore & { current: TokenSet | null } = {
      get current() {
        return base.current;
      },
      set current(v) {
        base.current = v;
      },
      load: () => base.load(),
      save: (t) => base.save(t),
      async clear() {
        clearing = true;
        await new Promise<void>((r) => (finishClear = r)); // a slow cleanup
        base.current = null;
      },
    };
    let url = "";
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () => reply("FRESH-AT")()) as unknown as typeof fetch,
      openBrowser: async (u) => {
        url = u;
      },
      now: () => 1_000,
    });

    const signOut = auth.signOut(); // still cleaning up...
    const signIn = outcome(auth.signIn()); // ...when `login` runs again
    await until(() => clearing);
    finishClear();
    await signOut;
    await until(() => url !== "");
    await redirectBack(url, { code: "C" });
    const result = await signIn;
    expect(result.ok && result.value.accessToken).toBe("FRESH-AT");
    expect(store.current?.accessToken).toBe("FRESH-AT");
  });

  it("a Graph request whose 401 retry spans a sign-out elsewhere stops instead of signing in (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const store = memoryStore({
      accessToken: "AT",
      refreshToken: "RT",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    });
    const server = processWith(store, dir, reply("UNUSED"));
    const cli = processWith(store, dir, reply("UNUSED"));
    let releaseResponse!: () => void;
    const responseGate = new Promise<void>((r) => (releaseResponse = r));
    let graphCalls = 0;
    const graph = createGraphClient({
      auth: server.auth,
      baseUrl: "https://graph.microsoft.com/v1.0",
      fetchImpl: (async () => {
        graphCalls += 1;
        await responseGate; // the response is slow…
        return jsonResponse({ error: { code: "InvalidAuthenticationToken" } }, 401);
      }) as unknown as typeof fetch,
      interactiveSignIn: { waitMs: 1_000 },
    });
    const call = outcome(graph.get("/me"));
    await until(() => graphCalls === 1);
    await cli.auth.signOut(); // …`logout` runs meanwhile…
    releaseResponse(); // …then the 401 triggers the retry.

    const result = await call;
    expect(!result.ok && result.error).toBeInstanceOf(SignedOutError);
    expect(server.authorizeUrl()).toBe(""); // no browser was opened
    expect(store.current).toBeNull();
  });

  it("a sign-out elsewhere waits for a cache read in progress, then removes what it read (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const base = memoryStore({
      accessToken: "PRE-LOGOUT",
      refreshToken: "RT",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    });
    let reading = false;
    let finishRead = (): void => undefined;
    const slowRead: TokenStore = {
      async load() {
        const read = base.current;
        if (!reading) {
          reading = true;
          await new Promise<void>((r) => (finishRead = r)); // the first read is slow
        }
        return read;
      },
      save: (t) => base.save(t),
      clear: (o) => base.clear(o),
    };
    const server = processWith(slowRead, dir, reply("UNUSED"));
    const cli = processWith(base, dir, reply("UNUSED"));

    const lookup = outcome(server.auth.getAccessToken({ interactive: false }));
    await until(() => reading);
    // The read holds the lock, so the logout can't finish (or even raise the
    // count) in the middle of it: the read is wholly before the logout.
    let loggedOut = false;
    const logout = cli.auth.signOut().then(() => (loggedOut = true));
    await new Promise((r) => setTimeout(r, 50));
    expect(loggedOut).toBe(false);
    finishRead();

    const result = await lookup;
    expect(result.ok && result.value).toBe("PRE-LOGOUT");
    await logout;
    expect(base.current).toBeNull(); // then the logout removed it
    const after = await outcome(server.auth.getAccessToken({ interactive: false }));
    expect(!after.ok && after.error).toBeInstanceOf(AuthError);
  });

  it("a lookup elsewhere during a sign-out's cleanup waits for it, and doesn't save the sign-in back (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const base = memoryStore({
      accessToken: "OLD",
      refreshToken: "RT",
      expiresAt: 500, // expired: a lookup would refresh it
      scopes: config.scopes,
    });
    let clearing = false;
    let finishClear = (): void => undefined;
    const slowClear: TokenStore = {
      load: () => base.load(),
      save: (t) => base.save(t),
      async clear(o) {
        clearing = true;
        await new Promise<void>((r) => (finishClear = r)); // a slow cleanup
        await base.clear(o);
      },
    };
    const refresh = vi.fn(() => reply("REFRESHED")());
    const server = processWith(base, dir, refresh);
    const cli = processWith(slowClear, dir, reply("UNUSED"));

    const logout = cli.auth.signOut();
    await until(() => clearing); // the count is raised; the cache is not yet deleted
    const lookup = outcome(server.auth.getAccessToken({ interactive: false }));
    await new Promise((r) => setTimeout(r, 50));
    finishClear();
    await logout;

    const result = await lookup;
    expect(!result.ok && result.error).toBeInstanceOf(AuthError); // not signed in
    expect(refresh).not.toHaveBeenCalled(); // it never read the old refresh token
    expect(base.current).toBeNull();
  });

  it("a sign-in whose count can't be stored saves nothing (Copilot review on #126)", async () => {
    const inner = createMemoryCoordinator();
    const diskFull = new Error("disk full");
    const coordinator: AuthCoordinator = {
      snapshot: () => inner.snapshot(),
      withLock: (fn) =>
        inner.withLock(async (state) => {
          const result = await fn(state, async () => {
            throw diskFull;
          });
          if (result.next) throw diskFull;
          return result;
        }),
    };
    const store = memoryStore(null);
    let url = "";
    const auth = createMicrosoftAuth({
      config,
      store,
      coordinator,
      fetchImpl: vi.fn(async () => reply("NEW-ACCOUNT")()) as unknown as typeof fetch,
      openBrowser: async (u) => {
        url = u;
      },
      now: () => 1_000,
    });
    const signIn = outcome(auth.signIn());
    await until(() => url !== "");
    await redirectBack(url, { code: "C" });
    const result = await signIn;
    expect(!result.ok && result.error).toBe(diskFull);
    // Not cached under the old count, where a pending refresh could overwrite it.
    expect(store.current).toBeNull();
  });

  it("signing in again after a sign-out elsewhere doesn't join the stale browser sign-in (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const store = memoryStore(null);
    const server = processWith(store, dir, reply("SERVER-AT"));
    const cli = processWith(store, dir, reply("UNUSED"));

    const stale = outcome(server.auth.signIn());
    await until(() => server.authorizeUrl() !== "");
    const staleUrl = server.authorizeUrl();
    await cli.auth.signOut(); // while that browser flow is still open

    const again = outcome(server.auth.signIn());
    await until(() => server.authorizeUrl() !== staleUrl); // a new flow, not the old one
    const staleResult = await stale;
    expect(!staleResult.ok && staleResult.error).toBeInstanceOf(SignedOutError);
    await redirectBack(server.authorizeUrl(), { code: "C" });
    const result = await again;
    expect(result.ok && result.value.accessToken).toBe("SERVER-AT");
    expect(store.current?.accessToken).toBe("SERVER-AT");
  });

  it("an on-demand sign-in after a sign-out elsewhere doesn't join the stale one either (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const store = memoryStore(null);
    const server = processWith(store, dir, reply("SERVER-AT"));
    const cli = processWith(store, dir, reply("UNUSED"));

    // A tool call opens the browser, then stops waiting for it.
    const first = await outcome(server.auth.getAccessToken({ interactiveWaitMs: 20 }));
    expect(!first.ok && first.error).toBeInstanceOf(SignInRequiredError);
    const staleUrl = server.authorizeUrl();
    expect(staleUrl).not.toBe("");
    await cli.auth.signOut();

    // The next call, after the logout, gets a sign-in that can still save.
    const next = outcome(server.auth.getAccessToken({ interactiveWaitMs: 2_000 }));
    await until(() => server.authorizeUrl() !== staleUrl);
    await redirectBack(server.authorizeUrl(), { code: "C" });
    const result = await next;
    expect(result.ok && result.value).toBe("SERVER-AT");
  });

  it("the first request to notice a sign-out elsewhere isn't itself stopped by it", async () => {
    const dir = sharedDir();
    const store = memoryStore({
      accessToken: "AT",
      refreshToken: "RT",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    });
    const server = processWith(store, dir, reply("UNUSED"));
    const cli = processWith(store, dir, reply("UNUSED"));
    await server.auth.getAccessToken({ interactive: false }); // server has seen the state
    await cli.auth.signOut();
    store.current = {
      accessToken: "NEW",
      refreshToken: "RT2",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    }; // and `login` again, say

    // A standalone lookup, and an operation whose epoch is noted after syncing.
    expect(await server.auth.getAccessToken({ interactive: false })).toBe("NEW");
    await cli.auth.signOut();
    const epoch = await server.auth.syncSignOuts();
    const result = await outcome(
      server.auth.getAccessToken({ interactive: false, signOutEpoch: epoch }),
    );
    // Just "not signed in": the operation began after the sign-out it saw.
    expect(!result.ok && result.error).toBeInstanceOf(AuthError);
    expect(!result.ok && result.error).not.toBeInstanceOf(SignedOutError);
  });

  it("credentials found by the fallback's second read, across a sign-out elsewhere, aren't used (Copilot review on #126)", async () => {
    const dir = sharedDir();
    const valid = {
      accessToken: "PRE-LOGOUT",
      refreshToken: "RT2",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    };
    let loads = 0;
    let reading = false;
    let finishRead = (): void => undefined;
    const store: TokenStore = {
      async load() {
        loads += 1;
        if (loads === 1) return { ...valid, accessToken: "OLD", expiresAt: 500 }; // expired
        const read = valid; // read before the logout…
        reading = true;
        await new Promise<void>((r) => (finishRead = r)); // …returned after it
        return read;
      },
      save: async () => undefined,
      clear: async () => undefined,
    };
    const server = processWith(store, dir, () => jsonResponse({ error: "invalid_grant" }, 400));
    const cli = processWith(memoryStore(null), dir, reply("UNUSED"));

    const lookup = outcome(server.auth.getAccessToken({ interactiveWaitMs: 1_000 }));
    await until(() => reading); // the refresh failed; the fallback reads the cache again
    await cli.auth.signOut();
    finishRead();

    const result = await lookup;
    expect(!result.ok && result.error).toBeInstanceOf(SignedOutError);
    expect(server.authorizeUrl()).toBe("");
  });
});

describe("sign-out during the start of an operation (Copilot review on #126)", () => {
  it("a sign-in immediately followed by a sign-out doesn't go on to sign in", async () => {
    const openBrowser = vi.fn(async () => undefined);
    const store = memoryStore(null);
    const auth = createMicrosoftAuth({
      config,
      store,
      fetchImpl: vi.fn(async () =>
        jsonResponse({ access_token: "AT", refresh_token: "RT", expires_in: 3600 }),
      ) as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });
    const signIn = auth.signIn().then(
      () => new Error("expected the sign-in to stop"),
      (e: unknown) => e as Error,
    );
    await auth.signOut();
    expect(await signIn).toBeInstanceOf(SignedOutError);
    expect(openBrowser).not.toHaveBeenCalled();
    expect(store.current).toBeNull();
  });

  it("an operation whose catch-up spans a local sign-out keeps the earlier epoch, so it stops", async () => {
    const inner = createMemoryCoordinator();
    let releaseSnapshot = (): void => undefined;
    let gated = true;
    const coordinator: AuthCoordinator = {
      async snapshot() {
        const state = await inner.snapshot();
        if (gated) {
          gated = false;
          await new Promise<void>((r) => (releaseSnapshot = r)); // a slow catch-up
        }
        return state;
      },
      withLock: (fn) => inner.withLock(fn),
    };
    const openBrowser = vi.fn(async () => undefined);
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore({
        accessToken: "AT",
        refreshToken: "RT",
        expiresAt: 10_000_000,
        scopes: config.scopes,
      }),
      coordinator,
      fetchImpl: vi.fn(async () =>
        jsonResponse({ access_token: "AT", refresh_token: "RT", expires_in: 3600 }),
      ) as unknown as typeof fetch,
      openBrowser,
      now: () => 1_000,
    });

    const starting = auth.syncSignOuts(); // a tool call starts…
    await auth.signOut(); // …the user signs out meanwhile…
    releaseSnapshot();
    const epoch = await starting;
    expect(epoch).not.toBe(auth.signOutEpoch()); // …so the call belongs to before it
    await expect(auth.getAccessToken({ signOutEpoch: epoch })).rejects.toBeInstanceOf(
      SignedOutError,
    );
    expect(openBrowser).not.toHaveBeenCalled();
  });

  it("deletes nothing if the shared sign-out count can't be stored (Copilot review on #126)", async () => {
    const inner = createMemoryCoordinator();
    const diskFull = new Error("disk full");
    // Storing counters fails, whether at once or when the locked section ends.
    const coordinator: AuthCoordinator = {
      snapshot: () => inner.snapshot(),
      withLock: (fn) =>
        inner.withLock(async (state) => {
          const result = await fn(state, async () => {
            throw diskFull;
          });
          if (result.next) throw diskFull;
          return result;
        }),
    };
    const saved = {
      accessToken: "AT",
      refreshToken: "RT",
      expiresAt: 10_000_000,
      scopes: config.scopes,
    };
    const store = memoryStore(saved);
    const auth = createMicrosoftAuth({
      config,
      store,
      coordinator,
      fetchImpl: vi.fn(async () => jsonResponse({})) as unknown as typeof fetch,
      openBrowser: async () => undefined,
      now: () => 1_000,
    });

    await expect(auth.signOut()).rejects.toBe(diskFull);
    // The sign-in is still there, so no other process can resurrect a deleted
    // one past an unchanged count; the user is told the sign-out failed.
    expect(store.current).toEqual(saved);
    expect(await inner.snapshot()).toEqual({ signOuts: 0, signIns: 0 });
  });

  it("a catch-up during this process's own sign-out doesn't take it for a remote one (Copilot review on #126)", async () => {
    const inner = createMemoryCoordinator();
    let paused = false;
    let published = false;
    let releasePublish = (): void => undefined;
    let releaseCleanup = (): void => undefined;
    // The sign-out pauses before storing its count, then again after storing
    // it (visible to snapshots) but before noting it as its own.
    const coordinator: AuthCoordinator = {
      snapshot: () => inner.snapshot(),
      withLock: (fn) =>
        inner.withLock((state, persist) =>
          fn(state, async (next) => {
            if (paused) return persist(next); // only the sign-out's write pauses
            paused = true;
            await new Promise<void>((r) => (releasePublish = r));
            await persist(next);
            published = true;
            await new Promise<void>((r) => (releaseCleanup = r));
          }),
        ),
    };
    let url = "";
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      coordinator,
      fetchImpl: vi.fn(async () =>
        jsonResponse({ access_token: "AT", refresh_token: "RT", expires_in: 3600 }),
      ) as unknown as typeof fetch,
      openBrowser: async (u) => {
        url = u;
      },
      now: () => 1_000,
    });
    await auth.syncSignOuts(); // the shared counters have been seen once

    const signOut = auth.signOut();
    await vi.waitFor(() => expect(paused).toBe(true));
    // A sign-in started after the sign-out, while the count isn't stored yet…
    const signIn = auth.signIn().then(
      (t) => t,
      (e: unknown) => e as Error,
    );
    await new Promise((r) => setTimeout(r, 20));
    // …then the count is stored, and a catch-up reads it before the sign-out
    // has noted it as its own.
    releasePublish();
    await vi.waitFor(() => expect(published).toBe(true));
    const sync = auth.syncSignOuts();
    await new Promise((r) => setTimeout(r, 20));
    releaseCleanup();
    await signOut;
    await sync;
    await vi.waitFor(() => expect(url).not.toBe(""));
    const u = new URL(url);
    const redirect = new URL(u.searchParams.get("redirect_uri")!);
    redirect.hostname = "127.0.0.1";
    await fetch(
      `${redirect.origin}/?${new URLSearchParams({ state: u.searchParams.get("state")!, code: "C" })}`,
    );
    // The sign-in completes: no second sign-out was imagined.
    const result = await signIn;
    expect(result).not.toBeInstanceOf(Error);
  });
});

describe("re-opening a pending sign-in page (#139)", () => {
  it("opens the same page again when a later call joins it, at most every 15 s", async () => {
    let clock = 1_000;
    const opened: string[] = [];
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: vi.fn(async () =>
        jsonResponse({ access_token: "AT", refresh_token: "RT", expires_in: 3600 }),
      ) as unknown as typeof fetch,
      openBrowser: async (u) => {
        opened.push(u);
      },
      now: () => clock,
    });
    const ask = () =>
      auth.getAccessToken({ interactiveWaitMs: 10 }).then(
        (t) => t,
        (e: unknown) => e as Error,
      );

    // The first call opens the page and stops waiting.
    expect(String(await ask())).toMatch(/has opened in your browser/);
    expect(opened).toHaveLength(1);
    // A call soon after joins without opening another tab.
    clock += 5_000;
    expect(String(await ask())).toMatch(/Still waiting/);
    expect(opened).toHaveLength(1);
    // Later, the user may have closed the tab: the same page opens again.
    clock += 20_000;
    expect(String(await ask())).toMatch(/opened again in your browser/);
    expect(opened).toEqual([opened[0], opened[0]]);

    // Finishing that page completes the original sign-in.
    const u = new URL(opened[0]!);
    const redirect = new URL(u.searchParams.get("redirect_uri")!);
    redirect.hostname = "127.0.0.1";
    await fetch(
      `${redirect.origin}/?${new URLSearchParams({ state: u.searchParams.get("state")!, code: "C" })}`,
    );
    await vi.waitFor(async () =>
      expect(await auth.getAccessToken({ interactive: false })).toBe("AT"),
    );
  });

  it("forgets the page once a sign-out cancels the sign-in", async () => {
    let clock = 1_000;
    const opened: string[] = [];
    const auth = createMicrosoftAuth({
      config,
      store: memoryStore(null),
      fetchImpl: vi.fn() as unknown as typeof fetch,
      openBrowser: async (u) => {
        opened.push(u);
      },
      now: () => clock,
    });
    await auth.getAccessToken({ interactiveWaitMs: 10 }).catch(() => undefined);
    await auth.signOut();
    clock += 60_000;
    // A new call starts a fresh sign-in (a new page), not the cancelled one.
    await auth.getAccessToken({ interactiveWaitMs: 10 }).catch(() => undefined);
    expect(opened).toHaveLength(2);
    expect(opened[1]).not.toBe(opened[0]);
  });
});
