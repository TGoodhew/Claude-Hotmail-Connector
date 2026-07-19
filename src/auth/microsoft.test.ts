import { createHash } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import { loadConfig } from "../config.js";
import { AuthError } from "../util/errors.js";
import {
  buildAuthorizeUrl,
  buildCodeExchangeBody,
  buildRefreshBody,
  createMicrosoftAuth,
  decodeIdTokenUsername,
  generatePkce,
  toTokenSet,
  type TokenResponse,
} from "./microsoft.js";
import type { TokenSet, TokenStore } from "./tokens.js";

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
  it("targets the common authority with PKCE and required params", () => {
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
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
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
});

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
}
