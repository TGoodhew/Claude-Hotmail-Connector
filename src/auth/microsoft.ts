/**
 * Upstream Microsoft sign-in for a **public (native) client**.
 *
 * In Mode A (local stdio) this is the ONLY OAuth we implement — the Claude <->
 * server hop is trusted (local subprocess), so no downstream PRM/DCR/PKCE is
 * needed. The flow is OAuth 2.0 authorization-code + **PKCE** against the
 * `common` authority (personal + work/school accounts), with a **loopback**
 * redirect and **no client secret**.
 *
 * The pure building blocks (PKCE, authorize-URL, token-request/response) are
 * exported for unit testing; the interactive browser + loopback listener is
 * exercised manually (documented in the README).
 */

import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import { createServer } from "node:http";
import { z } from "zod";
import type { AppConfig } from "../config.js";
import { AuthError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";
import type { TokenSet, TokenStore } from "./tokens.js";

/** Refresh a bit before the access token actually expires. */
const DEFAULT_MIN_TTL_MS = 60_000;
/** How long to wait for the user to complete the browser sign-in. */
const AUTH_TIMEOUT_MS = 5 * 60_000;

// ── PKCE ─────────────────────────────────────────────────────────────────────

function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export interface Pkce {
  verifier: string;
  challenge: string;
}

/** Generate a PKCE verifier (43 chars) and its S256 challenge. */
export function generatePkce(): Pkce {
  const verifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(verifier).digest());
  return { verifier, challenge };
}

// ── Authorize URL ────────────────────────────────────────────────────────────

export interface AuthorizeUrlParams {
  authority: string;
  clientId: string;
  redirectUri: string;
  scopes: string[];
  state: string;
  codeChallenge: string;
  prompt?: string;
}

/** Build the Microsoft authorize URL for the auth-code + PKCE flow. */
export function buildAuthorizeUrl(p: AuthorizeUrlParams): string {
  const u = new URL(`${p.authority}/oauth2/v2.0/authorize`);
  u.searchParams.set("client_id", p.clientId);
  u.searchParams.set("response_type", "code");
  u.searchParams.set("redirect_uri", p.redirectUri);
  u.searchParams.set("response_mode", "query");
  u.searchParams.set("scope", p.scopes.join(" "));
  u.searchParams.set("state", p.state);
  u.searchParams.set("code_challenge", p.codeChallenge);
  u.searchParams.set("code_challenge_method", "S256");
  u.searchParams.set("prompt", p.prompt ?? "select_account");
  return u.toString();
}

// ── Token requests / responses ───────────────────────────────────────────────

/** Form body for exchanging an authorization code for tokens. */
export function buildCodeExchangeBody(p: {
  clientId: string;
  code: string;
  redirectUri: string;
  codeVerifier: string;
  scopes: string[];
}): URLSearchParams {
  return new URLSearchParams({
    client_id: p.clientId,
    grant_type: "authorization_code",
    code: p.code,
    redirect_uri: p.redirectUri,
    code_verifier: p.codeVerifier,
    scope: p.scopes.join(" "),
  });
}

/** Form body for refreshing tokens. */
export function buildRefreshBody(p: {
  clientId: string;
  refreshToken: string;
  scopes: string[];
}): URLSearchParams {
  return new URLSearchParams({
    client_id: p.clientId,
    grant_type: "refresh_token",
    refresh_token: p.refreshToken,
    scope: p.scopes.join(" "),
  });
}

const TokenResponseSchema = z.object({
  access_token: z.string().min(1),
  refresh_token: z.string().optional(),
  expires_in: z.number(),
  scope: z.string().optional(),
  token_type: z.string().optional(),
  id_token: z.string().optional(),
});
export type TokenResponse = z.infer<typeof TokenResponseSchema>;

/** Best-effort extraction of a username from an id_token (non-verified hint). */
export function decodeIdTokenUsername(idToken: string | undefined): string | undefined {
  if (!idToken) return undefined;
  const parts = idToken.split(".");
  if (parts.length < 2) return undefined;
  try {
    const payload = JSON.parse(Buffer.from(parts[1]!, "base64").toString("utf8")) as Record<
      string,
      unknown
    >;
    const candidate = payload.preferred_username ?? payload.email ?? payload.upn;
    return typeof candidate === "string" ? candidate : undefined;
  } catch {
    return undefined;
  }
}

/** Convert a raw token response into a {@link TokenSet}. */
export function toTokenSet(
  resp: TokenResponse,
  now: number,
  fallback?: { refreshToken?: string; scopes?: string[]; account?: string },
): TokenSet {
  const refreshToken = resp.refresh_token ?? fallback?.refreshToken;
  if (!refreshToken) {
    throw new AuthError(
      "Token response did not include a refresh token (offline_access missing?).",
    );
  }
  return {
    accessToken: resp.access_token,
    refreshToken,
    expiresAt: now + resp.expires_in * 1000,
    scopes: resp.scope ? resp.scope.split(" ") : (fallback?.scopes ?? []),
    tokenType: resp.token_type ?? "Bearer",
    account: decodeIdTokenUsername(resp.id_token) ?? fallback?.account,
  };
}

// ── Auth service ─────────────────────────────────────────────────────────────

export interface MicrosoftAuthDeps {
  config: AppConfig;
  store: TokenStore;
  /** Injectable fetch (defaults to global fetch). */
  fetchImpl?: typeof fetch;
  /** Injectable browser launcher (defaults to `open`). */
  openBrowser?: (url: string) => Promise<void>;
  /** Injectable clock (defaults to Date.now). */
  now?: () => number;
  logger?: Logger;
}

export interface GetAccessTokenOptions {
  /** Allow launching the browser for interactive sign-in (default true). */
  interactive?: boolean;
  /** Refresh if fewer than this many ms remain (default 60s). */
  minTtlMs?: number;
}

export interface MicrosoftAuth {
  getAccessToken(options?: GetAccessTokenOptions): Promise<string>;
  signIn(): Promise<TokenSet>;
  signOut(): Promise<void>;
  getAccount(): Promise<string | undefined>;
}

export function createMicrosoftAuth(deps: MicrosoftAuthDeps): MicrosoftAuth {
  const { config, store } = deps;
  const doFetch = deps.fetchImpl ?? fetch;
  const now = deps.now ?? Date.now;
  const log = deps.logger ?? createLogger({ name: "ms-auth" });
  const tokenEndpoint = `${config.authority}/oauth2/v2.0/token`;

  async function postToken(body: URLSearchParams): Promise<TokenResponse> {
    const res = await doFetch(tokenEndpoint, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
      body,
    });
    const text = await res.text();
    let json: unknown;
    try {
      json = JSON.parse(text);
    } catch {
      json = undefined;
    }
    if (!res.ok) {
      const err = json as { error?: string; error_description?: string } | undefined;
      // error_description can be verbose but does not contain the code/verifier.
      throw new AuthError(
        `Microsoft token endpoint returned ${res.status}: ${err?.error ?? "unknown_error"}`,
      );
    }
    const parsed = TokenResponseSchema.safeParse(json);
    if (!parsed.success) {
      throw new AuthError("Microsoft token response was malformed.");
    }
    return parsed.data;
  }

  async function refresh(current: TokenSet): Promise<TokenSet> {
    const resp = await postToken(
      buildRefreshBody({
        clientId: config.microsoftClientId,
        refreshToken: current.refreshToken,
        scopes: config.scopes,
      }),
    );
    return toTokenSet(resp, now(), {
      refreshToken: current.refreshToken,
      scopes: current.scopes,
      account: current.account,
    });
  }

  async function interactiveSignIn(): Promise<TokenSet> {
    const pkce = generatePkce();
    const state = base64url(randomBytes(16));
    return await new Promise<TokenSet>((resolve, reject) => {
      const server = createServer((req, res) => {
        const url = new URL(req.url ?? "/", `http://${config.loopbackHost}`);
        const code = url.searchParams.get("code");
        const returnedState = url.searchParams.get("state");
        const oauthError = url.searchParams.get("error");
        if (!code && !oauthError) {
          // e.g. favicon.ico — ignore.
          res.statusCode = 404;
          res.end();
          return;
        }
        // Respond to the browser first, then finish the exchange.
        res.setHeader("content-type", "text/html; charset=utf-8");
        if (oauthError) {
          res.end(page("Sign-in failed", "You can close this tab and try again."));
          cleanup();
          reject(new AuthError(`Authorization failed: ${oauthError}`));
          return;
        }
        if (!returnedState || !safeEqual(returnedState, state)) {
          res.end(page("Sign-in error", "State mismatch. Please try again."));
          cleanup();
          reject(new AuthError("OAuth state mismatch (possible CSRF)."));
          return;
        }
        res.end(page("Signed in", "You're all set — you can close this tab and return to Claude."));
        void exchange(code!);
      });

      const timer = setTimeout(() => {
        cleanup();
        reject(new AuthError("Timed out waiting for Microsoft sign-in."));
      }, AUTH_TIMEOUT_MS);

      function cleanup(): void {
        clearTimeout(timer);
        server.close();
      }

      async function exchange(code: string): Promise<void> {
        try {
          const address = server.address();
          const port = typeof address === "object" && address ? address.port : 0;
          const redirectUri = `http://localhost:${port}`;
          const resp = await postToken(
            buildCodeExchangeBody({
              clientId: config.microsoftClientId,
              code,
              redirectUri,
              codeVerifier: pkce.verifier,
              scopes: config.scopes,
            }),
          );
          const tokens = toTokenSet(resp, now(), { scopes: config.scopes });
          cleanup();
          resolve(tokens);
        } catch (e) {
          cleanup();
          reject(e instanceof AuthError ? e : new AuthError("Token exchange failed.", e));
        }
      }

      server.on("error", (e) => {
        cleanup();
        reject(new AuthError("Failed to start the local sign-in listener.", e));
      });

      server.listen(0, config.loopbackHost, () => {
        const address = server.address();
        const port = typeof address === "object" && address ? address.port : 0;
        const redirectUri = `http://localhost:${port}`;
        const authorizeUrl = buildAuthorizeUrl({
          authority: config.authority,
          clientId: config.microsoftClientId,
          redirectUri,
          scopes: config.scopes,
          state,
          codeChallenge: pkce.challenge,
        });
        log.info("Opening browser for Microsoft sign-in", { redirectUri });
        const launch = deps.openBrowser
          ? deps.openBrowser(authorizeUrl)
          : import("open").then((m) => m.default(authorizeUrl).then(() => undefined));
        launch.catch(() => {
          log.warn("Could not open a browser automatically. Visit the URL printed on stderr.");
          // Print to stderr so the user can copy it (never stdout).
          process.stderr.write(`\nOpen this URL to sign in:\n${authorizeUrl}\n\n`);
        });
      });
    });
  }

  return {
    async getAccessToken(options: GetAccessTokenOptions = {}): Promise<string> {
      const interactive = options.interactive ?? true;
      const minTtlMs = options.minTtlMs ?? DEFAULT_MIN_TTL_MS;

      let cached: TokenSet | null = null;
      try {
        cached = await store.load();
      } catch (e) {
        // Corrupt cache: reset it and continue to (re)authenticate.
        log.warn("Token cache unreadable; clearing", { error: (e as Error).message });
        await store.clear();
      }

      if (cached && cached.expiresAt - now() > minTtlMs) {
        return cached.accessToken;
      }

      if (cached?.refreshToken) {
        try {
          const refreshed = await refresh(cached);
          await store.save(refreshed);
          return refreshed.accessToken;
        } catch (e) {
          log.warn("Silent token refresh failed", { error: (e as Error).message });
          if (!interactive) throw e instanceof AuthError ? e : new AuthError("Refresh failed.", e);
        }
      }

      if (!interactive) {
        throw new AuthError("Not signed in. Interactive Microsoft sign-in is required.");
      }

      const tokens = await interactiveSignIn();
      await store.save(tokens);
      return tokens.accessToken;
    },

    async signIn(): Promise<TokenSet> {
      const tokens = await interactiveSignIn();
      await store.save(tokens);
      return tokens;
    },

    async signOut(): Promise<void> {
      await store.clear();
    },

    async getAccount(): Promise<string | undefined> {
      try {
        return (await store.load())?.account;
      } catch {
        return undefined;
      }
    },
  };
}

function safeEqual(a: string, b: string): boolean {
  const ba = Buffer.from(a);
  const bb = Buffer.from(b);
  return ba.length === bb.length && timingSafeEqual(ba, bb);
}

function page(title: string, body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:system-ui;padding:2rem"><h2>${title}</h2><p>${body}</p></body></html>`;
}
