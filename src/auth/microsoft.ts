/**
 * Upstream Microsoft sign-in for a **public (native) client**.
 *
 * In Mode A (local stdio) this is the ONLY OAuth we implement — the Claude <->
 * server hop is trusted (local subprocess), so no downstream PRM/DCR/PKCE is
 * needed. The flow is OAuth 2.0 authorization-code + **PKCE** against the
 * `common` authority (personal + work/school accounts), with a **loopback**
 * redirect and **no client secret**.
 *
 * Sign-in runs either explicitly (`login`) or on demand when a tool needs a
 * token and none is valid (see `GetAccessTokenOptions.interactiveWaitMs`).
 * The pure building blocks (PKCE, authorize-URL, token-request/response) are
 * exported for unit testing; the loopback listener is tested by driving it
 * over HTTP with an injected browser launcher.
 */

import type { ChildProcess } from "node:child_process";
import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import { createServer } from "node:http";
import { z } from "zod";
import type { AppConfig } from "../config.js";
import { AuthError, LOGIN_HINT, SignedOutError, SignInRequiredError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";
import {
  createMemoryCoordinator,
  type AuthCoordinator,
  type SharedAuthState,
} from "./coordination.js";
import type { TokenSet, TokenStore } from "./tokens.js";

/** Refresh a bit before the access token actually expires. */
const DEFAULT_MIN_TTL_MS = 60_000;
/** How long to wait for the user to complete the browser sign-in. */
const AUTH_TIMEOUT_MS = 5 * 60_000;
/**
 * After an on-demand sign-in fails or is cancelled, don't open another browser
 * window for this long, so automatic model retries can't loop on the user.
 */
export const SIGN_IN_RETRY_COOLDOWN_MS = 30_000;

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
  /**
   * If the browser can't be opened, print the sign-in URL on stderr and keep
   * waiting (for the `login` command, where a person is at the terminal).
   * Otherwise (default; the MCP server) the attempt ends at once, so a retry
   * isn't stuck joining a sign-in that can never finish.
   */
  printUrlIfBrowserFails?: boolean;
  /** Injectable clock (defaults to Date.now). */
  now?: () => number;
  logger?: Logger;
  /** Limit on one token-endpoint request, body included (default 15s, #37). */
  tokenTimeoutMs?: number;
  /**
   * Orders saves and sign-outs with other processes sharing the token cache
   * (#95; see coordination.ts). Defaults to in-process only.
   */
  coordinator?: AuthCoordinator;
}

const DEFAULT_TOKEN_TIMEOUT_MS = 15_000;

/** Another process's browser sign-in finished first; this one's tokens are dropped. */
class SignedInElsewhereError extends Error {}

export interface GetAccessTokenOptions {
  /** Allow launching the browser for interactive sign-in (default true). */
  interactive?: boolean;
  /**
   * With `interactive`: wait at most this long for the browser sign-in, then
   * fail with a {@link SignInRequiredError} telling the user to finish it. The
   * sign-in carries on in the background, and later calls join it. Unset:
   * wait for it to finish (or time out after 5 minutes).
   */
  interactiveWaitMs?: number;
  /** Refresh if fewer than this many ms remain (default 60s). */
  minTtlMs?: number;
  /**
   * The sign-out epoch of the operation this lookup belongs to (from
   * `signOutEpoch()` when the operation began). A Graph request makes several
   * lookups (retries, 401 recovery); passing its epoch makes one that outlived
   * a sign-out stop instead of signing in again. Defaults to "now".
   */
  signOutEpoch?: number;
}

export interface MicrosoftAuth {
  getAccessToken(options?: GetAccessTokenOptions): Promise<string>;
  signIn(): Promise<TokenSet>;
  signOut(): Promise<void>;
  getAccount(): Promise<string | undefined>;
  /** Current sign-out epoch: note it when an operation starts (see `signOutEpoch` above). */
  signOutEpoch(): number;
  /**
   * Catch up with sign-outs in other processes (#95) and return the epoch for
   * an operation starting now: use it rather than `signOutEpoch()`. A sign-out
   * elsewhere that is first seen later ends the operations of the epochs
   * before it; a local sign-out during the catch-up ends this one.
   */
  syncSignOuts(): Promise<number>;
}

export function createMicrosoftAuth(deps: MicrosoftAuthDeps): MicrosoftAuth {
  const { config, store } = deps;
  const doFetch = deps.fetchImpl ?? fetch;
  const now = deps.now ?? Date.now;
  const log = deps.logger ?? createLogger({ name: "ms-auth" });
  const coordinator = deps.coordinator ?? createMemoryCoordinator();
  const tokenEndpoint = `${config.authority}/oauth2/v2.0/token`;

  async function postToken(body: URLSearchParams): Promise<TokenResponse> {
    // A stalled token endpoint must not hang the tool call that needed a token:
    // the limit covers the wait for headers and the body.
    const timeoutMs = deps.tokenTimeoutMs ?? DEFAULT_TOKEN_TIMEOUT_MS;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let res: Response;
    let text: string;
    try {
      res = await doFetch(tokenEndpoint, {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          accept: "application/json",
        },
        body,
        signal: controller.signal,
      });
      text = await res.text();
    } catch (e) {
      throw new AuthError(
        controller.signal.aborted
          ? `Microsoft token endpoint didn't answer within ${timeoutMs}ms.`
          : "Couldn't reach the Microsoft token endpoint.",
        e,
      );
    } finally {
      clearTimeout(timer);
    }
    let json: unknown;
    try {
      json = JSON.parse(text);
    } catch {
      json = undefined;
    }
    if (!res.ok) {
      const err = json as { error?: string; error_description?: string } | undefined;
      // Deliberately surface only the short error code, never error_description
      // (verbose upstream detail we don't want echoed into the thrown message).
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
    const refreshed = toTokenSet(resp, now(), {
      refreshToken: current.refreshToken,
      scopes: current.scopes,
      account: current.account,
    });
    // Still the same sign-in: it keeps its generation.
    return current.signInGeneration === undefined
      ? refreshed
      : { ...refreshed, signInGeneration: current.signInGeneration };
  }

  // Single-flight guard: collapse concurrent refreshes into one network call.
  // MSA refresh tokens rotate, so two parallel refreshes would spend the same
  // token twice — the loser gets `invalid_grant` and their `store.save()` could
  // clobber the freshly-rotated set. Concurrent callers await the same promise.
  // ── Sign-out ordering ──────────────────────────────────────────────────────
  // `signOuts` is bumped synchronously by signOut(). Every request notes it
  // when it starts, and anything still in flight from before a sign-out
  // (a sign-in, a refresh, a tool call's token lookup) stops with
  // SignedOutError instead of saving tokens or falling back to a new sign-in.
  // Token saves and the sign-out cleanup run one at a time: a save already
  // writing finishes first (and the cleanup then removes it), while a save
  // queued after the sign-out sees it and is dropped.
  let signOuts = 0;
  // The sign-out cleanup in progress, if any. Requests that start meanwhile
  // already belong to the new epoch, so they must wait for it rather than
  // read (and refresh) the credentials being deleted.
  let pendingSignOut: Promise<void> | null = null;
  // The shared sign-out count last seen here (#95). A rise in it is another
  // process's sign-out, which ends everything in flight here just as a local
  // one does (see `observe`).
  let sharedSignOutsSeen: number | undefined;
  // signOut() calls in this process, as opposed to sign-outs seen from elsewhere.
  let localSignOuts = 0;
  let commits: Promise<unknown> = Promise.resolve();
  function serialized<T>(op: () => Promise<T>): Promise<T> {
    const run = commits.then(op, op);
    commits = run.catch(() => undefined);
    return run;
  }
  /**
   * Save tokens from an operation that began at sign-out epoch `epoch` (this
   * process) with the shared counters `shared` (all processes, #95). Under the
   * cross-process lock: a sign-out anywhere since then wins (SignedOutError),
   * and so does a browser sign-in that finished since (SignedInElsewhereError),
   * so neither a late sign-in nor a slow refresh of the old account can
   * replace the account just signed in.
   */
  function saveUnlessSignedOut(
    tokens: TokenSet,
    epoch: number,
    shared: SharedAuthState,
    kind: "refresh" | "sign-in",
    onSaved?: () => void,
  ) {
    return serialized(async () => {
      if (signOuts !== epoch) throw new SignedOutError();
      await coordinator.withLock(async (current, persist) => {
        observe(current);
        if (current.signOuts !== shared.signOuts) throw new SignedOutError();
        if (current.signIns !== shared.signIns) throw new SignedInElsewhereError();
        // A sign-in's count goes up before its tokens are saved: if it can't be
        // stored, nothing is, so a new account is never cached under the old
        // count for a pending refresh of the old one to overwrite. (Lookups
        // read the counters and the cache together under this lock, so none
        // can pair the new count with the tokens being replaced.)
        if (kind === "sign-in") {
          const signIns = current.signIns + 1;
          await persist({ ...current, signIns });
          // Marked with its generation, so a sign-in that loses to this one can
          // tell whether this save happened (see savedElsewhere).
          await store.save({ ...tokens, signInGeneration: signIns });
        } else {
          await store.save(tokens);
        }
        return { value: undefined };
      });
      onSaved?.();
    });
  }

  /**
   * After SignedInElsewhereError: the sign-in that another process saved, read
   * under the lock (so a sign-out can't slip in between), if it is still for
   * this operation (no sign-out since `shared` and `epoch`) and usable.
   */
  async function savedElsewhere(epoch: number, shared: SharedAuthState): Promise<TokenSet | null> {
    const saved = await coordinator.withLock(async (current) => {
      observe(current);
      if (current.signOuts !== shared.signOuts) throw new SignedOutError();
      const found = await store.load().catch(() => null);
      // Only the sign-in that won: a raised count whose save then failed
      // leaves an older account cached, which is no one's answer now
      // (Copilot review on #126).
      return { value: found?.signInGeneration === current.signIns ? found : null };
    });
    if (signOuts !== epoch) throw new SignedOutError();
    return saved && saved.expiresAt - now() > DEFAULT_MIN_TTL_MS ? saved : null;
  }

  /**
   * End everything in flight here: called at a sign-out, synchronously, and
   * when one in another process is first seen. Requests compare the epoch they
   * noted with `signOuts`, so they all stop (SignedOutError) rather than save,
   * join a now-useless browser sign-in, or start a new one for an operation from
   * before the sign-out.
   */
  function invalidate(): void {
    signOuts += 1;
    lastSignedIn = undefined;
    // Cancel and detach a pending browser sign-in, so its listener closes and
    // a fresh sign-in can start straight away instead of joining it.
    inflightAbort?.abort();
    inflightSignIn = null;
    inflightAbort = null;
    inflightRefresh = null;
  }

  /** Note the shared counters; a sign-out elsewhere since they were last seen invalidates. */
  function observe(state: SharedAuthState): SharedAuthState {
    if (sharedSignOutsSeen !== undefined && state.signOuts > sharedSignOutsSeen) invalidate();
    if (sharedSignOutsSeen === undefined || state.signOuts > sharedSignOutsSeen) {
      sharedSignOutsSeen = state.signOuts;
    }
    return state;
  }

  async function sharedState(): Promise<SharedAuthState> {
    // An unlocked read can catch this process's own sign-out between
    // publishing its count and noting it as seen, and would take it for one
    // from elsewhere. So wait out a local sign-out's cleanup first, and read
    // again if one started meanwhile (Copilot review on #126). Reads under the
    // lock can't see that in-between state: the cleanup holds it throughout.
    for (;;) {
      while (pendingSignOut) await pendingSignOut.catch(() => undefined);
      const signOutsBefore = localSignOuts;
      const state = await coordinator.snapshot();
      if (!pendingSignOut && localSignOuts === signOutsBefore) return observe(state);
    }
  }

  let inflightRefresh: { epoch: number; id: object; promise: Promise<TokenSet> } | null = null;
  function refreshAndSave(
    current: TokenSet,
    epoch: number,
    shared: SharedAuthState,
  ): Promise<TokenSet> {
    // Only join a refresh from the same sign-out epoch.
    if (inflightRefresh?.epoch !== epoch) {
      const id = {}; // identity of this refresh, for its own finalizer
      const promise = (async () => {
        try {
          const refreshed = await refresh(current);
          await saveUnlessSignedOut(refreshed, epoch, shared, "refresh");
          return refreshed;
        } finally {
          if (inflightRefresh?.id === id) inflightRefresh = null;
        }
      })();
      inflightRefresh = { epoch, id, promise };
    }
    return inflightRefresh.promise;
  }

  async function interactiveSignIn(signal?: AbortSignal): Promise<TokenSet> {
    const pkce = generatePkce();
    const state = base64url(randomBytes(16));
    return await new Promise<TokenSet>((resolve, reject) => {
      if (signal?.aborted) {
        reject(new SignedOutError());
        return;
      }
      // Sign-out cancels the attempt: close the listener now rather than
      // leaving an orphaned flow pending until its timeout.
      signal?.addEventListener(
        "abort",
        () => {
          cleanup();
          reject(new SignedOutError());
        },
        { once: true },
      );
      let completed = false;
      const server = createServer((req, res) => {
        const url = new URL(req.url ?? "/", `http://${config.loopbackHost}`);
        const code = url.searchParams.get("code");
        const returnedState = url.searchParams.get("state");
        const oauthError = url.searchParams.get("error");
        if (completed || (!code && !oauthError)) {
          // e.g. favicon.ico, or anything after the real callback — ignore.
          res.statusCode = 404;
          res.end();
          return;
        }
        res.setHeader("content-type", "text/html; charset=utf-8");
        // Only Microsoft's redirect carries our random state. Anything else on
        // this port (a stray or hostile local request) gets a 400 and is ignored:
        // it must neither abort the pending sign-in nor complete it (#41).
        if (!returnedState || !safeEqual(returnedState, state)) {
          res.statusCode = 400;
          res.end(page("Sign-in error", "This request doesn't belong to the current sign-in."));
          log.warn("Ignored a loopback request with a missing or wrong state");
          return;
        }
        completed = true;
        // Respond to the browser first, then finish the exchange.
        if (oauthError) {
          res.end(page("Sign-in failed", "You can close this tab and try again."));
          cleanup();
          reject(new AuthError(`Authorization failed: ${oauthError}`));
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
          : import("open").then((m) => m.default(authorizeUrl).then(launcherSucceeded));
        launch.catch((e: unknown) => {
          // The browser evidently worked: Microsoft's redirect has arrived and the
          // token exchange is under way, so a late launcher failure is irrelevant.
          if (completed) return;
          if (deps.printUrlIfBrowserFails) {
            log.warn("Could not open a browser automatically. Visit the URL printed on stderr.");
            // Print to stderr so the user can copy it (never stdout).
            process.stderr.write(`\nOpen this URL to sign in:\n${authorizeUrl}\n\n`);
            return;
          }
          log.warn("Could not open a browser for Microsoft sign-in", {
            error: e instanceof Error ? e.message : String(e),
          });
          cleanup();
          reject(
            new SignInRequiredError(
              "Couldn't open a web browser for the Microsoft sign-in. Check that a default web " +
                `browser is set, then ask me to try again ${LOGIN_HINT}.`,
            ),
          );
        });
      });
    });
  }

  // Single-flight interactive sign-in. An on-demand caller may stop waiting
  // (`interactiveWaitMs`) while the browser flow carries on; later callers join
  // the same attempt rather than opening another window.
  let inflightSignIn: Promise<TokenSet> | null = null;
  let inflightAbort: AbortController | null = null;
  let lastSignInFailureAt: number | undefined;
  // Completed sign-ins in this process, with the latest result: lets a caller
  // whose refresh failed notice that a sign-in finished meanwhile, even if its
  // own cache read raced with that save.
  let signInsCompleted = 0;
  let lastSignedIn: TokenSet | undefined;
  /**
   * Start (or join) the browser sign-in. `since` is the shared counters noted
   * when the requesting operation began, so a sign-out in another process after
   * that point stops this attempt from saving (#95); without it (an explicit
   * `signIn()`), they are noted now.
   */
  function signInAndSave(since?: SharedAuthState): Promise<TokenSet> {
    if (!inflightSignIn) {
      const epoch = signOuts;
      // Also this attempt's identity: sign-out replaces it, so `inflightAbort ===
      // abort` means "this attempt is still the current one".
      const abort = new AbortController();
      const attempt = (async () => {
        const succeeded = (tokens: TokenSet) => {
          lastSignInFailureAt = undefined;
          lastSignedIn = tokens;
          signInsCompleted += 1;
        };
        try {
          let shared = since;
          if (!shared) {
            // Right after a local sign-out, let its cleanup (and its shared
            // count) land first, so this attempt belongs to the new epoch.
            while (pendingSignOut) await pendingSignOut.catch(() => undefined);
            if (signOuts !== epoch) throw new SignedOutError();
            // Noted before the browser opens: a sign-out or sign-in in another
            // process from here on decides whether this one may save (#95).
            shared = await sharedState();
            if (signOuts !== epoch) throw new SignedOutError();
          }
          const tokens = await interactiveSignIn(abort.signal);
          try {
            await saveUnlessSignedOut(tokens, epoch, shared, "sign-in", () => succeeded(tokens));
            return tokens;
          } catch (e) {
            if (!(e instanceof SignedInElsewhereError)) throw e;
            // Another process's sign-in finished first and is saved: use that
            // one rather than replacing it with this (possibly other) account.
            const saved = await savedElsewhere(epoch, shared);
            if (!saved) {
              // The other sign-in's save failed, or its tokens are no longer
              // usable: neither it nor this one is signed in, so say try again.
              throw new SignInRequiredError(
                "Another Microsoft sign-in was in progress at the same time, so this one wasn't " +
                  "saved. Ask me to try again.",
              );
            }
            log.info("Using the sign-in that finished first in another process");
            succeeded(saved);
            return saved;
          }
        } catch (e) {
          // Only the current attempt's genuine failures start the cooldown: not
          // a sign-out, and not an attempt that sign-out already detached.
          if (!(e instanceof SignedOutError) && inflightAbort === abort) {
            lastSignInFailureAt = now();
          }
          throw e;
        } finally {
          if (inflightAbort === abort) {
            inflightSignIn = null;
            inflightAbort = null;
          }
        }
      })();
      // A caller that stopped waiting no longer observes the result; never let
      // a later failure surface as an unhandled rejection.
      attempt.catch(() => undefined);
      inflightSignIn = attempt;
      inflightAbort = abort;
    }
    return inflightSignIn;
  }

  async function signInOnDemand(
    waitMs: number | undefined,
    epoch: number,
    shared: SharedAuthState,
  ): Promise<string> {
    // A request from before a sign-out, here or in another process (#95), must
    // not start or join a new sign-in after it, nor take a token from one.
    const ensureStillCurrent = async () => {
      if (signOuts !== epoch) throw new SignedOutError();
      const latest = await sharedState();
      // Both after the await: a local sign-out may have happened during it.
      if (signOuts !== epoch || latest.signOuts !== shared.signOuts) throw new SignedOutError();
    };
    await ensureStillCurrent();
    const joining = inflightSignIn !== null;
    if (
      !joining &&
      lastSignInFailureAt !== undefined &&
      now() - lastSignInFailureAt < SIGN_IN_RETRY_COOLDOWN_MS
    ) {
      throw new SignInRequiredError(
        `Microsoft sign-in didn't complete. Ask me to try again in a minute ${LOGIN_HINT}.`,
      );
    }
    const pending = signInAndSave(shared);
    const tokens = waitMs === undefined ? await pending : await settleWithin(pending, waitMs);
    if (tokens) {
      await ensureStillCurrent();
      return tokens.accessToken;
    }
    throw new SignInRequiredError(
      joining
        ? "Still waiting for you to finish the Microsoft sign-in in your browser. " +
            "Once you have, ask me to try again."
        : "A Microsoft sign-in window has opened in your browser. Finish signing in there, " +
            "then ask me to try again. If no browser window appeared, check that a default web " +
            `browser is set and ask me again ${LOGIN_HINT}.`,
    );
  }

  return {
    async getAccessToken(options: GetAccessTokenOptions = {}): Promise<string> {
      const interactive = options.interactive ?? true;
      const minTtlMs = options.minTtlMs ?? DEFAULT_MIN_TTL_MS;
      // Noted before any await (or taken from the calling operation): if the user
      // signs out while this request is in flight, including while it waits
      // below, it stops (see "Sign-out ordering").
      let epoch = options.signOutEpoch ?? signOuts;
      const localAtStart = localSignOuts;
      const ensureNotSignedOut = () => {
        if (signOuts !== epoch) throw new SignedOutError();
      };
      ensureNotSignedOut();
      // A sign-out still cleaning up: wait, so this request can't read the
      // credentials being deleted and refresh them back into existence.
      while (pendingSignOut) await pendingSignOut.catch(() => undefined);
      ensureNotSignedOut();
      const signInsAtStart = signInsCompleted;
      // The counters and the cache are read together, under the cross-process
      // lock (#95). A sign-in or sign-out elsewhere is then wholly before or
      // after this read: it can't pair the new counters with the credentials
      // being replaced or deleted (and wait for a sign-out's cleanup if one is
      // running). A sign-out after this point stops a refresh of what was read
      // from saving it back.
      const { shared, cached } = await coordinator.withLock(async (current) => {
        observe(current);
        let found: TokenSet | null = null;
        try {
          found = await store.load();
        } catch (e) {
          // Corrupt cache: reset it (under the lock) and go on to authenticate.
          log.warn("Token cache unreadable; clearing", { error: (e as Error).message });
          await store.clear();
        }
        return { value: { shared: current, cached: found } };
      });
      // A lookup that isn't part of a larger operation starts from what it first
      // sees: a sign-out elsewhere that this read is the first to notice came
      // before it. (An operation's own epoch was noted after syncSignOuts().)
      if (options.signOutEpoch === undefined && localSignOuts === localAtStart) epoch = signOuts;
      ensureNotSignedOut();

      if (cached && cached.expiresAt - now() > minTtlMs) {
        return cached.accessToken;
      }

      // A browser sign-in is already pending (e.g. a previous call's window):
      // join it, rather than racing it with a refresh whose failure could open
      // a second window after the user has just signed in.
      if (interactive && inflightSignIn) {
        return signInOnDemand(options.interactiveWaitMs, epoch, shared);
      }

      if (cached?.refreshToken) {
        try {
          const refreshed = await refreshAndSave(cached, epoch, shared);
          ensureNotSignedOut();
          return refreshed.accessToken;
        } catch (e) {
          // A sign-out mid-request ends it: no interactive fallback.
          if (e instanceof SignedOutError) throw e;
          // A browser sign-in finished elsewhere while we refreshed the old
          // account: use that sign-in rather than putting the old one back.
          if (e instanceof SignedInElsewhereError) {
            const saved = await savedElsewhere(epoch, shared);
            if (saved) return saved.accessToken;
          }
          log.warn("Silent token refresh failed", { error: (e as Error).message });
          if (!interactive) throw e instanceof AuthError ? e : new AuthError("Refresh failed.", e);
        }
      }
      ensureNotSignedOut();
      // A sign-out in another process since this request began ends it too,
      // before it can use a later sign-in or fall back to a browser one (#95).
      await sharedState();
      ensureNotSignedOut();

      if (!interactive) {
        throw new AuthError("Not signed in. Interactive Microsoft sign-in is required.");
      }

      // A sign-in may have completed while the refresh was in flight: use it.
      // A *different* token than the one we started with is new, so judge it by
      // the normal threshold. The 401-recovery path forces a refresh with a huge
      // minTtlMs, which no fresh token could otherwise satisfy.
      const latest = await store.load().catch(() => null);
      // This read may also have raced a sign-out elsewhere: catch up again
      // before using what it found (or this process's own latest sign-in).
      await sharedState();
      ensureNotSignedOut();
      if (latest) {
        const isNew = latest.accessToken !== cached?.accessToken;
        const threshold = isNew ? Math.min(minTtlMs, DEFAULT_MIN_TTL_MS) : minTtlMs;
        if (latest.expiresAt - now() > threshold) return latest.accessToken;
      }
      // That read may itself have been stale: if a sign-in in this process
      // completed since we started, use its tokens rather than opening another.
      if (
        signInsCompleted !== signInsAtStart &&
        lastSignedIn &&
        lastSignedIn.expiresAt - now() > DEFAULT_MIN_TTL_MS
      ) {
        return lastSignedIn.accessToken;
      }

      return signInOnDemand(options.interactiveWaitMs, epoch, shared);
    },

    async signIn(): Promise<TokenSet> {
      // Catch up first, so this belongs to the current epoch and doesn't join a
      // pending sign-in that a sign-out elsewhere has made useless. A local
      // sign-out during the catch-up came after this call, so it ends it.
      const local = localSignOuts;
      await sharedState();
      if (localSignOuts !== local) throw new SignedOutError();
      return signInAndSave();
    },

    async signOut(): Promise<void> {
      // Synchronously, before any await: everything in flight is now stale.
      localSignOuts += 1;
      invalidate();
      // Cache and key: nothing of the sign-in stays on disk. Queued behind any
      // save already in progress, so that save can't land after the cleanup.
      // Under the cross-process lock the shared sign-out count goes up too, even
      // if the cleanup fails, so operations that began before it in any process
      // won't save afterwards (#95).
      const cleanup = serialized(async () => {
        const failure = await coordinator.withLock(async (current, persist) => {
          observe(current);
          // The shared count goes up before anything is deleted. If it can't be
          // stored, nothing is, and the sign-out fails: otherwise an operation
          // elsewhere could still pass the old count and save the sign-in again.
          const next = { ...current, signOuts: current.signOuts + 1 };
          await persist(next);
          // This sign-out's own bump, already acted on above, isn't one from elsewhere.
          sharedSignOutsSeen = next.signOuts;
          let error: unknown;
          try {
            await store.clear({ includeKey: true });
          } catch (e) {
            error = e ?? new Error("Sign-out cleanup failed.");
          }
          return { value: error };
        });
        if (failure !== undefined) throw failure;
      });
      pendingSignOut = cleanup;
      try {
        await cleanup;
      } finally {
        if (pendingSignOut === cleanup) pendingSignOut = null;
      }
    },

    signOutEpoch(): number {
      return signOuts;
    },

    async syncSignOuts(): Promise<number> {
      const epoch = signOuts;
      const local = localSignOuts;
      await sharedState();
      // Sign-outs elsewhere seen by the catch-up came before the operation; a
      // local one during it came after, so the operation keeps the earlier epoch
      // and stops.
      return localSignOuts === local ? signOuts : epoch;
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

/** How long a launcher (xdg-open, `start`, `open`) gets to fail before we assume it worked. */
const LAUNCHER_GRACE_MS = 5_000;

/**
 * `open()` resolves once the launcher process has *started*. A launcher can
 * still fail afterwards, e.g. `xdg-open` with no browser configured. Treat a
 * non-zero exit within the grace period as a failed launch. A launcher still
 * running (or exiting 0) counts as success; never wait for the browser itself.
 */
export function launcherSucceeded(
  child: Pick<ChildProcess, "once" | "exitCode">,
  graceMs = LAUNCHER_GRACE_MS,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const fail = (code: number | null) =>
      reject(new Error(`Browser launcher exited with code ${code}.`));
    if (child.exitCode !== null) {
      if (child.exitCode === 0) resolve();
      else fail(child.exitCode);
      return;
    }
    const timer = setTimeout(resolve, graceMs);
    timer.unref?.();
    child.once("exit", (code: number | null, signal?: string | null) => {
      clearTimeout(timer);
      // Only a clean exit counts; a null code means it was killed by a signal.
      if (code === 0) resolve();
      else if (code === null)
        reject(new Error(`Browser launcher was terminated (${signal ?? "signal"}).`));
      else fail(code);
    });
    child.once("error", (e: Error) => {
      clearTimeout(timer);
      reject(e);
    });
  });
}

/** `p`'s value if it settles within `ms`, else `undefined` (rejections propagate). */
function settleWithin<T>(p: Promise<T>, ms: number): Promise<T | undefined> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => resolve(undefined), ms);
    p.then(
      (v) => {
        clearTimeout(timer);
        resolve(v);
      },
      (e: unknown) => {
        clearTimeout(timer);
        reject(e);
      },
    );
  });
}

function safeEqual(a: string, b: string): boolean {
  const ba = Buffer.from(a);
  const bb = Buffer.from(b);
  return ba.length === bb.length && timingSafeEqual(ba, bb);
}

function page(title: string, body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:system-ui;padding:2rem"><h2>${title}</h2><p>${body}</p></body></html>`;
}
