/**
 * The send confirmation page (#57). Sending mail is the one action that can
 * move mailbox content out, so it isn't left to the model or to the client's
 * approval prompt (which Claude Desktop doesn't show for extension tools,
 * #138): every email is shown to the user on a page served by the connector
 * itself, and it is sent only when they click Send there.
 *
 * The page is served on 127.0.0.1 at a one-time, unguessable URL that's opened
 * in the user's browser and never given to the model. It expires after a few
 * minutes; a decision can be made once. The tool waits a while for the click
 * and then reports that the email is waiting; a later click still sends it.
 */

import { randomBytes } from "node:crypto";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { clientSafeMessage, ValidationError } from "../util/errors.js";
import { createLogger, type Logger } from "../util/log.js";

export interface Recipient {
  address: string;
  name?: string;
  /** Whether this address has ever sent the user mail; null if that couldn't be checked. */
  known: boolean | null;
}

/** An email as it will be sent, which is exactly what the page shows. */
export interface OutgoingEmail {
  kind: "new" | "reply";
  to: Recipient[];
  cc: Recipient[];
  subject: string;
  /** Plain text. */
  body: string;
  /** For a reply: the message it answers. */
  inReplyTo?: { subject: string | null; from: string | null };
}

export type ConfirmOutcome =
  | { status: "sent" }
  | { status: "cancelled" }
  /** The user clicked Send, but sending failed; nothing was sent. */
  | { status: "failed"; message: string }
  /** No decision yet; the page stays open until it expires, and a click still sends. */
  | { status: "pending"; expiresAt: number };

export interface SendConfirmer {
  /**
   * Show `email` for confirmation and wait up to `waitMs` for the user's
   * decision. `send` runs only when they click Send.
   */
  confirm(email: OutgoingEmail, send: () => Promise<void>, waitMs: number): Promise<ConfirmOutcome>;
  /** Stop serving pages; any still pending are dropped (nothing is sent). */
  close(): Promise<void>;
}

export interface SendConfirmerDeps {
  openBrowser: (url: string) => Promise<void>;
  /** Loopback address to serve on (default 127.0.0.1). */
  host?: string;
  /** How long a page stays usable (default 5 minutes). */
  expiresMs?: number;
  /** Most confirmations pending at once (default 3). */
  maxPending?: number;
  now?: () => number;
  logger?: Logger;
}

export const CONFIRM_EXPIRES_MS = 5 * 60_000;
export const MAX_PENDING_CONFIRMATIONS = 3;
/** Longest request body accepted: the form only ever posts `decision=send|cancel`. */
const MAX_FORM_BYTES = 1_024;

const TOKEN_PATH = /^\/send\/([0-9a-f]{32})$/;

/** Headers on every page: nothing from elsewhere, no framing, nothing cached or leaked. */
const SECURITY_HEADERS: Record<string, string> = {
  "content-type": "text/html; charset=utf-8",
  "content-security-policy":
    "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; frame-ancestors 'none'; " +
    "base-uri 'none'",
  "x-frame-options": "DENY",
  "x-content-type-options": "nosniff",
  // Not no-referrer: under it, browsers send `Origin: null` on the page's own
  // form post. same-origin still keeps the one-time URL from other sites.
  "referrer-policy": "same-origin",
  "cache-control": "no-store",
};

interface Entry {
  email: OutgoingEmail;
  send: () => Promise<void>;
  expiresAt: number;
  state: "waiting" | "sending" | "done";
  /** Resolves with the decision (or the send's result). */
  decided: Promise<ConfirmOutcome>;
  settle: (outcome: ConfirmOutcome) => void;
  expiry: ReturnType<typeof setTimeout>;
}

export function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!,
  );
}

export function createSendConfirmer(deps: SendConfirmerDeps): SendConfirmer {
  const host = deps.host ?? "127.0.0.1";
  const expiresMs = deps.expiresMs ?? CONFIRM_EXPIRES_MS;
  const maxPending = deps.maxPending ?? MAX_PENDING_CONFIRMATIONS;
  const now = deps.now ?? Date.now;
  const log = deps.logger ?? createLogger({ name: "send" });
  const pending = new Map<string, Entry>();
  let server: Server | undefined;
  let listening: Promise<number> | undefined;

  /** The server, started on first use; it stops again once nothing is pending. */
  function port(): Promise<number> {
    listening ??= new Promise<number>((resolve, reject) => {
      server = createServer((req, res) => {
        handle(req, res).catch(() => {
          if (!res.headersSent) res.writeHead(500, SECURITY_HEADERS);
          res.end();
        });
      });
      server.once("error", reject);
      server.listen(0, host, () => {
        const address = server!.address();
        resolve(typeof address === "object" && address ? address.port : 0);
      });
    });
    return listening;
  }

  async function stopIfIdle(): Promise<void> {
    if (pending.size > 0 || !server) return;
    const s = server;
    server = undefined;
    listening = undefined;
    await new Promise<void>((resolve) => s.close(() => resolve()));
  }

  function finish(token: string, entry: Entry, outcome: ConfirmOutcome): void {
    entry.state = "done";
    clearTimeout(entry.expiry);
    pending.delete(token);
    entry.settle(outcome);
    void stopIfIdle();
  }

  async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? "/", `http://${host}`);
    const token = TOKEN_PATH.exec(url.pathname)?.[1];
    const entry = token ? pending.get(token) : undefined;
    if (!token || !entry) {
      res.writeHead(404, SECURITY_HEADERS);
      res.end(
        page(
          "This confirmation isn't open",
          "<p>It has expired, was already used, or never existed. Nothing was sent from this page.</p>",
        ),
      );
      return;
    }
    if (req.method === "GET") {
      res.writeHead(200, SECURITY_HEADERS);
      res.end(confirmPage(entry.email, entry.expiresAt));
      return;
    }
    if (req.method !== "POST") {
      res.writeHead(405, { ...SECURITY_HEADERS, allow: "GET, POST" });
      res.end();
      return;
    }
    // Only this page may decide: refuse a post from another site, even though
    // it couldn't know the URL. A browser can still send `Origin: null` for
    // the page's own post (under some referrer policies); that's accepted only
    // when the browser itself marks the request same-origin, which another
    // site can't make it do.
    const origin = req.headers.origin;
    const fetchSite = req.headers["sec-fetch-site"];
    const ownOrigin = `http://${host}:${await port()}`;
    const sameOrigin = fetchSite === "same-origin";
    if (
      (origin && origin !== ownOrigin && !(origin === "null" && sameOrigin)) ||
      (fetchSite && !["same-origin", "none"].includes(String(fetchSite)))
    ) {
      // What the browser sent (no addresses or content) helps tell a real
      // cross-site post from a browser quirk.
      log.warn("Refused a confirmation post that didn't come from the page", {
        origin: String(origin ?? ""),
        fetchSite: String(fetchSite ?? ""),
      });
      res.writeHead(403, SECURITY_HEADERS);
      res.end(page("Not allowed", "<p>This request didn't come from the confirmation page.</p>"));
      return;
    }
    const decision = new URLSearchParams(await readBody(req)).get("decision");
    if (entry.state !== "waiting") {
      res.writeHead(409, SECURITY_HEADERS);
      res.end(page("Already decided", "<p>This email is already being sent.</p>"));
      return;
    }
    if (decision === "cancel") {
      log.info("Email not sent: the user cancelled it");
      finish(token, entry, { status: "cancelled" });
      res.writeHead(200, SECURITY_HEADERS);
      res.end(page("Not sent", "<p>The email wasn't sent. You can close this tab.</p>"));
      return;
    }
    if (decision !== "send") {
      res.writeHead(400, SECURITY_HEADERS);
      res.end(page("Not sent", "<p>Choose Send or Don't send.</p>"));
      return;
    }
    entry.state = "sending";
    try {
      await entry.send();
      log.info("Email sent after the user confirmed it", {
        recipients: entry.email.to.length + entry.email.cc.length,
      });
      finish(token, entry, { status: "sent" });
      res.writeHead(200, SECURITY_HEADERS);
      res.end(page("Sent", "<p>The email was sent. You can close this tab.</p>"));
    } catch (e) {
      const message = clientSafeMessage(e);
      // Counts only: Microsoft's error text can quote recipients or content.
      log.warn("A confirmed email didn't send cleanly", {
        recipients: entry.email.to.length + entry.email.cc.length,
      });
      finish(token, entry, { status: "failed", message });
      res.writeHead(502, SECURITY_HEADERS);
      // Neutral: the message itself says whether it may have gone anyway.
      res.end(page("Sending didn't complete", `<p>${escapeHtml(message)}</p>`));
    }
  }

  return {
    async confirm(email, send, waitMs) {
      if (pending.size >= maxPending) {
        throw new ValidationError(
          `${maxPending} emails are already waiting for confirmation in the browser. Send or ` +
            "cancel those first. Nothing was sent.",
        );
      }
      const token = randomBytes(16).toString("hex");
      let settle!: (o: ConfirmOutcome) => void;
      const decided = new Promise<ConfirmOutcome>((r) => (settle = r));
      const expiresAt = now() + expiresMs;
      const entry: Entry = {
        email,
        send,
        expiresAt,
        state: "waiting",
        decided,
        settle,
        expiry: setTimeout(() => {
          if (entry.state !== "waiting") return;
          log.info("A send confirmation expired; nothing was sent");
          pending.delete(token);
          entry.state = "done";
          void stopIfIdle();
        }, expiresMs),
      };
      entry.expiry.unref?.();
      pending.set(token, entry);
      const started = now();
      let launchTimer: ReturnType<typeof setTimeout> | undefined;
      try {
        // The browser launch counts towards the wait: a slow one mustn't push
        // the call past its deadline. The page stays open either way.
        await Promise.race([
          deps.openBrowser(`http://${host}:${await port()}/send/${token}`),
          new Promise<void>((resolve) => {
            launchTimer = setTimeout(resolve, waitMs);
          }),
        ]);
      } catch {
        clearTimeout(entry.expiry);
        pending.delete(token);
        void stopIfIdle();
        throw new ValidationError(
          "Couldn't open your web browser to confirm the email. Check that a default web browser " +
            "is set, then try again. Nothing was sent.",
        );
      } finally {
        clearTimeout(launchTimer);
      }
      log.info("Opened a send confirmation in the browser", {
        recipients: email.to.length + email.cc.length,
      });
      let timer: ReturnType<typeof setTimeout> | undefined;
      const waited = new Promise<ConfirmOutcome>((resolve) => {
        timer = setTimeout(
          () => resolve({ status: "pending", expiresAt }),
          Math.max(0, waitMs - (now() - started)),
        );
      });
      try {
        return await Promise.race([decided, waited]);
      } finally {
        clearTimeout(timer);
      }
    },

    async close() {
      for (const [token, entry] of pending) {
        clearTimeout(entry.expiry);
        pending.delete(token);
      }
      await stopIfIdle();
    },
  };
}

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks: Buffer[] = [];
    req.on("data", (c: Buffer) => {
      size += c.length;
      if (size > MAX_FORM_BYTES) {
        reject(new Error("form too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function page(title: string, body: string): string {
  return (
    `<!doctype html><html lang="en"><head><meta charset="utf-8">` +
    `<meta name="viewport" content="width=device-width, initial-scale=1">` +
    `<title>${escapeHtml(title)}</title><style>${STYLE}</style></head>` +
    `<body><main><h1>${escapeHtml(title)}</h1>${body}</main></body></html>`
  );
}

const STYLE =
  "body{font-family:system-ui,sans-serif;margin:0;background:#f5f5f4;color:#1c1917}" +
  "main{max-width:44rem;margin:2rem auto;padding:1.5rem;background:#fff;border-radius:8px;" +
  "box-shadow:0 1px 3px #0002}h1{font-size:1.4rem;margin-top:0}" +
  ".warn{background:#fef3c7;border:1px solid #f59e0b;padding:.75rem;border-radius:6px}" +
  "dl{display:grid;grid-template-columns:max-content 1fr;gap:.4rem 1rem}dt{font-weight:600}" +
  "dd{margin:0;word-break:break-word}ul{margin:0;padding-left:1.1rem}" +
  ".new{color:#b91c1c;font-weight:600}.unknown{color:#57534e}" +
  "pre{white-space:pre-wrap;word-break:break-word;background:#fafaf9;border:1px solid #e7e5e4;" +
  "padding:1rem;border-radius:6px;font:inherit}" +
  ".actions{display:flex;gap:1rem;margin-top:1.5rem}button{font:inherit;padding:.6rem 1.4rem;" +
  "border-radius:6px;border:1px solid #a8a29e;cursor:pointer;background:#fff}" +
  "button.send{background:#1d4ed8;color:#fff;border-color:#1d4ed8}" +
  "@media (prefers-color-scheme:dark){body{background:#1c1917;color:#f5f5f4}" +
  "main{background:#292524}pre{background:#1c1917;border-color:#44403c}" +
  ".warn{background:#422006;border-color:#b45309}button{background:#292524;color:#f5f5f4}" +
  ".new{color:#fca5a5}.unknown{color:#a8a29e}}";

function recipientList(list: Recipient[]): string {
  return (
    "<ul>" +
    list
      .map((r) => {
        const who = r.name
          ? `${escapeHtml(r.name)} &lt;${escapeHtml(r.address)}&gt;`
          : escapeHtml(r.address);
        const flag =
          r.known === false
            ? ' <span class="new">has never emailed you</span>'
            : r.known === null
              ? ' <span class="unknown">(couldn\'t check past mail)</span>'
              : "";
        return `<li>${who}${flag}</li>`;
      })
      .join("") +
    "</ul>"
  );
}

function confirmPage(email: OutgoingEmail, expiresAt: number): string {
  const anyNew = [...email.to, ...email.cc].some((r) => r.known === false);
  const reply = email.inReplyTo
    ? `<dt>Reply to</dt><dd>${escapeHtml(email.inReplyTo.subject ?? "(no subject)")}` +
      (email.inReplyTo.from ? ` from ${escapeHtml(email.inReplyTo.from)}` : "") +
      "</dd>"
    : "";
  const expires = new Date(expiresAt).toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
  });
  return page(
    "Send this email?",
    `<p class="warn">Claude asked to send this email from your account. Send it only if you asked ` +
      `Claude to, and it says what you meant. Emails can't be unsent.` +
      (anyNew
        ? " <strong>It goes to someone who has never emailed you</strong>: check that's intended."
        : "") +
      `</p>` +
      `<dl>${reply}<dt>To</dt><dd>${recipientList(email.to)}</dd>` +
      (email.cc.length ? `<dt>Cc</dt><dd>${recipientList(email.cc)}</dd>` : "") +
      `<dt>Subject</dt><dd>${escapeHtml(email.subject || "(no subject)")}</dd></dl>` +
      `<pre>${escapeHtml(email.body)}</pre>` +
      `<form method="post" class="actions">` +
      `<button class="send" name="decision" value="send">Send</button>` +
      `<button name="decision" value="cancel">Don't send</button></form>` +
      `<p class="unknown">This page stops working at ${escapeHtml(expires)}.</p>`,
  );
}
