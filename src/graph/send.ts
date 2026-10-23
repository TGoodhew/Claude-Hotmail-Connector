/**
 * Sending mail (#57, Mail.Send): new messages, and replies to the sender of a
 * message. Only ever called after the user has confirmed the exact email on
 * the connector's own page (see src/send/confirm.ts); nothing here decides
 * to send. No forwarding, reply-all, attachments or BCC; plain text only.
 */

import { GraphError, TimeoutError, ValidationError } from "../util/errors.js";
import type { GraphClient, ODataPage } from "./client.js";

/** Most recipients (To and Cc together) one email may have. */
export const MAX_RECIPIENTS = 10;
/** Longest body sent, in characters. */
export const MAX_SEND_BODY_CHARS = 20_000;
/** Longest subject. */
export const MAX_SUBJECT_CHARS = 255;

/** A plain address: no display name, comments, spaces or lists. */
const ADDRESS = /^[^\s@<>()",;:[\]\\]+@[^\s@<>()",;:[\]\\]+\.[^\s@<>()",;:[\]\\]+$/;

export interface NewEmailInput {
  to: string[];
  cc?: string[];
  subject: string;
  body: string;
}

export interface CheckedEmail {
  to: string[];
  cc: string[];
  subject: string;
  body: string;
}

function checkAddresses(list: string[], label: string): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const raw of list) {
    const address = raw.trim();
    if (!ADDRESS.test(address)) {
      throw new ValidationError(
        `${label} must be plain email addresses like name@example.com; one isn't. Nothing was sent.`,
      );
    }
    if (!seen.has(address.toLowerCase())) {
      seen.add(address.toLowerCase());
      out.push(address);
    }
  }
  return out;
}

/** The body and subject limits, shared by new emails and replies. */
function checkText(subject: string, body: string): void {
  if (subject.length > MAX_SUBJECT_CHARS) {
    throw new ValidationError(
      `The subject can be at most ${MAX_SUBJECT_CHARS} characters. Nothing was sent.`,
    );
  }
  if (!body.trim()) throw new ValidationError("The email has no text. Nothing was sent.");
  if (body.length > MAX_SEND_BODY_CHARS) {
    throw new ValidationError(
      `The email can be at most ${MAX_SEND_BODY_CHARS.toLocaleString("en-US")} characters. ` +
        "Nothing was sent.",
    );
  }
}

/** Check and tidy a new email: plain addresses, no duplicates, the limits. */
export function checkNewEmail(input: NewEmailInput): CheckedEmail {
  const to = checkAddresses(input.to, "To");
  if (to.length === 0) throw new ValidationError("Say who the email is to. Nothing was sent.");
  const toSet = new Set(to.map((a) => a.toLowerCase()));
  // Someone on both lines is sent to once, on To.
  const cc = checkAddresses(input.cc ?? [], "Cc").filter((a) => !toSet.has(a.toLowerCase()));
  if (to.length + cc.length > MAX_RECIPIENTS) {
    throw new ValidationError(
      `An email can go to at most ${MAX_RECIPIENTS} people (To and Cc together). Nothing was sent.`,
    );
  }
  const subject = input.subject.trim();
  checkText(subject, input.body);
  return { to, cc, subject, body: input.body };
}

/**
 * Whether `address` has ever sent the user mail; null if that couldn't be
 * checked. A first-time recipient is how an injected "send this to me" shows
 * itself, so the confirmation page flags it.
 */
export async function hasEmailedYou(graph: GraphClient, address: string): Promise<boolean | null> {
  try {
    const page = await graph.get<ODataPage<{ id: string }>>("/me/messages", {
      query: {
        // OData string literal: a quote is doubled.
        $filter: `from/emailAddress/address eq '${address.replace(/'/g, "''")}'`,
        $top: 1,
        $select: "id",
      },
    });
    return (page.value ?? []).length > 0;
  } catch (e) {
    // A query Graph won't answer means "couldn't check"; anything else (a
    // sign-in needed, a sign-out, the call cancelled) still ends the call.
    if (e instanceof GraphError && e.reason !== "cancelled") return null;
    throw e;
  }
}

/** {@link hasEmailedYou} for several addresses, three lookups at a time. */
export async function knownSenders(
  graph: GraphClient,
  addresses: string[],
): Promise<Map<string, boolean | null>> {
  const out = new Map<string, boolean | null>();
  const queue = [...addresses];
  const worker = async () => {
    for (let a = queue.shift(); a !== undefined; a = queue.shift()) {
      out.set(a.toLowerCase(), await hasEmailedYou(graph, a));
    }
  };
  await Promise.all([worker(), worker(), worker()]);
  return out;
}

/**
 * A send may have reached Microsoft even when no answer came back, or when the
 * answer was a server error (it can fail after the work is done): never say
 * "nothing was sent" then. Sends aren't retried, for the same reason.
 */
function uncertain(e: unknown): never {
  if (
    e instanceof TimeoutError ||
    (e instanceof GraphError && (e.status === undefined || [500, 502, 503, 504].includes(e.status)))
  ) {
    throw new ValidationError(
      "Microsoft didn't confirm the email was sent, but it may have been. Check Sent Items " +
        "before sending it again.",
    );
  }
  throw e;
}

const recipients = (list: { address: string; name?: string }[]) =>
  list.map((r) => ({ emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) } }));

/** Send a new plain-text email, keeping a copy in Sent Items. */
export async function sendNewEmail(graph: GraphClient, email: CheckedEmail): Promise<void> {
  try {
    await graph.post("/me/sendMail", {
      message: {
        subject: email.subject,
        body: { contentType: "Text", content: email.body },
        toRecipients: recipients(email.to.map((address) => ({ address }))),
        ccRecipients: recipients(email.cc.map((address) => ({ address }))),
      },
      saveToSentItems: true,
    });
  } catch (e) {
    uncertain(e);
  }
}

export interface PreparedReply {
  messageId: string;
  /** Who the reply goes to: the message's Reply-To, else its sender. */
  to: { address: string; name?: string }[];
  /** The subject as Outlook will send it. */
  subject: string;
  body: string;
  original: { subject: string | null; from: string | null };
}

interface GraphAddress {
  emailAddress?: { address?: string; name?: string };
}

/**
 * Work out a reply before showing it: its recipients are the message's
 * Reply-To if it has one, else its sender, and the page shows exactly those.
 */
export async function prepareReply(
  graph: GraphClient,
  messageId: string,
  body: string,
): Promise<PreparedReply> {
  const m = await graph.get<{ subject?: string; from?: GraphAddress; replyTo?: GraphAddress[] }>(
    `/me/messages/${encodeURIComponent(messageId)}`,
    { query: { $select: "subject,from,replyTo" } },
  );
  const pick = (a?: GraphAddress) =>
    a?.emailAddress?.address
      ? {
          address: a.emailAddress.address,
          ...(a.emailAddress.name ? { name: a.emailAddress.name } : {}),
        }
      : undefined;
  const replyTo = (m.replyTo ?? []).map(pick).filter((a) => a !== undefined);
  const sender = pick(m.from);
  const to = replyTo.length > 0 ? replyTo : sender ? [sender] : [];
  if (to.length === 0) {
    throw new ValidationError("That message has no sender to reply to. Nothing was sent.");
  }
  if (to.length > MAX_RECIPIENTS) {
    throw new ValidationError(
      `That message asks for replies to more than ${MAX_RECIPIENTS} people, so it can't be ` +
        "answered here. Nothing was sent.",
    );
  }
  const original = m.subject?.trim() ?? "";
  const subject = /^re:/i.test(original) ? original : `RE: ${original}`;
  checkText(subject, body);
  return {
    messageId,
    to,
    subject,
    body,
    original: { subject: m.subject ?? null, from: sender?.address ?? null },
  };
}

/**
 * Send a prepared reply, in the message's thread. The recipients and body are
 * given explicitly, so what's sent is what the page showed: the reply doesn't
 * quote the original message, and goes only to the people listed.
 */
export async function sendReply(graph: GraphClient, reply: PreparedReply): Promise<void> {
  try {
    await graph.post(`/me/messages/${encodeURIComponent(reply.messageId)}/reply`, {
      message: {
        toRecipients: recipients(reply.to),
        body: { contentType: "Text", content: reply.body },
      },
    });
  } catch (e) {
    uncertain(e);
  }
}
