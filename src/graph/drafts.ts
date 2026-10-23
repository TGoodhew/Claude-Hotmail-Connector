/**
 * Saving drafts (#184, Mail.ReadWrite): a new message, or a reply to a
 * message, saved to Drafts for the user to review, edit and send from Outlook.
 * Nothing here sends: the only requests are `POST /me/messages` and
 * `POST /me/messages/{id}/createReply`, both of which only create a draft.
 * No existing message is changed, moved or deleted (spec v2 §6, §7.10).
 */

import { GraphError, TimeoutError, ValidationError } from "../util/errors.js";
import type { CheckedEmail } from "./send.js";
import type { GraphClient } from "./client.js";

export interface SavedDraft {
  id: string;
  subject: string;
  to: string[];
  cc: string[];
  /** Outlook on the web's link to the draft, as Graph gives it. */
  webLink: string | null;
}

interface GraphAddress {
  emailAddress?: { address?: string };
}

interface GraphDraft {
  id?: string;
  subject?: string | null;
  toRecipients?: GraphAddress[];
  ccRecipients?: GraphAddress[];
  webLink?: string | null;
}

/**
 * A draft may have been saved even when no answer came back, or when the
 * answer was a server error: never say "nothing was saved" then. The Graph
 * client doesn't repeat these POSTs (they carry no transactionId), so there's
 * no duplicate draft from a retry.
 */
function uncertain(e: unknown): never {
  if (
    e instanceof TimeoutError ||
    (e instanceof GraphError && (e.status === undefined || [500, 502, 503, 504].includes(e.status)))
  ) {
    throw new ValidationError(
      "Microsoft didn't confirm the draft was saved, but it may have been. Check Drafts in " +
        "Outlook before trying again.",
    );
  }
  throw e;
}

const addresses = (list?: GraphAddress[]) =>
  (list ?? []).map((r) => r.emailAddress?.address).filter((a): a is string => !!a);

function saved(d: GraphDraft): SavedDraft {
  // The draft exists: say so, so it isn't saved again by a retry.
  if (!d.id) {
    throw new ValidationError(
      "Microsoft saved the draft but didn't say which one it is. Check Drafts in Outlook " +
        "before trying again.",
    );
  }
  return {
    id: d.id,
    subject: d.subject ?? "",
    to: addresses(d.toRecipients),
    cc: addresses(d.ccRecipients),
    webLink: d.webLink ?? null,
  };
}

const recipients = (list: string[]) => list.map((address) => ({ emailAddress: { address } }));

/** Save a new plain-text email to Drafts, without sending it. */
export async function createDraft(graph: GraphClient, email: CheckedEmail): Promise<SavedDraft> {
  let d: GraphDraft;
  try {
    d = await graph.post<GraphDraft>("/me/messages", {
      subject: email.subject,
      body: { contentType: "Text", content: email.body },
      toRecipients: recipients(email.to),
      ccRecipients: recipients(email.cc),
    });
  } catch (e) {
    uncertain(e);
  }
  return saved(d);
}

/**
 * Save a reply to a message as a draft, without sending it. Graph puts the
 * `comment` above the quoted original, and keeps the thread and the message's
 * Reply-To, so the draft is just what Outlook's own Reply would start.
 */
export async function createReplyDraft(
  graph: GraphClient,
  messageId: string,
  body: string,
): Promise<SavedDraft> {
  let d: GraphDraft;
  try {
    d = await graph.post<GraphDraft>(`/me/messages/${encodeURIComponent(messageId)}/createReply`, {
      comment: body,
    });
  } catch (e) {
    uncertain(e);
  }
  return saved(d);
}

/** Outlook on the web's hosts: a link elsewhere isn't opened. */
const OUTLOOK_HOSTS = new Set(["outlook.live.com", "outlook.office.com", "outlook.office365.com"]);

/** `webLink`'s origin if it's Outlook on the web over HTTPS, else undefined. */
function outlookOrigin(webLink: string | null): string | undefined {
  if (!webLink) return undefined;
  try {
    const u = new URL(webLink);
    return u.protocol === "https:" && OUTLOOK_HOSTS.has(u.hostname) ? u.origin : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Ways to open a draft in Outlook on the web. The live check (#184) prints
 * each, to find the one that opens the draft for editing.
 */
export function draftUrlCandidates(draft: Pick<SavedDraft, "id" | "webLink">): {
  webLink?: string;
  composeDeeplink?: string;
  draftsFolder?: string;
} {
  const origin = outlookOrigin(draft.webLink);
  if (!origin) return {};
  const id = encodeURIComponent(draft.id);
  return {
    webLink: draft.webLink!,
    composeDeeplink: `${origin}/mail/0/deeplink/compose/${id}`,
    draftsFolder: `${origin}/mail/0/drafts/id/${id}`,
  };
}

/**
 * The link `create_draft` opens: undefined if Graph gave none on Outlook's own
 * hosts. Graph's `webLink` until the live check (#184) shows which form opens
 * the draft for editing.
 */
export function draftUrl(draft: Pick<SavedDraft, "id" | "webLink">): string | undefined {
  return draftUrlCandidates(draft).webLink;
}
