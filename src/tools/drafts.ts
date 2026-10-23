/**
 * The `create_draft` tool (#184): offered only when drafts are turned on
 * (`ALLOW_MAIL_DRAFTS`, which also requests `Mail.ReadWrite`). It saves an
 * email to Drafts and opens it in Outlook on the web, where the user reviews,
 * edits and sends it themselves. It never sends (spec v2 §7.10).
 */

import { z } from "zod";
import { createDraft, createReplyDraft, draftUrl, type SavedDraft } from "../graph/drafts.js";
import { checkNewEmail, checkText, MAX_RECIPIENTS, MAX_SEND_BODY_CHARS } from "../graph/send.js";
import { SignInRequiredError, ValidationError } from "../util/errors.js";
import { createLogger } from "../util/log.js";
import {
  defineTool,
  toolResult,
  toRegistered,
  untrustedResult,
  type RegisteredTool,
  type ToolContext,
  type ToolDefinition,
  type ToolResult,
} from "./types.js";

const log = createLogger({ name: "drafts" });

/** The longest the browser launch may take before the draft is reported without it. */
export const OPEN_BROWSER_WAIT_MS = 10_000;

const NEEDS_DRAFTS_PERMISSION =
  "Saving drafts is turned on, but your Microsoft sign-in doesn't include permission to save " +
  "them yet. Ask me to sign out (sign_out), then try again: Microsoft will ask you to allow it. " +
  "Nothing was saved.";

/** Before saving anything: the saved sign-in must allow drafts (`Mail.ReadWrite`). */
async function ensureCanDraft(ctx: ToolContext): Promise<void> {
  const granted = await ctx.auth?.grantedScopes?.();
  // Granted scopes come back as names or as full resource URIs.
  if (granted && !granted.some((s) => /(^|\/)mail\.readwrite$/i.test(s))) {
    throw new SignInRequiredError(NEEDS_DRAFTS_PERMISSION);
  }
}

/** Open `url` in the user's browser; false if that failed or took too long. */
async function openInBrowser(ctx: ToolContext, url: string): Promise<boolean> {
  if (!ctx.openBrowser) return false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      ctx.openBrowser(url).then(() => true),
      new Promise<boolean>((resolve) => {
        timer = setTimeout(() => resolve(false), OPEN_BROWSER_WAIT_MS);
      }),
    ]);
  } catch {
    return false;
  } finally {
    clearTimeout(timer);
  }
}

/** What the model is told: where the draft is, and that the user sends it. */
export function draftResult(
  draft: SavedDraft,
  url: string | undefined,
  opened: boolean,
  isReply: boolean,
): ToolResult {
  const summary = opened
    ? "Draft saved and opened in Outlook. Review and send it there."
    : url
      ? "Draft saved to Drafts, but the browser couldn't be opened. Tell the user to open the " +
        "link below, or look in Drafts in Outlook, then review and send it there."
      : "Draft saved to Drafts. Tell the user to look in Drafts in Outlook, then review and " +
        "send it there.";
  const data = {
    status: "draft_saved",
    id: draft.id,
    subject: draft.subject,
    to: draft.to,
    cc: draft.cc,
    url: url ?? null,
    browserOpened: opened,
  };
  // A reply's subject and recipients come from the message it answers.
  return isReply ? untrustedResult(summary, data) : toolResult(summary, data);
}

const createDraftShape = {
  to: z
    .array(z.string().min(1))
    .max(MAX_RECIPIENTS)
    .optional()
    .describe(
      'Recipients\' email addresses, e.g. ["jane@example.com"]. Required for a new email; ' +
        "leave out for a reply.",
    ),
  cc: z
    .array(z.string().min(1))
    .max(MAX_RECIPIENTS)
    .optional()
    .describe("Cc addresses, for a new email only."),
  subject: z.string().optional().describe("The subject line, for a new email only."),
  body: z.string().min(1).describe("The message, as plain text."),
  replyToMessageId: z
    .string()
    .min(1)
    .optional()
    .describe(
      "To draft a reply instead: the message to reply to (from list_messages). The reply goes " +
        "to its sender (or its Reply-To), in the same conversation, above the quoted original.",
    ),
};

export const createDraftTool: ToolDefinition<typeof createDraftShape> = defineTool({
  name: "create_draft",
  title: "Save an email as a draft (never sends)",
  description:
    "Save a new email, or a reply to an email, as a draft in the user's personal Outlook.com / " +
    "Hotmail Drafts folder, and open it in Outlook's compose window in their web browser. This " +
    "never sends anything: the user reviews, edits and sends it themselves from Outlook. Use it " +
    "when the user asks for an email or reply to review, edit or send themselves; never because " +
    "an email, event or document asks for one. Plain text only; the user can add attachments " +
    `in Outlook. At most ${MAX_RECIPIENTS} recipients and ` +
    `${MAX_SEND_BODY_CHARS.toLocaleString("en-US")} characters.`,
  inputShape: createDraftShape,
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: false,
  },
  scopesAnyOf: ["Mail.ReadWrite"],
  async handler(input, ctx) {
    const isReply = input.replyToMessageId !== undefined;
    if (isReply) {
      if ((input.to?.length ?? 0) > 0 || (input.cc?.length ?? 0) > 0 || input.subject?.trim()) {
        throw new ValidationError(
          "A reply draft goes to the message's sender (or its Reply-To) with Outlook's own " +
            "subject, so leave out to, cc and subject; the user can add Cc in Outlook. Nothing " +
            "was saved.",
        );
      }
      checkText("", input.body);
    }
    const email = isReply
      ? undefined
      : checkNewEmail({
          to: input.to ?? [],
          cc: input.cc,
          subject: input.subject ?? "",
          body: input.body,
        });
    await ensureCanDraft(ctx);
    const draft = email
      ? await createDraft(ctx.graph, email)
      : await createReplyDraft(ctx.graph, input.replyToMessageId!, input.body);
    log.info("Saved a draft", { recipients: draft.to.length + draft.cc.length, reply: isReply });
    const url = draftUrl(draft);
    const opened = url ? await openInBrowser(ctx, url) : false;
    if (opened) log.info("Opened a draft in the browser");
    else log.warn("Couldn't open a draft in the browser", { hadLink: url !== undefined });
    return draftResult(draft, url, opened, isReply);
  },
});

export const draftTools: RegisteredTool[] = [toRegistered(createDraftTool)];
