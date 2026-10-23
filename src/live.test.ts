/**
 * Opt-in live test against the real Microsoft Graph (spec v2 §13, #48).
 *
 * Skipped unless HOTMAIL_LIVE_TESTS=1, so CI never runs it. To run it from a
 * source checkout with your own app registration (see the README):
 *
 *   node dist/index.js login      # once: signs in and caches tokens
 *   npm run test:live
 *
 * It reads a little mail (read-only), then creates throwaway events far in
 * the future in a DST zone: one checks its local time survived the round trip
 * and repeats the create with the same transactionId to see how Graph answers;
 * another checks that a write carrying a stale version tag (If-Match) is
 * refused. Test events are cancelled (moved to Deleted Items) whatever
 * happens. It never opens a browser: with no cached sign-in it fails and asks
 * you to run `login`.
 *
 * With ALLOW_MAIL_DRAFTS=true (and a sign-in that granted Mail.ReadWrite), it
 * also saves two drafts (#184): a new one to yourself, and a reply to your
 * newest Inbox message. It checks they're unsent drafts in Drafts, that the
 * reply keeps its thread with your text above the quote, and prints the ways
 * to open each in Outlook on the web. They're deleted (moved to Deleted Items)
 * afterwards, unless HOTMAIL_LIVE_KEEP_DRAFTS=1 keeps them so you can try the
 * links by hand; then discard them in Outlook. Nothing is ever sent.
 */

import { randomUUID } from "node:crypto";
import { describe, expect, it } from "vitest";
import { createMicrosoftAuth } from "./auth/microsoft.js";
import { defaultKeyProtector } from "./auth/protect.js";
import { createFileTokenStore } from "./auth/tokens.js";
import { getConfig } from "./config.js";
import { cancelEvent, createEvent, listEvents, updateEvent } from "./graph/calendar.js";
import { createGraphClient, scopeToOperation, type GraphClient } from "./graph/client.js";
import {
  createDraft,
  createReplyDraft,
  draftUrlCandidates,
  type SavedDraft,
} from "./graph/drafts.js";
import { getMessage, listMessages } from "./graph/mail.js";
import { whoami } from "./graph/user.js";
import { GraphError } from "./util/errors.js";

const LIVE = process.env.HOTMAIL_LIVE_TESTS === "1";
/** Leave the test drafts in Drafts, to try their links by hand (#184). */
const KEEP_DRAFTS = process.env.HOTMAIL_LIVE_KEEP_DRAFTS === "1";
/** Built from a char code, so no escape sequence can be mangled in editing. */
const NEWLINE = String.fromCharCode(10);
/** Two lines, to check a reply draft keeps the line break. */
const TEST_DRAFT_BODY = ["Hotmail Connector live test: don't send.", "Second line."].join(NEWLINE);
/** Deadlines for the write test's main steps and for removing what it created. */
const MAIN_MS = 70_000;
const CLEANUP_MS = 35_000;
/** Part of the cleanup time for finding events by subject; the rest is for deleting. */
const LOOKUP_MS = 10_000;

/**
 * Only the test event's own 09:30-10:15 slot (with a margin), so unrelated
 * events that day can't push it past the listing limit.
 */
const TEST_WINDOW = {
  startDateTime: "2031-01-15T09:00:00",
  endDateTime: "2031-01-15T10:45:00",
  limit: 100,
};

/** A listing that hit its limit may have left the test event out: say so. */
function warnIfCapped(truncated: boolean, subject: string): void {
  if (truncated) {
    console.warn(`[live] The calendar listing hit its limit; check for "${subject}" by hand.`);
  }
}

/** A subject unique to this run (a timestamp plus a random suffix), and its stamp. */
function uniqueSubject(): { stamp: string; subject: string } {
  const time = new Date()
    .toISOString()
    .replace(/[^0-9]/g, "")
    .slice(0, 14);
  const stamp = `${time}-${randomUUID().slice(0, 8)}`;
  return { stamp, subject: `Hotmail Connector live test ${stamp}` };
}

/**
 * Cancel the test events: the ids known, plus any event in the test window
 * whose subject starts with `prefix`. A create can succeed at Graph yet fail
 * here (a lost response, the deadline), and a test may have renamed its event,
 * hence the lookup by prefix (each event is confirmed with its own title).
 */
async function removeTestEvents(
  base: GraphClient,
  known: Map<string, string>,
  prefix: string,
  timeZone: string,
): Promise<void> {
  const cleanup = scopeToOperation(base, { deadline: Date.now() + CLEANUP_MS });
  const events = new Map(known);
  try {
    // A shorter deadline of its own, so a stalled lookup leaves time to
    // delete the events whose ids are already known.
    const lookup = scopeToOperation(base, { deadline: Date.now() + LOOKUP_MS });
    const day = await listEvents(lookup, { ...TEST_WINDOW, timeZone });
    warnIfCapped(day.truncated, prefix);
    for (const e of day.events) {
      if (e.subject?.startsWith(prefix)) events.set(e.id, e.subject);
    }
  } catch (e) {
    console.warn(`[live] Couldn't look for test events; check for "${prefix}".`, e);
  }
  for (const [id, subject] of events) {
    await cancelEvent(cleanup, { id, confirmSubject: subject, timeZone }).catch((e: unknown) =>
      console.warn(`[live] Couldn't remove test event ${id}; delete "${subject}" by hand.`, e),
    );
  }
}

/** What a write that should be refused actually got: "412" if it was refused that way. */
async function outcomeOf(write: () => Promise<unknown>): Promise<string> {
  try {
    await write();
    return "applied";
  } catch (e) {
    if (e instanceof GraphError && e.status !== undefined) return String(e.status);
    return `failed: ${e instanceof Error ? e.message : String(e)}`;
  }
}

/**
 * Print the ways to open a draft in Outlook on the web (#184), including the
 * draft's EWS id form, which Outlook's own URLs may expect. Try each by hand
 * with HOTMAIL_LIVE_KEEP_DRAFTS=1 to see which opens the draft for editing.
 */
async function printLinks(g: GraphClient, label: string, draft: SavedDraft): Promise<void> {
  const links: Record<string, string | undefined> = { ...draftUrlCandidates(draft) };
  try {
    const t = await g.post<{ value?: { targetId?: string }[] }>("/me/translateExchangeIds", {
      inputIds: [draft.id],
      sourceIdType: "restId",
      targetIdType: "ewsId",
    });
    const ews = t.value?.[0]?.targetId;
    const ewsDraft = ews ? draftUrlCandidates({ id: ews, webLink: draft.webLink }) : {};
    links.composeDeeplinkEwsId = ewsDraft.composeDeeplink;
    links.draftsFolderEwsId = ewsDraft.draftsFolder;
  } catch (e) {
    console.log(
      `[live] translateExchangeIds failed: ${e instanceof Error ? e.message : String(e)}`,
    );
  }
  console.log(`[live] Links for the ${label}:`);
  for (const [name, url] of Object.entries(links)) if (url) console.log(`  ${name}: ${url}`);
}

describe.runIf(LIVE)("live Microsoft Graph (opt-in, #48)", () => {
  const config = LIVE ? getConfig() : undefined;
  const graph = () => {
    const store = createFileTokenStore(config!.tokenCacheFile, {
      protector: defaultKeyProtector(),
    });
    const auth = createMicrosoftAuth({ config: config!, store });
    return createGraphClient({ auth, baseUrl: config!.graphBaseUrl }); // non-interactive
  };

  it("reads the signed-in account and some mail", { timeout: 60_000 }, async () => {
    const g = graph();
    const me = await whoami(g);
    expect(me.mail ?? me.userPrincipalName).toBeTruthy();

    const { messages } = await listMessages(g, { limit: 3 });
    if (messages.length > 0) {
      const msg = await getMessage(g, { id: messages[0]!.id, format: "text" });
      expect(msg.id).toBe(messages[0]!.id);
      expect(msg.body.contentType).toBe("text");
    }
  });

  it(
    "creates an event at the intended local time, de-duplicates by transactionId, and cleans up",
    { timeout: MAIN_MS + CLEANUP_MS + 15_000 },
    async () => {
      // The Graph client retries (with Retry-After waits) until a deadline, and a
      // Vitest timeout wouldn't cancel it or run `finally`. So the main steps get
      // a deadline, and cleanup gets its own fresh one, both inside the timeout.
      const base = graph();
      const g = scopeToOperation(base, { deadline: Date.now() + MAIN_MS });
      const timeZone = "Australia/Sydney"; // a DST zone (#64)
      // Unique per run, so two runs on one account can't share a subject or
      // transactionId; the repeat below reuses it.
      const { stamp, subject } = uniqueSubject();
      const input = {
        subject,
        start: { dateTime: "2031-01-15T09:30:00", timeZone },
        end: { dateTime: "2031-01-15T10:15:00", timeZone },
        transactionId: `live-${stamp}`,
      };
      const created = new Map<string, string>();
      try {
        const first = await createEvent(g, input);
        created.set(first.id, subject);

        // The wall-clock time is preserved, not shifted to UTC. The create
        // response's zone isn't requested, so this is checked through a listing
        // rendered in Sydney time.
        const { events: listed, truncated } = await listEvents(g, { ...TEST_WINDOW, timeZone });
        warnIfCapped(truncated, subject);
        const found = listed.find((e) => e.id === first.id);
        expect(found?.start?.dateTime?.startsWith("2031-01-15T09:30")).toBe(true);
        expect(found?.end?.dateTime?.startsWith("2031-01-15T10:15")).toBe(true);

        // How Graph answers a repeat of the same transactionId (#48, D7).
        let repeat: string;
        try {
          const second = await createEvent(g, input);
          if (second.id !== first.id) created.set(second.id, subject);
          repeat = second.id === first.id ? "returned the same event" : "created a second event";
        } catch (e) {
          repeat = `failed: ${e instanceof Error ? e.message : String(e)}`;
        }
        console.log(`[live] Repeating a create with the same transactionId ${repeat}.`);
        expect(repeat).toBe("returned the same event");
      } finally {
        await removeTestEvents(base, created, subject, timeZone);
      }
    },
  );

  it(
    "refuses a PATCH with a stale If-Match (412); notes that a DELETE ignores it (spec §7.6, §7.8)",
    { timeout: MAIN_MS + CLEANUP_MS + 15_000 },
    async () => {
      const base = graph();
      const g = scopeToOperation(base, { deadline: Date.now() + MAIN_MS });
      const timeZone = "Australia/Sydney";
      const { stamp, subject } = uniqueSubject();
      const created = new Map<string, string>();
      try {
        const event = await createEvent(g, {
          subject,
          start: { dateTime: "2031-01-15T09:30:00", timeZone },
          end: { dateTime: "2031-01-15T10:15:00", timeZone },
          transactionId: `live-etag-${stamp}`,
        });
        created.set(event.id, subject);
        const path = `/me/events/${encodeURIComponent(event.id)}`;
        const version = async () =>
          (await g.get<{ "@odata.etag"?: string }>(path, { query: { $select: "id" } }))[
            "@odata.etag"
          ];
        const stale = await version();
        expect(stale, "Graph returned no @odata.etag for the event").toBeTruthy();

        // A normal update_event: it reads the event, then writes the checked version.
        const renamed = `${subject} (renamed)`;
        const { changes } = await updateEvent(g, { id: event.id, subject: renamed });
        created.set(event.id, renamed);
        expect(changes).toEqual([{ field: "subject", before: subject, after: renamed }]);
        expect(await version()).not.toBe(stale); // the write gave it a new version

        // A PATCH carrying the version from before that update is refused…
        const patch = await outcomeOf(() =>
          g.patch(
            path,
            { subject: `${subject} (stale write)` },
            { headers: { "if-match": stale! } },
          ),
        );
        console.log(`[live] A PATCH with a stale If-Match got: ${patch}.`);
        expect(patch).toBe("412");
        const after = await g.get<{ subject?: string }>(path, { query: { $select: "subject" } });
        expect(after.subject).toBe(renamed); // …and had no effect.

        // A DELETE isn't: Graph ignores If-Match on event deletes (found by this
        // test, 2026-10-22; spec §7.8). Recorded here so a change in Graph shows
        // up. The event goes to Deleted Items, so there's nothing left to clean.
        const remove = await outcomeOf(() => g.delete(path, { headers: { "if-match": stale! } }));
        console.log(`[live] A DELETE with a stale If-Match got: ${remove}.`);
        expect(remove).toBe("applied");
        created.delete(event.id);
        expect(await outcomeOf(() => g.get(path, { query: { $select: "id" } }))).toBe("404");
      } finally {
        await removeTestEvents(base, created, subject, timeZone);
      }
    },
  );
  it.runIf(LIVE && getConfig().allowMailDrafts)(
    "saves a new draft and a reply draft without sending, and prints their Outlook links (#184)",
    { timeout: MAIN_MS + CLEANUP_MS + 15_000 },
    async () => {
      const base = graph();
      const g = scopeToOperation(base, { deadline: Date.now() + MAIN_MS });
      const { subject } = uniqueSubject();
      const saved: SavedDraft[] = [];
      try {
        const me = await whoami(g);
        const self = me.mail ?? me.userPrincipalName;
        expect(self).toBeTruthy();
        const drafts = await g.get<{ id: string }>("/me/mailFolders/drafts", {
          query: { $select: "id" },
        });
        type Saved = {
          isDraft?: boolean;
          parentFolderId?: string;
          conversationId?: string;
          body?: { content?: string };
        };
        const read = (id: string) =>
          g.get<Saved>(`/me/messages/${encodeURIComponent(id)}`, {
            query: { $select: "isDraft,parentFolderId,conversationId,body" },
            headers: { prefer: 'outlook.body-content-type="text"' },
          });

        // A new draft, to yourself.
        const fresh = await createDraft(g, {
          to: [self!],
          cc: [],
          subject,
          body: TEST_DRAFT_BODY,
        });
        saved.push(fresh);
        const freshRead = await read(fresh.id);
        expect(freshRead.isDraft).toBe(true);
        expect(freshRead.parentFolderId).toBe(drafts.id);
        await printLinks(g, "new draft", fresh);

        // A reply draft to the newest Inbox message.
        const inbox = await g.get<{ value?: { id: string; conversationId?: string }[] }>(
          "/me/mailFolders/inbox/messages",
          { query: { $top: 1, $orderby: "receivedDateTime desc", $select: "id,conversationId" } },
        );
        const original = inbox.value?.[0];
        if (!original) {
          console.warn("[live] The Inbox is empty, so no reply draft was tried.");
          return;
        }
        const reply = await createReplyDraft(g, original.id, TEST_DRAFT_BODY);
        saved.push(reply);
        const replyRead = await read(reply.id);
        expect(replyRead.isDraft).toBe(true);
        expect(replyRead.parentFolderId).toBe(drafts.id);
        expect(replyRead.conversationId).toBe(original.conversationId);
        const text = replyRead.body?.content ?? "";
        const first = text.indexOf("Hotmail Connector live test");
        const second = text.indexOf("Second line.");
        // Lines, trimmed (Graph's text may use CRLF), to find the quote's header
        // and whether the comment's line break survived.
        const lines = text.split(NEWLINE).map((l) => l.trim());
        const quote = lines.findIndex((l) => /^(From|De|Von):/.test(l));
        const kept = lines.some(
          (l, i) => l.endsWith("don't send.") && lines[i + 1] === "Second line.",
        );
        console.log(
          `[live] Reply draft text: comment at ${first}, second line at ${second}, quote header ` +
            `on line ${quote}; line break kept: ${kept}.`,
        );
        // Links first, so they're printed even if a check below fails.
        await printLinks(g, "reply draft", reply);
        expect(first).toBeGreaterThanOrEqual(0);
        expect(second).toBeGreaterThan(first);
        // Graph must keep the comment's line break, else replies need HTML (#184).
        expect(kept, "the reply draft lost the comment's line break").toBe(true);
        if (quote >= 0) {
          expect(quote).toBeGreaterThan(lines.findIndex((l) => l === "Second line."));
        }
      } finally {
        if (KEEP_DRAFTS) {
          console.log(
            `[live] Kept ${saved.length} test draft(s) in Drafts (HOTMAIL_LIVE_KEEP_DRAFTS=1). ` +
              "Try the links above, then discard them in Outlook. Don't send them.",
          );
        } else {
          const cleanup = scopeToOperation(base, { deadline: Date.now() + CLEANUP_MS });
          for (const d of saved) {
            await cleanup
              .delete(`/me/messages/${encodeURIComponent(d.id)}`)
              .catch((e: unknown) =>
                console.warn(
                  `[live] Couldn't remove test draft "${d.subject}"; discard it by hand.`,
                  e,
                ),
              );
          }
        }
      }
    },
  );
});
