/**
 * Opt-in live test against the real Microsoft Graph (spec v2 §13, #48).
 *
 * Skipped unless HOTMAIL_LIVE_TESTS=1, so CI never runs it. To run it from a
 * source checkout with your own app registration (see the README):
 *
 *   node dist/index.js login      # once: signs in and caches tokens
 *   npm run test:live
 *
 * It reads a little mail (read-only), then creates one throwaway event far in
 * the future in a DST zone, checks its local time survived the round trip,
 * repeats the create with the same transactionId to see how Graph answers, and
 * cancels the event (moved to Deleted Items) whatever happens. It never opens
 * a browser: with no cached sign-in it fails and asks you to run `login`.
 */

import { describe, expect, it } from "vitest";
import { createMicrosoftAuth } from "./auth/microsoft.js";
import { defaultKeyProtector } from "./auth/protect.js";
import { createFileTokenStore } from "./auth/tokens.js";
import { getConfig } from "./config.js";
import { cancelEvent, createEvent, listEvents } from "./graph/calendar.js";
import { createGraphClient, scopeToOperation } from "./graph/client.js";
import { getMessage, listMessages } from "./graph/mail.js";
import { whoami } from "./graph/user.js";

const LIVE = process.env.HOTMAIL_LIVE_TESTS === "1";
/** Deadlines for the write test's main steps and for removing what it created. */
const MAIN_MS = 70_000;
const CLEANUP_MS = 35_000;
/** Part of the cleanup time for finding events by subject; the rest is for deleting. */
const LOOKUP_MS = 10_000;

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
      const stamp = new Date()
        .toISOString()
        .replace(/[^0-9]/g, "")
        .slice(0, 14);
      const subject = `Hotmail Connector live test ${stamp}`;
      const input = {
        subject,
        start: { dateTime: "2031-01-15T09:30:00", timeZone },
        end: { dateTime: "2031-01-15T10:15:00", timeZone },
        transactionId: `live-${stamp}`,
      };
      const created: string[] = [];
      try {
        const first = await createEvent(g, input);
        created.push(first.id);

        // The wall-clock time is preserved, not shifted to UTC. The create
        // response's zone isn't requested, so this is checked through a listing
        // rendered in Sydney time.
        const listed = await listEvents(g, {
          startDateTime: "2031-01-15",
          endDateTime: "2031-01-16",
          timeZone,
          limit: 50,
        });
        const found = listed.find((e) => e.id === first.id);
        expect(found?.start?.dateTime?.startsWith("2031-01-15T09:30")).toBe(true);
        expect(found?.end?.dateTime?.startsWith("2031-01-15T10:15")).toBe(true);

        // How Graph answers a repeat of the same transactionId (#48, D7).
        let repeat: string;
        try {
          const second = await createEvent(g, input);
          if (second.id !== first.id) created.push(second.id);
          repeat = second.id === first.id ? "returned the same event" : "created a second event";
        } catch (e) {
          repeat = `failed: ${e instanceof Error ? e.message : String(e)}`;
        }
        console.log(`[live] Repeating a create with the same transactionId ${repeat}.`);
        expect(repeat).toBe("returned the same event");
      } finally {
        const cleanup = scopeToOperation(base, { deadline: Date.now() + CLEANUP_MS });
        // A create can succeed at Graph yet fail here (a lost response, the
        // deadline), so also look the event up by its unique subject.
        const ids = new Set(created);
        try {
          // A shorter deadline of its own, so a stalled lookup leaves time to
          // delete the events whose ids are already known.
          const lookup = scopeToOperation(base, { deadline: Date.now() + LOOKUP_MS });
          const day = await listEvents(lookup, {
            startDateTime: "2031-01-15",
            endDateTime: "2031-01-16",
            timeZone,
            limit: 50,
          });
          for (const e of day) if (e.subject === subject) ids.add(e.id);
        } catch (e) {
          console.warn(`[live] Couldn't look for test events; check for "${subject}".`, e);
        }
        for (const id of ids) {
          await cancelEvent(cleanup, { id, confirmSubject: subject, timeZone }).catch(
            (e: unknown) =>
              console.warn(
                `[live] Couldn't remove test event ${id}; delete "${subject}" by hand.`,
                e,
              ),
          );
        }
      }
    },
  );
});
