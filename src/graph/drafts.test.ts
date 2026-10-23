import { describe, expect, it, vi } from "vitest";
import { clientSafeMessage, GraphError, TimeoutError, ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { createDraft, createReplyDraft, draftUrl, draftUrlCandidates } from "./drafts.js";

type Post = (path: string, body: unknown) => Promise<unknown>;

const graphWith = (post: Post) =>
  ({
    get: vi.fn(),
    post: vi.fn(post),
    patch: vi.fn(),
    delete: vi.fn(),
  }) as unknown as GraphClient & {
    get: ReturnType<typeof vi.fn>;
    post: ReturnType<typeof vi.fn<Post>>;
    patch: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  };

const WEB_LINK =
  "https://outlook.live.com/owa/?ItemID=AQMk%2Fab&exvsurl=1&viewmodel=ReadMessageItem";

const draftResponse = {
  id: "AQMk/ab+c=",
  subject: "Hi",
  toRecipients: [{ emailAddress: { address: "jane@example.com" } }],
  ccRecipients: [{ emailAddress: { address: "sam@example.com" } }],
  webLink: WEB_LINK,
};

/** Never a request that sends, or that touches an existing message. */
function expectNothingSentOrChanged(g: ReturnType<typeof graphWith>): void {
  for (const [path] of g.post.mock.calls) {
    expect(path).not.toMatch(/\/send$|\/sendMail$|\/reply$|\/replyAll$|\/forward$/i);
  }
  expect(g.patch).not.toHaveBeenCalled();
  expect(g.delete).not.toHaveBeenCalled();
}

describe("createDraft (#184)", () => {
  it("saves a plain-text draft to exactly the checked recipients, without sending", async () => {
    const g = graphWith(async () => draftResponse);
    const saved = await createDraft(g, {
      to: ["jane@example.com"],
      cc: ["sam@example.com"],
      subject: "Hi",
      body: "Hello\nthere",
    });
    expect(g.post).toHaveBeenCalledWith("/me/messages", {
      subject: "Hi",
      body: { contentType: "Text", content: "Hello\nthere" },
      toRecipients: [{ emailAddress: { address: "jane@example.com" } }],
      ccRecipients: [{ emailAddress: { address: "sam@example.com" } }],
    });
    expect(saved).toEqual({
      id: "AQMk/ab+c=",
      subject: "Hi",
      to: ["jane@example.com"],
      cc: ["sam@example.com"],
      webLink: WEB_LINK,
    });
    expectNothingSentOrChanged(g);
  });

  it("doesn't claim nothing was saved when the answer was lost", async () => {
    for (const e of [
      new TimeoutError("slow"),
      new GraphError("network", { reason: "network" }),
      new GraphError("server", { status: 503 }),
    ]) {
      const g = graphWith(async () => {
        throw e;
      });
      const err = await createDraft(g, { to: ["a@example.com"], cc: [], subject: "", body: "b" })
        .then(() => undefined)
        .catch((x: unknown) => x);
      expect(err).toBeInstanceOf(ValidationError);
      expect(String(err)).toMatch(/may have been\. Check Drafts/);
      expect(g.post).toHaveBeenCalledOnce();
    }
  });

  it("passes on a refusal unchanged", async () => {
    const refused = new GraphError("forbidden", { status: 403 });
    const g = graphWith(async () => {
      throw refused;
    });
    await expect(
      createDraft(g, { to: ["a@example.com"], cc: [], subject: "", body: "b" }),
    ).rejects.toBe(refused);
  });
});

describe("createReplyDraft (#184)", () => {
  it("drafts a reply in the thread with the text above the quote, without sending", async () => {
    const g = graphWith(async () => ({
      ...draftResponse,
      subject: "RE: Your booking",
      ccRecipients: [],
    }));
    const saved = await createReplyDraft(g, "msg/1", "Thanks!\nSee you then.");
    // Graph's createReply keeps the quote, the thread and the Reply-To; the
    // comment goes above the quote. Nothing else is written.
    expect(g.post).toHaveBeenCalledExactlyOnceWith("/me/messages/msg%2F1/createReply", {
      comment: "Thanks!\nSee you then.",
    });
    expect(saved).toMatchObject({ subject: "RE: Your booking", to: ["jane@example.com"], cc: [] });
    expectNothingSentOrChanged(g);
  });

  it("says to check Drafts when Graph saved one but gave no id, so it isn't saved twice", async () => {
    const g = graphWith(async () => ({ subject: "x" }));
    const err = await createReplyDraft(g, "m", "b").catch((x: unknown) => x);
    // A ValidationError's message reaches the user as is (clientSafeMessage).
    expect(err).toBeInstanceOf(ValidationError);
    expect(clientSafeMessage(err)).toMatch(
      /saved the draft.*Check Drafts in Outlook before trying again/,
    );
  });
});

describe("draft links (#184)", () => {
  it("offers Graph's webLink and compose forms on the same Outlook host", () => {
    const c = draftUrlCandidates({ id: "AQMk/ab+c=", webLink: WEB_LINK });
    expect(c).toEqual({
      webLink: WEB_LINK,
      composeDeeplink: "https://outlook.live.com/mail/0/deeplink/compose/AQMk%2Fab%2Bc%3D",
      draftsFolder: "https://outlook.live.com/mail/0/drafts/id/AQMk%2Fab%2Bc%3D",
    });
    expect(draftUrl({ id: "x", webLink: WEB_LINK })).toBe(WEB_LINK);
  });

  it("opens nothing that isn't Outlook on the web over HTTPS", () => {
    for (const webLink of [
      null,
      "",
      "not a url",
      "http://outlook.live.com/owa/?ItemID=x",
      "https://outlook.live.com.evil.example/owa/?ItemID=x",
      "https://evil.example/owa/?ItemID=x",
      "file:///C:/x",
    ]) {
      expect(draftUrl({ id: "x", webLink }), String(webLink)).toBeUndefined();
      expect(draftUrlCandidates({ id: "x", webLink }), String(webLink)).toEqual({});
    }
    expect(draftUrl({ id: "x", webLink: "https://outlook.office.com/owa/?ItemID=x" })).toBe(
      "https://outlook.office.com/owa/?ItemID=x",
    );
  });
});
