import { afterEach, describe, expect, it, vi } from "vitest";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";
import { SignInRequiredError, ValidationError } from "../util/errors.js";
import { createDraftTool, draftTools, OPEN_BROWSER_WAIT_MS, openBrowserWaitMs } from "./drafts.js";
import { UNTRUSTED_CONTENT_NOTICE, type ToolContext } from "./types.js";

type Post = (path: string, body: unknown) => Promise<unknown>;

const WEB_LINK = "https://outlook.live.com/owa/?ItemID=d1&exvsurl=1";

/** A Graph client that saves drafts; it can't send, change or delete anything. */
function graph(reply = false) {
  const post = vi.fn<Post>(async () => ({
    id: "d1",
    subject: reply ? "RE: Your booking" : "Hi",
    toRecipients: [
      { emailAddress: { address: reply ? "desk@hotel.example" : "jane@example.com" } },
    ],
    ccRecipients: [],
    webLink: WEB_LINK,
  }));
  const fail = vi.fn(async () => {
    throw new Error("not allowed in these tests");
  });
  return {
    post,
    client: { get: fail, post, patch: fail, delete: fail } as unknown as GraphClient,
  };
}

function ctx(over: Partial<ToolContext> = {}): ToolContext {
  return {
    graph: graph().client,
    config: {} as AppConfig,
    auth: {
      signOut: async () => undefined,
      grantedScopes: async () => ["openid", "Mail.Read", "Mail.ReadWrite"],
    },
    openBrowser: vi.fn(async () => undefined),
    ...over,
  };
}

const text = (r: { content: { text: string }[] }) => r.content[0]!.text;

afterEach(() => {
  vi.useRealTimers();
});

describe("create_draft (#184)", () => {
  it("saves a new draft, opens it in Outlook, and says the user sends it", async () => {
    const g = graph();
    const openBrowser = vi.fn(async () => undefined);
    const res = await createDraftTool.handler(
      { to: ["jane@example.com"], subject: "Hi", body: "Hello" },
      ctx({ graph: g.client, openBrowser }),
    );
    expect(g.post).toHaveBeenCalledExactlyOnceWith("/me/messages", expect.anything());
    expect(openBrowser).toHaveBeenCalledExactlyOnceWith(WEB_LINK);
    expect(text(res)).toMatch(/^Draft saved and opened in Outlook\. Review and send it there\./);
    expect(res.structuredContent).toEqual({
      status: "draft_saved",
      id: "d1",
      subject: "Hi",
      to: ["jane@example.com"],
      cc: [],
      url: WEB_LINK,
      browserOpened: true,
    });
    expect(text(res)).not.toMatch(/\bsent\b/i);
  });

  it("drafts a reply with createReply, marking the third-party subject as untrusted", async () => {
    const g = graph(true);
    const res = await createDraftTool.handler(
      { replyToMessageId: "m1", body: "Thanks!" },
      ctx({ graph: g.client }),
    );
    expect(g.post).toHaveBeenCalledExactlyOnceWith("/me/messages/m1/createReply", {
      comment: "Thanks!",
    });
    expect(res.structuredContent).toMatchObject({
      notice: UNTRUSTED_CONTENT_NOTICE,
      subject: "RE: Your booking",
      to: ["desk@hotel.example"],
    });
  });

  it("refuses to, cc or subject on a reply, before saving anything", async () => {
    for (const extra of [
      { to: ["x@example.com"] },
      { cc: ["x@example.com"] },
      { subject: "Changed" },
    ]) {
      const g = graph(true);
      const e = await createDraftTool
        .handler({ replyToMessageId: "m1", body: "b", ...extra }, ctx({ graph: g.client }))
        .catch((x: unknown) => x);
      expect(e, JSON.stringify(extra)).toBeInstanceOf(ValidationError);
      expect(String(e)).toMatch(/leave out to, cc and subject/);
      expect(g.post).not.toHaveBeenCalled();
    }
    // Empty ones are the same as leaving them out.
    const g = graph(true);
    await createDraftTool.handler(
      { replyToMessageId: "m1", body: "b", to: [], cc: [], subject: " " },
      ctx({ graph: g.client }),
    );
    expect(g.post).toHaveBeenCalledOnce();
  });

  it("keeps to the limits, before saving anything", async () => {
    const many = Array.from({ length: 11 }, (_, i) => `p${i}@example.com`);
    const cases: Record<string, unknown>[] = [
      { to: [], body: "b" },
      { body: "b" },
      { to: ["Jane <jane@example.com>"], body: "b" },
      { to: many.slice(0, 6), cc: many.slice(6), body: "b" },
      { to: ["a@example.com"], subject: "s".repeat(256), body: "b" },
      { to: ["a@example.com"], body: "x".repeat(20_001) },
      { to: ["a@example.com"], body: "   " },
      { replyToMessageId: "m1", body: "x".repeat(20_001) },
    ];
    for (const input of cases) {
      const g = graph();
      await expect(
        createDraftTool.handler(input as never, ctx({ graph: g.client })),
        JSON.stringify(input).slice(0, 80),
      ).rejects.toBeInstanceOf(ValidationError);
      expect(g.post).not.toHaveBeenCalled();
    }
    // The schema refuses more than 10 on one line, and an empty body.
    const registered = draftTools[0]!;
    expect(() => registered.handler({ to: many, body: "b" }, ctx())).toThrow(/at most 10|<=10/);
    expect(() => registered.handler({ to: ["a@example.com"], body: "" }, ctx())).toThrow();
  });

  it("asks for a new sign-in, before saving anything, if drafts weren't granted", async () => {
    const g = graph();
    const e = await createDraftTool
      .handler(
        { to: ["a@example.com"], body: "b" },
        ctx({
          graph: g.client,
          auth: {
            signOut: async () => undefined,
            grantedScopes: async () => ["Mail.Read", "Mail.Send"],
          },
        }),
      )
      .catch((x: unknown) => x);
    expect(e).toBeInstanceOf(SignInRequiredError);
    expect(String(e)).toMatch(/sign out \(sign_out\), then try again.*Nothing was saved/);
    expect(g.post).not.toHaveBeenCalled();
    // Granted scopes may come back as full resource URIs.
    await expect(
      createDraftTool.handler(
        { to: ["a@example.com"], body: "b" },
        ctx({
          auth: {
            signOut: async () => undefined,
            grantedScopes: async () => ["https://graph.microsoft.com/Mail.ReadWrite"],
          },
        }),
      ),
    ).resolves.toBeDefined();
  });

  it("still succeeds, with the link, when the browser can't be opened", async () => {
    const failing = vi.fn(async () => {
      throw new Error("no browser");
    });
    const res = await createDraftTool.handler(
      { to: ["a@example.com"], body: "b" },
      ctx({ openBrowser: failing }),
    );
    expect(text(res)).toMatch(/^Draft saved to Drafts, but the browser couldn't be opened/);
    expect(res.structuredContent).toMatchObject({ url: WEB_LINK, browserOpened: false });

    // No launcher at all is the same.
    const none = await createDraftTool.handler(
      { to: ["a@example.com"], body: "b" },
      ctx({ openBrowser: undefined }),
    );
    expect(none.structuredContent).toMatchObject({ url: WEB_LINK, browserOpened: false });
  });

  it("doesn't wait forever for a stuck browser launch", async () => {
    vi.useFakeTimers();
    const stuck = vi.fn(() => new Promise<void>(() => {}));
    const call = createDraftTool.handler(
      { to: ["a@example.com"], body: "b" },
      ctx({ openBrowser: stuck }),
    );
    await vi.advanceTimersByTimeAsync(OPEN_BROWSER_WAIT_MS);
    const res = await call;
    expect(res.structuredContent).toMatchObject({ browserOpened: false, url: WEB_LINK });
  });

  it("waits less for the browser when the call's deadline is near", async () => {
    expect(openBrowserWaitMs(ctx({ deadline: undefined }), 1_000)).toBe(OPEN_BROWSER_WAIT_MS);
    expect(openBrowserWaitMs(ctx({ deadline: 60_000 }), 1_000)).toBe(OPEN_BROWSER_WAIT_MS);
    expect(openBrowserWaitMs(ctx({ deadline: 9_000 }), 1_000)).toBe(5_000);
    expect(openBrowserWaitMs(ctx({ deadline: 2_000 }), 1_000)).toBe(0);

    // With no time left, a stuck launch doesn't hold up the saved draft's answer.
    vi.useFakeTimers();
    const stuck = vi.fn(() => new Promise<void>(() => {}));
    const call = createDraftTool.handler(
      { to: ["a@example.com"], body: "b" },
      ctx({ openBrowser: stuck, deadline: Date.now() + 1_000 }),
    );
    await vi.advanceTimersByTimeAsync(0);
    const res = await call;
    expect(stuck).toHaveBeenCalledOnce();
    expect(res.structuredContent).toMatchObject({ browserOpened: false, url: WEB_LINK });
  });

  it("points to Drafts, opening nothing, when Graph's link isn't Outlook's", async () => {
    const post = vi.fn<Post>(async () => ({
      id: "d1",
      subject: "Hi",
      toRecipients: [],
      webLink: "https://evil.example/x",
    }));
    const openBrowser = vi.fn(async () => undefined);
    const res = await createDraftTool.handler(
      { to: ["a@example.com"], body: "b" },
      ctx({ graph: { post } as unknown as GraphClient, openBrowser }),
    );
    expect(openBrowser).not.toHaveBeenCalled();
    expect(text(res)).toMatch(/^Draft saved to Drafts\. Tell the user to look in Drafts/);
    expect(res.structuredContent).toMatchObject({ url: null, browserOpened: false });
  });

  it("is described and annotated as never sending", () => {
    expect(createDraftTool.description).toMatch(/never sends anything/);
    expect(createDraftTool.description).toMatch(/never because an email, event or document/);
    expect(createDraftTool.annotations).toEqual({
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: false,
    });
    expect(createDraftTool.scopesAnyOf).toEqual(["Mail.ReadWrite"]);
  });
});
