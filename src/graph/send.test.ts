import { describe, expect, it, vi } from "vitest";
import { GraphError, SignInRequiredError, TimeoutError, ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import {
  checkNewEmail,
  hasEmailedYou,
  knownSenders,
  MAX_RECIPIENTS,
  MAX_SEND_BODY_CHARS,
  prepareReply,
  sendNewEmail,
  sendReply,
} from "./send.js";

type Get = (path: string, opts?: { query?: Record<string, unknown> }) => Promise<unknown>;
type Post = (path: string, body: unknown) => Promise<unknown>;

const graphWith = (get?: Get, post?: Post) =>
  ({ get: vi.fn(get), post: vi.fn(post) }) as unknown as GraphClient & {
    get: ReturnType<typeof vi.fn<Get>>;
    post: ReturnType<typeof vi.fn<Post>>;
  };

describe("checkNewEmail (#57)", () => {
  const ok = { to: ["jane@example.com"], subject: " Hi ", body: "Hello" };

  it("tidies the recipients: trimmed, no duplicates, nobody on both To and Cc", () => {
    expect(
      checkNewEmail({
        ...ok,
        to: [" jane@example.com", "JANE@example.com", "pat@example.com"],
        cc: ["pat@example.com", "sam@example.com", "sam@example.com"],
      }),
    ).toEqual({
      to: ["jane@example.com", "pat@example.com"],
      cc: ["sam@example.com"],
      subject: "Hi",
      body: "Hello",
    });
  });

  it("takes plain addresses only", () => {
    for (const bad of [
      "Jane <jane@example.com>",
      "jane@example",
      "a,b@example.com",
      "jane @example.com",
      "",
    ]) {
      expect(() => checkNewEmail({ ...ok, to: [bad] }), bad).toThrow(ValidationError);
    }
  });

  it("keeps to the limits, saying nothing was sent", () => {
    const many = Array.from({ length: MAX_RECIPIENTS + 1 }, (_, i) => `p${i}@example.com`);
    expect(() => checkNewEmail({ ...ok, to: many })).toThrow(/at most 10 people/);
    expect(() => checkNewEmail({ ...ok, to: many.slice(0, 6), cc: many.slice(6) })).toThrow(
      /at most 10/,
    );
    expect(() => checkNewEmail({ ...ok, to: [] })).toThrow(/who the email is to/);
    expect(() => checkNewEmail({ ...ok, body: "   " })).toThrow(/no text/);
    expect(() => checkNewEmail({ ...ok, body: "x".repeat(MAX_SEND_BODY_CHARS + 1) })).toThrow(
      /at most 20,000 characters/,
    );
    expect(() => checkNewEmail({ ...ok, subject: "s".repeat(256) })).toThrow(/subject/);
    expect(() => checkNewEmail({ ...ok, body: "" })).toThrow(/Nothing was sent/);
  });
});

describe("hasEmailedYou / knownSenders (#57)", () => {
  it("looks for one message from the address, quoting it safely", async () => {
    const graph = graphWith(async () => ({ value: [{ id: "m1" }] }));
    expect(await hasEmailedYou(graph, "o'brien@example.com")).toBe(true);
    expect(graph.get.mock.calls[0]![1]?.query).toEqual({
      $filter: "from/emailAddress/address eq 'o''brien@example.com'",
      $top: 1,
      $select: "id",
    });
    expect(
      await hasEmailedYou(
        graphWith(async () => ({ value: [] })),
        "a@example.com",
      ),
    ).toBe(false);
  });

  it("says it couldn't check when Graph won't answer, but doesn't hide a sign-in", async () => {
    const refused = graphWith(async () => {
      throw new GraphError("bad filter", { status: 400 });
    });
    expect(await hasEmailedYou(refused, "a@example.com")).toBeNull();
    const signIn = graphWith(async () => {
      throw new SignInRequiredError("Sign in first.");
    });
    await expect(hasEmailedYou(signIn, "a@example.com")).rejects.toBeInstanceOf(
      SignInRequiredError,
    );
  });

  it("checks several addresses, keyed in lower case", async () => {
    const graph = graphWith(async (_p, o) => ({
      value: String(o?.query?.$filter).toLowerCase().includes("known") ? [{ id: "x" }] : [],
    }));
    const map = await knownSenders(graph, ["Known@example.com", "new@example.com"]);
    expect(map).toEqual(
      new Map([
        ["known@example.com", true],
        ["new@example.com", false],
      ]),
    );
  });
});

describe("sendNewEmail (#57)", () => {
  it("sends plain text to exactly the checked recipients, saving it to Sent Items", async () => {
    const graph = graphWith(undefined, async () => undefined);
    await sendNewEmail(graph, {
      to: ["jane@example.com"],
      cc: ["pat@example.com"],
      subject: "Hi",
      body: "Hello <b>",
    });
    expect(graph.post).toHaveBeenCalledWith("/me/sendMail", {
      message: {
        subject: "Hi",
        body: { contentType: "Text", content: "Hello <b>" },
        toRecipients: [{ emailAddress: { address: "jane@example.com" } }],
        ccRecipients: [{ emailAddress: { address: "pat@example.com" } }],
      },
      saveToSentItems: true,
    });
  });

  it("doesn't claim nothing was sent when the answer was lost", async () => {
    const email = { to: ["a@example.com"], cc: [], subject: "s", body: "b" };
    for (const lost of [new TimeoutError("t"), new GraphError("network", {})]) {
      const graph = graphWith(undefined, async () => {
        throw lost;
      });
      await expect(sendNewEmail(graph, email)).rejects.toThrow(/may have been/);
    }
    const refused = graphWith(undefined, async () => {
      throw new GraphError("forbidden", { status: 403 });
    });
    await expect(sendNewEmail(refused, email)).rejects.toBeInstanceOf(GraphError);
  });
});

describe("replies (#57)", () => {
  const message = (m: Record<string, unknown>) => graphWith(async () => m);

  it("answers the Reply-To address if there is one, else the sender", async () => {
    const withReplyTo = await prepareReply(
      message({
        subject: "Booking",
        from: { emailAddress: { address: "noreply@hotel.example", name: "Hotel" } },
        replyTo: [{ emailAddress: { address: "desk@hotel.example", name: "Front desk" } }],
      }),
      "m/1",
      "Thanks!",
    );
    expect(withReplyTo).toEqual({
      messageId: "m/1",
      to: [{ address: "desk@hotel.example", name: "Front desk" }],
      subject: "RE: Booking",
      body: "Thanks!",
      original: { subject: "Booking", from: "noreply@hotel.example" },
    });
    const plain = await prepareReply(
      message({ subject: "Re: Booking", from: { emailAddress: { address: "jane@example.com" } } }),
      "m2",
      "OK",
    );
    expect(plain.to).toEqual([{ address: "jane@example.com" }]);
    expect(plain.subject).toBe("Re: Booking");
  });

  it("refuses a message with nobody to reply to, or too many", async () => {
    await expect(prepareReply(message({ subject: "x" }), "m", "hi")).rejects.toThrow(
      /no sender to reply to/,
    );
    const many = Array.from({ length: 11 }, (_, i) => ({
      emailAddress: { address: `p${i}@example.com` },
    }));
    await expect(prepareReply(message({ subject: "x", replyTo: many }), "m", "hi")).rejects.toThrow(
      /more than 10 people/,
    );
  });

  it("sends exactly the recipients and text shown, in the thread, without quoting", async () => {
    const graph = graphWith(undefined, async () => undefined);
    await sendReply(graph, {
      messageId: "m/1",
      to: [{ address: "desk@hotel.example", name: "Front desk" }],
      subject: "RE: Booking",
      body: "Thanks!",
      original: { subject: "Booking", from: null },
    });
    expect(graph.post).toHaveBeenCalledWith("/me/messages/m%2F1/reply", {
      message: {
        toRecipients: [{ emailAddress: { address: "desk@hotel.example", name: "Front desk" } }],
        body: { contentType: "Text", content: "Thanks!" },
      },
    });
  });
});
