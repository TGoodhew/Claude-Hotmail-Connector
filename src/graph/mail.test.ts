import { describe, expect, it, vi } from "vitest";
import { GraphError, ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import {
  buildMessagesQuery,
  DEFAULT_MAX_BODY_CHARS,
  getMessage,
  listMailFolders,
  listMessages,
  MAX_EXCLUDED_SUBFOLDERS,
  MAX_MAIL_FOLDERS,
  sliceBody,
  type ListMessagesInput,
} from "./mail.js";

function baseInput(over: Partial<ListMessagesInput> = {}): ListMessagesInput {
  return { limit: 10, ...over };
}

describe("buildMessagesQuery", () => {
  it("uses $search for free-text query (no $orderby)", () => {
    const { query } = buildMessagesQuery(baseInput({ query: "monsoon aquatics" }));
    expect(query.$search).toBe('"monsoon aquatics"');
    expect(query.$orderby).toBeUndefined();
    expect(query.$top).toBe(10);
    expect(String(query.$select)).toContain("receivedDateTime");
  });

  it("maps from/subjectContains into KQL qualifiers", () => {
    const { query } = buildMessagesQuery(
      baseInput({ from: "bookings@x.com", subjectContains: "confirmation" }),
    );
    expect(query.$search).toBe('"from:bookings@x.com subject:confirmation"');
  });

  it("uses $filter + $orderby for date-only queries", () => {
    const { query } = buildMessagesQuery(
      baseInput({ receivedAfter: "2026-07-01", receivedBefore: "2026-07-31" }),
    );
    expect(query.$search).toBeUndefined();
    expect(String(query.$filter)).toContain("receivedDateTime ge 2026-07-01");
    expect(String(query.$filter)).toContain("receivedDateTime le 2026-07-31");
    expect(query.$orderby).toBe("receivedDateTime desc");
  });

  it("returns a client date filter when search + dates are combined", () => {
    const built = buildMessagesQuery(baseInput({ query: "hotel", receivedAfter: "2026-07-01" }));
    expect(built.query.$search).toBe('"hotel"');
    expect(built.clientDateFilter?.afterMs).toBe(Date.parse("2026-07-01T00:00:00.000Z"));
  });

  it("rejects an invalid date", () => {
    expect(() => buildMessagesQuery(baseInput({ receivedAfter: "not-a-date" }))).toThrow(
      ValidationError,
    );
  });
});

/** Answers the well-known folder lookups (#53) with "F-<name>"; the rest goes to `get`. */
function withFolders(get?: unknown) {
  return async (path: string, opts?: unknown) => {
    const m = /^\/me\/mailFolders\/([a-z]+)$/.exec(path);
    if (m) return { id: `F-${m[1]}` };
    return (get as (p: string, o?: unknown) => Promise<unknown>)(path, opts);
  };
}

function graphWith(request: unknown, get?: unknown): GraphClient {
  return { request, get: withFolders(get) } as unknown as GraphClient;
}

type RequestMock = (
  method: string,
  path: string,
  opts?: unknown,
) => Promise<{ status: number; data: unknown; headers: Headers }>;
type GetMock = (path: string, opts?: unknown) => Promise<unknown>;

describe("listMessages", () => {
  const rawPage = {
    value: [
      {
        id: "m1",
        subject: "Booking confirmation",
        from: { emailAddress: { name: "Monsoon", address: "book@monsoon.com" } },
        receivedDateTime: "2026-07-10T09:00:00Z",
        webLink: "https://outlook/1",
        bodyPreview: "Your booking...",
        hasAttachments: false,
      },
    ],
    "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
  };

  it("shapes messages and returns the nextPageToken", async () => {
    const request = vi.fn<RequestMock>(async () => ({
      status: 200,
      data: rawPage,
      headers: new Headers(),
    }));
    const res = await listMessages(graphWith(request), baseInput({ query: "booking" }));
    expect(res.messages).toEqual([
      {
        id: "m1",
        subject: "Booking confirmation",
        from: { name: "Monsoon", address: "book@monsoon.com" },
        receivedDateTime: "2026-07-10T09:00:00Z",
        webLink: "https://outlook/1",
        bodyPreview: "Your booking...",
        hasAttachments: false,
        folderId: null,
        folder: null,
      },
    ]);
    // An opaque cursor, never the raw URL (#30).
    expect(res.nextPageToken).toBeTruthy();
    expect(res.nextPageToken).not.toContain("https://");
    expect(res.nextPageToken).toMatch(/^[A-Za-z0-9_-]+$/);
    // eventual consistency header is sent for $search.
    expect(request.mock.calls[0]![2]).toMatchObject({ headers: { consistencylevel: "eventual" } });
  });

  it("follows its own cursor as a path relative to Graph, keeping the search header", async () => {
    const request = vi.fn<RequestMock>(async () => ({
      status: 200,
      data: rawPage,
      headers: new Headers(),
    }));
    const first = await listMessages(graphWith(request), baseInput({ query: "booking" }));
    await listMessages(graphWith(request), baseInput({ pageToken: first.nextPageToken }));
    const [, path, opts] = request.mock.calls[1]!;
    expect(path).toBe("me/messages?$skip=10");
    expect(opts).toMatchObject({ headers: { consistencylevel: "eventual" } });
  });

  it("keeps search date bounds on later pages (#31)", async () => {
    const page = (ids: [string, string][], next?: string) => ({
      status: 200,
      data: {
        value: ids.map(([id, receivedDateTime]) => ({ id, receivedDateTime })),
        ...(next ? { "@odata.nextLink": next } : {}),
      },
      headers: new Headers(),
    });
    const request = vi
      .fn<RequestMock>()
      .mockResolvedValueOnce(
        page(
          [
            ["p1-new", "2026-07-10T00:00:00Z"],
            ["p1-old", "2026-06-01T00:00:00Z"],
          ],
          "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc",
        ),
      )
      .mockResolvedValueOnce(
        page([
          ["p2-new", "2026-07-11T00:00:00Z"],
          ["p2-old", "2026-05-01T00:00:00Z"],
        ]),
      );
    const input = baseInput({ query: "hotel", receivedAfter: "2026-07-01" });
    const p1 = await listMessages(graphWith(request), input);
    const p2 = await listMessages(graphWith(request), { ...input, pageToken: p1.nextPageToken });
    expect(p1.messages.map((m) => m.id)).toEqual(["p1-new"]);
    expect(p2.messages.map((m) => m.id)).toEqual(["p2-new"]);
    expect(p2.nextPageToken).toBeUndefined();
  });

  it("rejects a pageToken that is a URL (prompt-injection channel, #30)", async () => {
    const request = vi.fn<RequestMock>();
    await expect(
      listMessages(graphWith(request), baseInput({ pageToken: "https://evil.example/steal" })),
    ).rejects.toBeInstanceOf(ValidationError);
    expect(request).not.toHaveBeenCalled();
  });

  it("rejects a forged cursor that points at an absolute URL", async () => {
    const forged = Buffer.from(
      JSON.stringify({ v: 1, next: "https://evil.example/steal", mode: "filter" }),
    ).toString("base64url");
    const request = vi.fn<RequestMock>();
    await expect(
      listMessages(graphWith(request), baseInput({ pageToken: forged })),
    ).rejects.toThrow(/not a valid cursor/);
    expect(request).not.toHaveBeenCalled();
  });

  it("drops a nextLink outside Graph instead of handing it back", async () => {
    const request = vi.fn<RequestMock>(async () => ({
      status: 200,
      data: { value: [], "@odata.nextLink": "https://evil.example/next" },
      headers: new Headers(),
    }));
    const res = await listMessages(graphWith(request), baseInput());
    expect(res.nextPageToken).toBeUndefined();
  });

  it("applies the client-side date filter for search + dates", async () => {
    const page = {
      value: [
        { id: "old", receivedDateTime: "2026-06-01T00:00:00Z" },
        { id: "new", receivedDateTime: "2026-07-10T00:00:00Z" },
      ],
    };
    const request = vi.fn(async () => ({ status: 200, data: page, headers: new Headers() }));
    const res = await listMessages(
      graphWith(request),
      baseInput({ query: "x", receivedAfter: "2026-07-01" }),
    );
    expect(res.messages.map((m) => m.id)).toEqual(["new"]);
  });
});

describe("getMessage", () => {
  it("converts an HTML body to text and lists attachment metadata", async () => {
    const get = vi.fn(async (path: string) => {
      if (path.includes("/attachments")) {
        return {
          value: [
            {
              id: "a1",
              name: "ticket.pdf",
              contentType: "application/pdf",
              size: 1234,
              isInline: false,
            },
          ],
        };
      }
      return {
        id: "m1",
        subject: "Confirmation",
        from: { emailAddress: { name: "X", address: "x@y.com" } },
        receivedDateTime: "2026-07-10T09:00:00Z",
        webLink: "https://outlook/1",
        hasAttachments: true,
        body: { contentType: "html", content: "<p>Ref <b>ABC123</b></p>" },
      };
    });
    const msg = await getMessage(graphWith(undefined, get), { id: "m1", format: "text" });
    expect(msg.body).toEqual({
      contentType: "text",
      content: "Ref ABC123",
      truncated: false,
      totalChars: 10,
    });
    expect(msg.attachments).toEqual([
      { id: "a1", name: "ticket.pdf", contentType: "application/pdf", size: 1234, isInline: false },
    ]);
    expect(msg.from).toEqual({ name: "X", address: "x@y.com" });
  });

  it("returns raw HTML when format=html and skips attachment fetch when none", async () => {
    const get = vi.fn(async () => ({
      id: "m2",
      subject: "Plain",
      hasAttachments: false,
      body: { contentType: "html", content: "<p>Keep me</p>" },
    }));
    const msg = await getMessage(graphWith(undefined, get), { id: "m2", format: "html" });
    expect(msg.body.contentType).toBe("html");
    expect(msg.body.content).toBe("<p>Keep me</p>");
    expect(msg.attachments).toEqual([]);
    expect(get).toHaveBeenCalledTimes(1); // no attachments call
  });

  it("reports the content type the body really has (#35)", async () => {
    const plain = vi.fn(async () => ({
      id: "m4",
      hasAttachments: false,
      body: { contentType: "text", content: "Just text" },
    }));
    // Asked for HTML, but the message is plain text: say so.
    const asHtml = await getMessage(graphWith(undefined, plain), { id: "m4", format: "html" });
    expect(asHtml.body).toMatchObject({ contentType: "text", content: "Just text" });
    const asText = await getMessage(graphWith(undefined, plain), { id: "m4", format: "text" });
    expect(asText.body.contentType).toBe("text");
  });

  it("url-encodes the message id", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "x",
      hasAttachments: false,
      body: { content: "" },
    }));
    await getMessage(graphWith(undefined, get), { id: "AA=BB/CC", format: "text" });
    expect(get.mock.calls[0]![0]).toContain(encodeURIComponent("AA=BB/CC"));
  });

  it("reads a message with a malformed numeric entity instead of failing (#32)", async () => {
    const get = vi.fn(async () => ({
      id: "m3",
      hasAttachments: false,
      body: { contentType: "html", content: "<p>Total&#99999999; &euro;240</p>" },
    }));
    const msg = await getMessage(graphWith(undefined, get), { id: "m3", format: "text" });
    expect(msg.body.content).toBe("Total� €240");
  });
});

describe("body size cap (#38)", () => {
  it("returns short bodies whole", () => {
    expect(sliceBody("hello", "text")).toEqual({
      contentType: "text",
      content: "hello",
      truncated: false,
      totalChars: 5,
    });
  });

  it("cuts long bodies at maxChars and says where to read on", () => {
    const body = sliceBody("abcdefghij", "text", 4);
    expect(body).toEqual({
      contentType: "text",
      content: "abcd",
      truncated: true,
      totalChars: 10,
      nextStartChar: 4,
    });
    expect(sliceBody("abcdefghij", "text", 4, 8)).toEqual({
      contentType: "text",
      content: "ij",
      truncated: false,
      totalChars: 10,
      startChar: 8,
    });
  });

  it("defaults to 20,000 characters", () => {
    const body = sliceBody("x".repeat(DEFAULT_MAX_BODY_CHARS + 5), "text");
    expect(body.content).toHaveLength(DEFAULT_MAX_BODY_CHARS);
    expect(body).toMatchObject({ truncated: true, nextStartChar: DEFAULT_MAX_BODY_CHARS });
  });

  it("never splits an emoji (surrogate pair) at either end", () => {
    const smile = String.fromCodePoint(0x1f600); // two UTF-16 units
    const text = `ab${smile}cd`; // a b [hi lo] c d
    expect(sliceBody(text, "text", 3)).toMatchObject({ content: "ab", nextStartChar: 2 });
    expect(sliceBody(text, "text", 3, 3)).toMatchObject({ content: `${smile}c`, startChar: 2 });
    // Too small to back off: the whole pair is taken, so reading always advances.
    expect(sliceBody(text, "text", 1, 2)).toMatchObject({ content: smile, nextStartChar: 4 });
  });

  it("isn't misled by a lone surrogate next to an emoji (Copilot review on #117)", () => {
    const smile = String.fromCodePoint(0x1f600);
    const lone = String.fromCharCode(0xdc00); // malformed: a low surrogate with no pair
    const text = "a".repeat(998) + smile + lone + "tail";
    const first = sliceBody(text, "text", 1_000);
    // The cut at 1000 falls between the emoji and the lone unit, which isn't a pair.
    expect(first.content.endsWith(smile)).toBe(true);
    expect(first.nextStartChar).toBe(1_000);
  });

  it("reads any body back exactly, part by part", () => {
    const smile = String.fromCodePoint(0x1f600);
    const lone = String.fromCharCode(0xdc00);
    const texts = [
      "a".repeat(998) + smile + lone + "tail",
      (smile + "x").repeat(700),
      smile.repeat(1_500),
      lone + smile + lone.repeat(3) + "z".repeat(2_500),
    ];
    for (const text of texts) {
      for (const max of [1, 2, 3, 1_000]) {
        let out = "";
        let start = 0;
        for (let guard = 0; guard < 10_000; guard += 1) {
          const part = sliceBody(text, "text", max, start);
          out += part.content;
          if (!part.truncated) break;
          expect(part.nextStartChar).toBeGreaterThan(start); // always advances
          start = part.nextStartChar!;
        }
        expect(out).toBe(text);
      }
    }
  });

  it("caps get_message bodies after converting HTML to text", async () => {
    const get = vi.fn(async () => ({
      id: "big",
      hasAttachments: false,
      body: { contentType: "html", content: `<p>${"word ".repeat(10_000)}</p>` },
    }));
    const msg = await getMessage(graphWith(undefined, get), {
      id: "big",
      format: "text",
      maxChars: 1_000,
    });
    expect(msg.body.content).toHaveLength(1_000);
    expect(msg.body).toMatchObject({ truncated: true, totalChars: 49_999, nextStartChar: 1_000 });
  });
});

describe("mail folders (#53)", () => {
  const msg = (id: string, parentFolderId?: string) => ({
    id,
    subject: id,
    receivedDateTime: "2026-07-10T09:00:00Z",
    hasAttachments: false,
    ...(parentFolderId ? { parentFolderId } : {}),
  });
  const pageOf = (value: unknown[], next?: string) => ({
    status: 200,
    data: { value, ...(next ? { "@odata.nextLink": next } : {}) },
    headers: new Headers(),
  });
  const mixed = [
    msg("in", "F-inbox"),
    msg("junk", "F-junkemail"),
    msg("bin", "F-deleteditems"),
    msg("custom", "custom-folder-id"),
  ];

  it("leaves Junk Email and Deleted Items out by default, says how many, and labels the rest", async () => {
    const request = vi.fn<RequestMock>(async () => pageOf(mixed));
    const res = await listMessages(graphWith(request), baseInput());
    expect(request.mock.calls[0]![1]).toBe("/me/messages");
    expect(res.messages.map((m) => [m.id, m.folder, m.folderId])).toEqual([
      ["in", "inbox", "F-inbox"],
      ["custom", null, "custom-folder-id"],
    ]);
    expect(res.excludedCount).toBe(2);
  });

  it("keeps them, labelled, when asked", async () => {
    const request = vi.fn<RequestMock>(async () => pageOf(mixed));
    const res = await listMessages(graphWith(request), baseInput({ includeJunkAndDeleted: true }));
    expect(res.messages.map((m) => m.folder)).toEqual(["inbox", "junkemail", "deleteditems", null]);
    expect(res.excludedCount).toBe(0);
  });

  /** A get mock with a tree of subfolders beneath the well-known folders. */
  function treeGet(children: Record<string, string[]>, inFlight?: { now: number; max: number }) {
    const count = (id: string) => children[id]?.length ?? 0;
    return vi.fn<GetMock>(async (path) => {
      if (inFlight) inFlight.max = Math.max(inFlight.max, ++inFlight.now);
      await new Promise((r) => setTimeout(r, 1));
      if (inFlight) inFlight.now--;
      const known = /^\/me\/mailFolders\/([a-z]+)$/.exec(path);
      if (known) return { id: `F-${known[1]}`, childFolderCount: count(`F-${known[1]}`) };
      const kids = /^\/me\/mailFolders\/([^/]+)\/childFolders$/.exec(path);
      if (kids) {
        const parent = decodeURIComponent(kids[1]!);
        return {
          value: (children[parent] ?? []).map((id) => ({ id, childFolderCount: count(id) })),
        };
      }
      throw new Error(`unexpected ${path}`);
    });
  }

  it("leaves out folders beneath Deleted Items and Junk Email too", async () => {
    // A deleted folder moves under Deleted Items; its messages stay filed in it.
    const get = treeGet({
      "F-deleteditems": ["old-project"],
      "old-project": ["old-project/2025"],
      "F-junkemail": ["junk-sub"],
      "F-inbox": ["receipts"],
    });
    const request = vi.fn<RequestMock>(async () =>
      pageOf([
        msg("a", "old-project"),
        msg("b", "old-project/2025"),
        msg("c", "junk-sub"),
        msg("d", "receipts"),
      ]),
    );
    const res = await listMessages({ request, get } as unknown as GraphClient, baseInput());
    expect(res.messages.map((m) => m.id)).toEqual(["d"]);
    expect(res.excludedCount).toBe(3);
    // Only the excluded trees are walked.
    const walked = get.mock.calls.map(([p]) => p).filter((p) => p.endsWith("/childFolders"));
    expect(walked.sort()).toEqual([
      "/me/mailFolders/F-deleteditems/childFolders",
      "/me/mailFolders/F-junkemail/childFolders",
      "/me/mailFolders/old-project/childFolders",
    ]);
    // Hidden folders are walked too: the listing returns their messages.
    for (const [p, o] of get.mock.calls) {
      if (p.endsWith("/childFolders")) {
        expect((o as { query?: Record<string, unknown> }).query?.includeHiddenFolders).toBe("true");
      }
    }
  });

  it("keeps folder lookups to three at a time, alongside the listing", async () => {
    const inFlight = { now: 0, max: 0 };
    const get = treeGet(
      {
        "F-deleteditems": ["s1", "s2", "s3", "s4"],
        s1: ["t1"],
        s2: ["t2"],
        s3: ["t3"],
        s4: ["t4"],
      },
      inFlight,
    );
    const request = vi.fn<RequestMock>(async () => pageOf([msg("x", "t4")]));
    const res = await listMessages({ request, get } as unknown as GraphClient, baseInput());
    expect(res.excludedCount).toBe(1);
    expect(inFlight.max).toBe(3);
  });

  it("refuses, rather than leak, when the excluded folders have too many subfolders", async () => {
    const many = Array.from({ length: MAX_EXCLUDED_SUBFOLDERS + 1 }, (_, i) => `s${i}`);
    const request = vi.fn<RequestMock>(async () => pageOf([]));
    const tooMany = { request, get: treeGet({ "F-deleteditems": many }) };
    await expect(listMessages(tooMany as unknown as GraphClient, baseInput())).rejects.toThrow(
      /too many subfolders/,
    );
    // A folder with more children than one page holds, likewise.
    const paged = vi.fn<GetMock>(async (path) =>
      path.endsWith("/childFolders")
        ? { value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next" }
        : { id: `F-${path.split("/").pop()}`, childFolderCount: 1 },
    );
    await expect(
      listMessages({ request, get: paged } as unknown as GraphClient, baseInput()),
    ).rejects.toThrow(/too many subfolders/);
    // A huge tree is refused as soon as the count is passed, not walked in full.
    const wide = Array.from({ length: MAX_EXCLUDED_SUBFOLDERS }, (_, i) => `w${i}`);
    const tree: Record<string, string[]> = { "F-deleteditems": wide };
    for (const w of wide)
      tree[w] = Array.from({ length: MAX_EXCLUDED_SUBFOLDERS }, (_, i) => `${w}-${i}`);
    const huge = treeGet(tree);
    await expect(
      listMessages({ request, get: huge } as unknown as GraphClient, baseInput()),
    ).rejects.toThrow(/too many subfolders/);
    const walks = huge.mock.calls.filter(([p]) => p.endsWith("/childFolders")).length;
    expect(walks).toBeLessThanOrEqual(1 + 3); // the root's, then at most one batch in flight
    // Opting in doesn't need the walk.
    await listMessages(
      tooMany as unknown as GraphClient,
      baseInput({ includeJunkAndDeleted: true }),
    );
  });

  it("lists one folder by well-known name or id, without leaving anything out", async () => {
    const request = vi.fn<RequestMock>(async () => pageOf([msg("j", "F-junkemail")]));
    const junk = await listMessages(
      graphWith(request),
      baseInput({ folder: "junkemail", query: "hotel" }),
    );
    expect(request.mock.calls[0]![1]).toBe("/me/mailFolders/junkemail/messages");
    expect(junk.messages.map((m) => m.folder)).toEqual(["junkemail"]);
    await listMessages(graphWith(request), baseInput({ folder: "AAMk/x+y=" }));
    expect(request.mock.calls[1]![1]).toBe("/me/mailFolders/AAMk%2Fx%2By%3D/messages");
  });

  it("refuses a folder that would resolve away in the URL", async () => {
    // "/me/mailFolders/../messages" would become "/me/messages": everything, nothing left out.
    const request = vi.fn<RequestMock>(async () => pageOf([]));
    for (const folder of [".", "..", " .. ", "   "]) {
      await expect(listMessages(graphWith(request), baseInput({ folder }))).rejects.toBeInstanceOf(
        ValidationError,
      );
    }
    expect(request).not.toHaveBeenCalled();
  });

  it("keeps leaving them out on later pages", async () => {
    const request = vi
      .fn<RequestMock>()
      .mockResolvedValueOnce(
        pageOf([msg("a", "F-inbox")], "https://graph.microsoft.com/v1.0/me/messages?$skip=10"),
      )
      .mockResolvedValueOnce(pageOf([msg("b", "F-junkemail"), msg("c", "F-inbox")]));
    const p1 = await listMessages(graphWith(request), baseInput());
    const p2 = await listMessages(graphWith(request), baseInput({ pageToken: p1.nextPageToken }));
    expect(p2.messages.map((m) => m.id)).toEqual(["c"]);
    expect(p2.excludedCount).toBe(1);
    // …and an opted-in listing keeps including them on later pages.
    const r2 = vi
      .fn<RequestMock>()
      .mockResolvedValueOnce(
        pageOf([msg("a", "F-inbox")], "https://graph.microsoft.com/v1.0/me/messages?$skip=10"),
      )
      .mockResolvedValueOnce(pageOf([msg("b", "F-junkemail")]));
    const q1 = await listMessages(graphWith(r2), baseInput({ includeJunkAndDeleted: true }));
    const q2 = await listMessages(graphWith(r2), baseInput({ pageToken: q1.nextPageToken }));
    expect(q2.messages.map((m) => m.id)).toEqual(["b"]);
  });

  it("copes with a mailbox that lacks a well-known folder (e.g. no Archive)", async () => {
    const request = vi.fn<RequestMock>(async () => pageOf([msg("in", "F-inbox")]));
    const get = async (path: string) => {
      if (path === "/me/mailFolders/archive") throw new GraphError("nope", { status: 404 });
      return { id: `F-${path.slice(16)}` };
    };
    const res = await listMessages({ request, get } as unknown as GraphClient, baseInput());
    expect(res.messages.map((m) => m.folder)).toEqual(["inbox"]);
    // Any other failure still surfaces.
    const failing = async () => {
      throw new GraphError("boom", { status: 500 });
    };
    await expect(
      listMessages({ request, get: failing } as unknown as GraphClient, baseInput()),
    ).rejects.toThrow("boom");
  });

  it("get_message says which folder the message is in", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "m1",
      parentFolderId: "F-junkemail",
      body: { contentType: "text", content: "x" },
      hasAttachments: false,
    }));
    const res = await getMessage(graphWith(undefined, get), { id: "m1", format: "text" });
    expect(res).toMatchObject({ folderId: "F-junkemail", folder: "junkemail" });
  });

  it("lists the top-level folders, with their well-known names", async () => {
    const getAllPages = vi.fn(async () => [
      {
        id: "F-inbox",
        displayName: "Inbox",
        childFolderCount: 1,
        unreadItemCount: 3,
        totalItemCount: 40,
      },
      { id: "F-junkemail", displayName: "Junk Email" },
      { id: "trips", displayName: "Trips", parentFolderId: "root" },
    ]);
    const graph = { getAllPages, get: withFolders() } as unknown as GraphClient;
    const folders = await listMailFolders(graph);
    expect(folders).toEqual([
      {
        id: "F-inbox",
        name: "Inbox",
        wellKnownName: "inbox",
        parentFolderId: null,
        childFolderCount: 1,
        unreadItemCount: 3,
        totalItemCount: 40,
      },
      {
        id: "F-junkemail",
        name: "Junk Email",
        wellKnownName: "junkemail",
        parentFolderId: null,
        childFolderCount: 0,
        unreadItemCount: 0,
        totalItemCount: 0,
      },
      {
        id: "trips",
        name: "Trips",
        wellKnownName: null,
        parentFolderId: "root",
        childFolderCount: 0,
        unreadItemCount: 0,
        totalItemCount: 0,
      },
    ]);
    expect(getAllPages).toHaveBeenCalledWith(
      "/me/mailFolders",
      expect.objectContaining({ query: expect.objectContaining({ $top: MAX_MAIL_FOLDERS }) }),
      { maxItems: MAX_MAIL_FOLDERS },
    );
  });
});
