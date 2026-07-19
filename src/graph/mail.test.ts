import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { buildMessagesQuery, getMessage, listMessages, type ListMessagesInput } from "./mail.js";

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

function graphWith(request: unknown, get?: unknown): GraphClient {
  return { request, get } as unknown as GraphClient;
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
      },
    ]);
    expect(res.nextPageToken).toBe("https://graph.microsoft.com/v1.0/me/messages?$skip=10");
    // eventual consistency header is sent for $search.
    expect(request.mock.calls[0]![2]).toMatchObject({ headers: { consistencylevel: "eventual" } });
  });

  it("follows a pageToken verbatim", async () => {
    const request = vi.fn<RequestMock>(async () => ({
      status: 200,
      data: { value: [] },
      headers: new Headers(),
    }));
    await listMessages(graphWith(request), baseInput({ pageToken: "https://graph/next" }));
    expect(request.mock.calls[0]![1]).toBe("https://graph/next");
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
    expect(msg.body).toEqual({ contentType: "text", content: "Ref ABC123" });
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

  it("url-encodes the message id", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "x",
      hasAttachments: false,
      body: { content: "" },
    }));
    await getMessage(graphWith(undefined, get), { id: "AA=BB/CC", format: "text" });
    expect(get.mock.calls[0]![0]).toContain(encodeURIComponent("AA=BB/CC"));
  });
});
