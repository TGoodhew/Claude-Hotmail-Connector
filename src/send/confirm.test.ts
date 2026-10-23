import { afterEach, describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import {
  createSendConfirmer,
  escapeHtml,
  type OutgoingEmail,
  type SendConfirmer,
} from "./confirm.js";

const quiet = { debug() {}, info() {}, warn() {}, error() {} } as never;

const EMAIL: OutgoingEmail = {
  kind: "new",
  to: [
    { address: "jane@example.com", name: "Jane", known: true },
    { address: "stranger@evil.example", known: false },
  ],
  cc: [{ address: "pat@example.com", known: null }],
  subject: "Dinner <b>Friday</b>",
  body: "See you at 7.\n<script>alert(1)</script>",
};

const open: SendConfirmer[] = [];
afterEach(async () => {
  for (const c of open.splice(0)) await c.close();
});

/** A confirmer whose "browser" records the URL it was asked to open. */
function setup(opts: Partial<Parameters<typeof createSendConfirmer>[0]> = {}) {
  const urls: string[] = [];
  const confirmer = createSendConfirmer({
    openBrowser: async (u) => {
      urls.push(u);
    },
    logger: quiet,
    ...opts,
  });
  open.push(confirmer);
  return { confirmer, urls, lastUrl: () => urls[urls.length - 1]! };
}

const post = (url: string, decision: string, headers: Record<string, string> = {}) =>
  fetch(url, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded", ...headers },
    body: `decision=${decision}`,
  });

/** Wait until the page has been opened, then return its URL. */
async function opened(urls: string[]): Promise<string> {
  await vi.waitFor(() => expect(urls.length).toBeGreaterThan(0));
  return urls[urls.length - 1]!;
}

describe("send confirmation page (#57)", () => {
  it("shows exactly what will be sent, escaped, with new recipients flagged", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = confirmer.confirm(EMAIL, send, 10_000);
    const url = await opened(urls);
    expect(url).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/send\/[0-9a-f]{32}$/);

    const res = await fetch(url);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-security-policy")).toContain("default-src 'none'");
    expect(res.headers.get("x-frame-options")).toBe("DENY");
    expect(res.headers.get("cache-control")).toBe("no-store");
    const html = await res.text();
    expect(html).toContain("Jane &lt;jane@example.com&gt;");
    expect(html).toContain('stranger@evil.example <span class="new">has never emailed you</span>');
    expect(html).toContain(`pat@example.com <span class="unknown">(couldn't check past mail)`);
    expect(html).toContain("Dinner &lt;b&gt;Friday&lt;/b&gt;");
    expect(html).toContain("&lt;script&gt;alert(1)&lt;/script&gt;");
    expect(html).not.toContain("<script>");
    expect(html).toContain("goes to someone who has never emailed you");
    expect(send).not.toHaveBeenCalled(); // looking sends nothing

    await post(url, "cancel");
    await outcome;
  });

  it("sends only when Send is clicked, once, and then the page is gone", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = confirmer.confirm(EMAIL, send, 10_000);
    const url = await opened(urls);
    const res = await post(url, "send");
    expect(res.status).toBe(200);
    expect(await res.text()).toContain("The email was sent");
    expect(await outcome).toEqual({ status: "sent" });
    expect(send).toHaveBeenCalledOnce();
    // Used: nothing more can be done with it (and the idle server has stopped).
    const again = await post(url, "send").catch(() => undefined);
    expect(again === undefined || again.status === 404).toBe(true);
    expect(send).toHaveBeenCalledOnce();
  });

  it("sends nothing when the user chooses Don't send", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = confirmer.confirm(EMAIL, send, 10_000);
    const res = await post(await opened(urls), "cancel");
    expect(await res.text()).toContain("wasn't sent");
    expect(await outcome).toEqual({ status: "cancelled" });
    expect(send).not.toHaveBeenCalled();
  });

  it("reports waiting when there's no decision in time, and still sends on a later click", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = await confirmer.confirm(EMAIL, send, 20);
    expect(outcome.status).toBe("pending");
    expect(send).not.toHaveBeenCalled();
    const res = await post(await opened(urls), "send");
    expect(res.status).toBe(200);
    expect(send).toHaveBeenCalledOnce();
  });

  it("expires: an old page sends nothing", async () => {
    const { confirmer, urls } = setup({ expiresMs: 30 });
    const send = vi.fn(async () => undefined);
    await confirmer.confirm(EMAIL, send, 5);
    const url = await opened(urls);
    await new Promise((r) => setTimeout(r, 80));
    const res = await post(url, "send").catch(() => undefined);
    expect(res === undefined || res.status === 404).toBe(true);
    expect(send).not.toHaveBeenCalled();
  });

  it("refuses a post from another site, sending nothing", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = confirmer.confirm(EMAIL, send, 10_000);
    const url = await opened(urls);
    expect((await post(url, "send", { origin: "https://evil.example" })).status).toBe(403);
    expect((await post(url, "send", { "sec-fetch-site": "cross-site" })).status).toBe(403);
    expect(send).not.toHaveBeenCalled();
    // The page itself may still decide.
    const own = new URL(url).origin;
    expect((await post(url, "send", { origin: own, "sec-fetch-site": "same-origin" })).status).toBe(
      200,
    );
    expect(await outcome).toEqual({ status: "sent" });
  });

  it("ignores anything but a clear Send or Don't send, and unknown pages", async () => {
    const { confirmer, urls } = setup();
    const send = vi.fn(async () => undefined);
    const outcome = confirmer.confirm(EMAIL, send, 10_000);
    const url = await opened(urls);
    expect((await post(url, "maybe")).status).toBe(400);
    expect((await fetch(url, { method: "PUT" })).status).toBe(405);
    const base = new URL(url);
    expect((await fetch(`${base.origin}/send/${"0".repeat(32)}`)).status).toBe(404);
    expect((await fetch(`${base.origin}/`)).status).toBe(404);
    expect(send).not.toHaveBeenCalled();
    await post(url, "cancel");
    await outcome;
  });

  it("reports a failed send without claiming it went", async () => {
    const { confirmer, urls } = setup();
    const outcome = confirmer.confirm(
      EMAIL,
      async () => {
        throw new ValidationError("Recipient rejected.");
      },
      10_000,
    );
    const res = await post(await opened(urls), "send");
    expect(res.status).toBe(502);
    expect(await res.text()).toContain("Sending failed: Recipient rejected.");
    expect(await outcome).toEqual({ status: "failed", message: "Recipient rejected." });
  });

  it("allows only a few confirmations waiting at once", async () => {
    const { confirmer } = setup({ maxPending: 2 });
    const send = vi.fn(async () => undefined);
    await confirmer.confirm(EMAIL, send, 1);
    await confirmer.confirm(EMAIL, send, 1);
    await expect(confirmer.confirm(EMAIL, send, 1)).rejects.toThrow(/already waiting/);
  });

  it("sends nothing, and says so, when the browser can't be opened", async () => {
    const confirmer = createSendConfirmer({
      openBrowser: async () => {
        throw new Error("no browser");
      },
      logger: quiet,
    });
    open.push(confirmer);
    const send = vi.fn(async () => undefined);
    await expect(confirmer.confirm(EMAIL, send, 1_000)).rejects.toThrow(/Nothing was sent/);
    expect(send).not.toHaveBeenCalled();
  });

  it("escapes every character HTML treats specially", () => {
    expect(escapeHtml(`<a href="x" onclick='y'>&</a>`)).toBe(
      "&lt;a href=&quot;x&quot; onclick=&#39;y&#39;&gt;&amp;&lt;/a&gt;",
    );
  });
});
