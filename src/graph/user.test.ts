import { describe, expect, it, vi } from "vitest";
import type { GraphClient } from "./client.js";
import { whoami } from "./user.js";
import { whoamiTool } from "../tools/user.js";
import type { ToolContext } from "../tools/types.js";

type GetMock = (path: string, opts?: unknown) => Promise<unknown>;

function graphWith(get: GetMock): GraphClient {
  return { get } as unknown as GraphClient;
}

describe("whoami", () => {
  it("shapes the /me profile", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "abc",
      displayName: "Tony Goodhew",
      userPrincipalName: "tony_goodhew@hotmail.com",
      mail: "tony_goodhew@hotmail.com",
    }));
    const me = await whoami(graphWith(get));
    expect(me).toEqual({
      id: "abc",
      displayName: "Tony Goodhew",
      userPrincipalName: "tony_goodhew@hotmail.com",
      mail: "tony_goodhew@hotmail.com",
    });
    expect(get.mock.calls[0]![0]).toBe("/me");
  });

  it("nulls missing fields", async () => {
    const get = vi.fn<GetMock>(async () => ({ id: "abc" }));
    const me = await whoami(graphWith(get));
    expect(me).toEqual({ id: "abc", displayName: null, userPrincipalName: null, mail: null });
  });
});

describe("whoami tool", () => {
  it("summarises the signed-in account", async () => {
    const get = vi.fn<GetMock>(async () => ({
      displayName: "Tony Goodhew",
      mail: "tony_goodhew@hotmail.com",
    }));
    const ctx: ToolContext = {
      graph: graphWith(get),
      config: {} as ToolContext["config"],
    };
    const res = await whoamiTool.handler({}, ctx);
    expect(res.content[0]!.text).toContain("tony_goodhew@hotmail.com");
    expect(whoamiTool.annotations?.readOnlyHint).toBe(true);
  });
});
