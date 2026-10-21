import { describe, expect, it, vi } from "vitest";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";
import { ConfigError } from "../util/errors.js";
import type { ToolContext } from "./types.js";
import { signOutTool, userTools, whoamiTool } from "./user.js";

function ctx(over: Partial<ToolContext> = {}): ToolContext {
  return { graph: {} as GraphClient, config: {} as AppConfig, ...over };
}

describe("whoami tool", () => {
  it("summarises the signed-in account", async () => {
    const get = vi.fn(async () => ({ displayName: "Jane Doe", mail: "jane@outlook.com" }));
    const res = await whoamiTool.handler({}, ctx({ graph: { get } as unknown as GraphClient }));
    expect(res.content[0]!.text).toContain("Signed in as Jane Doe <jane@outlook.com>.");
  });
});

describe("sign_out tool (#43)", () => {
  it("signs out through the auth service and reports it", async () => {
    const signOut = vi.fn(async () => undefined);
    const res = await signOutTool.handler({}, ctx({ auth: { signOut } }));
    expect(signOut).toHaveBeenCalledOnce();
    expect(res.content[0]!.text).toContain("Signed out");
    expect(res.structuredContent).toEqual({ signedOut: true });
  });

  it("fails clearly when no auth service is wired in", async () => {
    await expect(signOutTool.handler({}, ctx())).rejects.toBeInstanceOf(ConfigError);
  });

  it("is a local, destructive, idempotent action that is always available", () => {
    expect(signOutTool.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: false,
    });
    expect(signOutTool.scopesAnyOf).toEqual(["openid"]);
    expect(signOutTool.description).toContain("Only call this when the user asks");
    expect(userTools.map((t) => t.name)).toEqual(["whoami", "sign_out"]);
  });
});
