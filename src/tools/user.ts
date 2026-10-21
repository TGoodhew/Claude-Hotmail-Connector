/**
 * MCP tools for the account: whoami (diagnostics) and sign_out.
 */

import { whoami } from "../graph/user.js";
import { ConfigError } from "../util/errors.js";
import {
  defineTool,
  toolResult,
  toRegistered,
  type RegisteredTool,
  type ToolDefinition,
} from "./types.js";

const whoamiShape = {};

export const whoamiTool: ToolDefinition<typeof whoamiShape> = defineTool({
  name: "whoami",
  title: "Show the connected account",
  description:
    "Return the signed-in Microsoft account's display name and email address. Useful to confirm " +
    "the correct personal account is connected. Read-only.",
  inputShape: whoamiShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["User.Read"],
  async handler(_input, ctx) {
    const me = await whoami(ctx.graph);
    const email = me.mail ?? me.userPrincipalName ?? "unknown";
    const summary = `Signed in as ${me.displayName ?? email} <${email}>.`;
    return toolResult(summary, me);
  },
});

const signOutShape = {};

// Extension installs have no command line, so this is their sign-out (#43).
// Misuse is low-stakes: an injected sign-out only means signing in again.
export const signOutTool: ToolDefinition<typeof signOutShape> = defineTool({
  name: "sign_out",
  title: "Sign out of the Microsoft account",
  description:
    "Sign out of the user's personal Outlook.com / Hotmail account on this computer: " +
    "deletes the connector's saved sign-in (token cache and its encryption key). Only call this " +
    "when the user asks to sign out or to remove the connector's data. The next request that " +
    "needs the account will open a Microsoft sign-in in the browser.",
  inputShape: signOutShape,
  annotations: {
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: true,
    openWorldHint: false,
  },
  // Needs no Graph permission; `openid` is always requested, so it's always available.
  scopesAnyOf: ["openid"],
  async handler(_input, ctx) {
    if (!ctx.auth) throw new ConfigError("Sign-out isn't available in this context.");
    await ctx.auth.signOut();
    return toolResult("Signed out. The saved Microsoft sign-in was removed from this computer.", {
      signedOut: true,
    });
  },
});

export const userTools: RegisteredTool[] = [toRegistered(whoamiTool), toRegistered(signOutTool)];
