/**
 * MCP tool: whoami (diagnostics).
 */

import { whoami } from "../graph/user.js";
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
  async handler(_input, ctx) {
    const me = await whoami(ctx.graph);
    const email = me.mail ?? me.userPrincipalName ?? "unknown";
    const summary = `Signed in as ${me.displayName ?? email} <${email}>.`;
    return toolResult(summary, me as unknown as Record<string, unknown>);
  },
});

export const userTools: RegisteredTool[] = [toRegistered(whoamiTool)];
