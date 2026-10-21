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
    "Return the display name and email address of the personal Outlook.com / Hotmail account " +
    "this connector is signed in to. Useful to confirm the correct account is connected. " +
    "Read-only.",
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

export const userTools: RegisteredTool[] = [toRegistered(whoamiTool)];
