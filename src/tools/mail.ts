/**
 * MCP tool definitions for mail (read-only).
 */

import { z } from "zod";
import { getMessage, listMessages } from "../graph/mail.js";
import {
  defineTool,
  toolResult,
  toRegistered,
  type RegisteredTool,
  type ToolDefinition,
} from "./types.js";

const listMessagesShape = {
  query: z
    .string()
    .optional()
    .describe("Free-text keyword search (maps to Graph $search, KQL-style)."),
  from: z.string().optional().describe("Sender email address to match."),
  subjectContains: z.string().optional().describe("Text the subject should contain."),
  receivedAfter: z
    .string()
    .optional()
    .describe("Only messages received on/after this ISO 8601 date/time."),
  receivedBefore: z
    .string()
    .optional()
    .describe("Only messages received on/before this ISO 8601 date/time."),
  limit: z.number().int().min(1).max(25).default(10).describe("Max results (1-25, default 10)."),
  pageToken: z.string().optional().describe("Opaque token from a previous call's nextPageToken."),
};

export const listMessagesTool: ToolDefinition<typeof listMessagesShape> = defineTool({
  name: "list_messages",
  title: "List / search email",
  description:
    "Search or list email messages (e.g. to find booking or confirmation emails). " +
    "Read-only. Returns id, subject, sender, receivedDateTime, webLink, a body preview, " +
    "and hasAttachments, plus a nextPageToken when more results exist. Use get_message to " +
    "read a full message body.",
  inputShape: listMessagesShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  async handler(input, ctx) {
    const result = await listMessages(ctx.graph, input);
    const summary =
      result.messages.length === 0
        ? "No matching messages found."
        : `Found ${result.messages.length} message(s)` +
          (result.nextPageToken ? " (more available — use nextPageToken)." : ".");
    return toolResult(summary, {
      count: result.messages.length,
      messages: result.messages,
      ...(result.nextPageToken ? { nextPageToken: result.nextPageToken } : {}),
    });
  },
});

const getMessageShape = {
  id: z.string().min(1).describe("The message id (from list_messages)."),
  format: z
    .enum(["text", "html"])
    .default("text")
    .describe("Body format. 'text' converts HTML to plain text (default, token-efficient)."),
};

export const getMessageTool: ToolDefinition<typeof getMessageShape> = defineTool({
  name: "get_message",
  title: "Read an email",
  description:
    "Read a single email message by id, returning subject, sender, receivedDateTime, webLink, " +
    "the body (HTML converted to plain text by default), and attachment metadata " +
    "(names/sizes only — no file contents). Read-only.",
  inputShape: getMessageShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  async handler(input, ctx) {
    const msg = await getMessage(ctx.graph, input);
    const summary = `${msg.subject ?? "(no subject)"} — from ${msg.from?.address ?? "unknown"}`;
    return toolResult(summary, msg as unknown as Record<string, unknown>);
  },
});

export const mailTools: RegisteredTool[] = [
  toRegistered(listMessagesTool),
  toRegistered(getMessageTool),
];
