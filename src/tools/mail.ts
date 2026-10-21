/**
 * MCP tool definitions for mail (read-only).
 */

import { z } from "zod";
import { DEFAULT_MAX_BODY_CHARS, getMessage, listMessages } from "../graph/mail.js";
import {
  defineTool,
  toRegistered,
  UNTRUSTED_CONTENT_DESCRIPTION,
  untrustedResult,
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
    "Search or list email in the user's personal Outlook.com / Hotmail mailbox (e.g. to find " +
    "booking or confirmation emails). " +
    "Read-only. Returns id, subject, sender, receivedDateTime, webLink, a body preview, " +
    "and hasAttachments, plus a nextPageToken when more results exist. Use get_message to " +
    "read a full message body. " +
    UNTRUSTED_CONTENT_DESCRIPTION,
  inputShape: listMessagesShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Mail.Read"],
  async handler(input, ctx) {
    const result = await listMessages(ctx.graph, input);
    const summary =
      result.messages.length === 0
        ? "No matching messages found."
        : `Found ${result.messages.length} message(s)` +
          (result.nextPageToken ? " (more available — use nextPageToken)." : ".");
    return untrustedResult(summary, {
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
  maxChars: z
    .number()
    .int()
    .min(1_000)
    .max(100_000)
    .default(DEFAULT_MAX_BODY_CHARS)
    .describe("Most body characters to return (1,000-100,000, default 20,000)."),
  startChar: z
    .number()
    .int()
    .min(0)
    .default(0)
    .describe("Where in the body to start: a previous result's body.nextStartChar."),
};

export const getMessageTool: ToolDefinition<typeof getMessageShape> = defineTool({
  name: "get_message",
  title: "Read an email",
  description:
    "Read a single email from the user's personal Outlook.com / Hotmail mailbox by id, " +
    "returning subject, sender, receivedDateTime, webLink, " +
    "the body (HTML converted to plain text by default), and attachment metadata " +
    "(names/sizes only — no file contents). Long bodies are cut at maxChars: body.truncated " +
    "is then true, and startChar=body.nextStartChar reads on. Read-only. " +
    UNTRUSTED_CONTENT_DESCRIPTION,
  inputShape: getMessageShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Mail.Read"],
  async handler(input, ctx) {
    const msg = await getMessage(ctx.graph, input);
    // The summary stays free of the message's own text; that is all in the data.
    const { body } = msg;
    const part = body.truncated
      ? ` The body is long: this is characters ${body.startChar ?? 0}-${body.nextStartChar} of ` +
        `${body.totalChars}; call again with startChar=${body.nextStartChar} for more.`
      : "";
    return untrustedResult(`Read the message; its details follow.${part}`, msg);
  },
});

export const mailTools: RegisteredTool[] = [
  toRegistered(listMessagesTool),
  toRegistered(getMessageTool),
];
