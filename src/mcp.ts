/**
 * MCP server assembly.
 *
 * Registers every tool with the MCP SDK's high-level McpServer. The SDK
 * validates incoming arguments against each tool's input shape; handler errors
 * are caught and returned as a client-safe tool result — canned messages for
 * internal failures, the validation message itself for input errors (see
 * clientSafeMessage) — never thrown to the transport, never leaking internals.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { calendarTools } from "./tools/calendar.js";
import { mailTools } from "./tools/mail.js";
import type { RegisteredTool, ToolContext } from "./tools/types.js";
import { userTools } from "./tools/user.js";
import { clientSafeMessage } from "./util/errors.js";
import { createLogger, type Logger } from "./util/log.js";

export const SERVER_NAME = "claude-hotmail-connector";
export const SERVER_VERSION = "0.1.0";

/** All tools exposed by the connector. */
export function allTools(): RegisteredTool[] {
  return [...userTools, ...mailTools, ...calendarTools];
}

/** Build an McpServer with every tool registered against the given context. */
export function createMcpServer(
  ctx: ToolContext,
  tools: RegisteredTool[] = allTools(),
  logger?: Logger,
): McpServer {
  const log = logger ?? createLogger({ name: "mcp" });
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });

  for (const tool of tools) {
    server.registerTool(
      tool.name,
      {
        title: tool.title ?? tool.annotations?.title,
        description: tool.description,
        inputSchema: tool.inputShape,
        annotations: tool.annotations,
      },
      async (args: Record<string, unknown>): Promise<CallToolResult> => {
        try {
          const result = await tool.handler(args ?? {}, ctx);
          return result as CallToolResult;
        } catch (err) {
          log.error("Tool handler failed", {
            tool: tool.name,
            error: err instanceof Error ? err.message : String(err),
          });
          return {
            content: [{ type: "text", text: clientSafeMessage(err) }],
            isError: true,
          };
        }
      },
    );
  }

  return server;
}
