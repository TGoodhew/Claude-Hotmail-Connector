/**
 * Shared types for MCP tool definitions.
 *
 * Each tool exposes a zod input *shape* (registered with the MCP SDK, which
 * validates incoming arguments) plus a handler that receives the validated
 * input and a {@link ToolContext}. {@link toRegistered} erases the specific
 * input type so a heterogeneous set of tools can be stored and registered
 * uniformly, while {@link defineTool} keeps full inference at the definition
 * site (and for direct unit testing).
 */

import { z, type ZodRawShape } from "zod";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";

export interface ToolContext {
  graph: GraphClient;
  config: AppConfig;
}

/** MCP behaviour hints surfaced to the client/model. */
export interface ToolAnnotations {
  title?: string;
  readOnlyHint?: boolean;
  destructiveHint?: boolean;
  idempotentHint?: boolean;
  openWorldHint?: boolean;
}

/** A minimal MCP tool result (compatible with the SDK's CallToolResult). */
export interface ToolResult {
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
  isError?: boolean;
}

export interface ToolDefinition<Shape extends ZodRawShape = ZodRawShape> {
  name: string;
  title?: string;
  description: string;
  inputShape: Shape;
  annotations?: ToolAnnotations;
  handler: (input: z.infer<z.ZodObject<Shape>>, ctx: ToolContext) => Promise<ToolResult>;
}

/** A tool whose handler input type has been erased for uniform registration. */
export interface RegisteredTool {
  name: string;
  title?: string;
  description: string;
  inputShape: ZodRawShape;
  annotations?: ToolAnnotations;
  handler: (input: Record<string, unknown>, ctx: ToolContext) => Promise<ToolResult>;
}

/** Helper to define a tool with full input-type inference. */
export function defineTool<Shape extends ZodRawShape>(
  def: ToolDefinition<Shape>,
): ToolDefinition<Shape> {
  return def;
}

/**
 * Erase a tool's specific input type, wrapping its handler so it validates the
 * incoming arguments against the tool's own schema before delegating.
 */
export function toRegistered<Shape extends ZodRawShape>(
  def: ToolDefinition<Shape>,
): RegisteredTool {
  const schema = z.object(def.inputShape);
  return {
    name: def.name,
    title: def.title,
    description: def.description,
    inputShape: def.inputShape,
    annotations: def.annotations,
    handler: (input, ctx) => def.handler(schema.parse(input), ctx),
  };
}

/** Build a standard tool result: a human summary followed by the JSON payload. */
export function toolResult<T extends object>(summary: string, data: T): ToolResult {
  return {
    content: [{ type: "text", text: `${summary}\n\n${JSON.stringify(data, null, 2)}` }],
    structuredContent: data as Record<string, unknown>,
  };
}
