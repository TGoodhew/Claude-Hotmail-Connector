/**
 * Opaque paging cursors (spec v2 §7.2, #30).
 *
 * A `nextPageToken` is never a raw URL: it is base64url JSON holding the next
 * page's path *relative to the Graph base* plus whatever later pages need to
 * behave like the first (each list tool adds its own fields). A relative path
 * can only ever be resolved against Graph, and the Graph client rejects
 * absolute URLs outside it besides.
 */

import { z } from "zod";
import { GRAPH_BASE_URL } from "../config.js";
import { ValidationError } from "../util/errors.js";

/** The fields every cursor has; a tool's cursor schema extends this. */
export const CursorBase = z.object({
  v: z.literal(1),
  next: z
    .string()
    .min(1)
    .refine((p) => /^[A-Za-z]/.test(p) && !p.includes("://") && !p.includes("\\"), {
      message: "must be a relative Graph path",
    }),
});

const GRAPH_PREFIX = GRAPH_BASE_URL.replace(/\/?$/, "/");

/**
 * The cursor for Graph's `@odata.nextLink`, carrying `fields` alongside it.
 * Graph's own nextLinks live under the base; anything else is dropped, never
 * followed, and there is then no cursor.
 */
export function encodeCursor(
  nextLink: string | undefined,
  fields: Record<string, unknown>,
): string | undefined {
  if (!nextLink?.startsWith(GRAPH_PREFIX)) return undefined;
  const cursor = { v: 1, next: nextLink.slice(GRAPH_PREFIX.length), ...fields };
  return Buffer.from(JSON.stringify(cursor)).toString("base64url");
}

/** Decode a cursor made by {@link encodeCursor} for `tool`, or refuse it. */
export function decodeCursor<S extends z.ZodTypeAny>(
  token: string,
  schema: S,
  tool: string,
): z.infer<S> {
  let raw: unknown;
  try {
    raw = JSON.parse(Buffer.from(token, "base64url").toString("utf8"));
  } catch {
    raw = undefined;
  }
  const parsed = schema.safeParse(raw);
  if (!parsed.success) {
    throw new ValidationError(
      `pageToken is not a valid cursor. Pass the nextPageToken from a previous ${tool} call ` +
        "unchanged (it is not a URL).",
    );
  }
  return parsed.data;
}
