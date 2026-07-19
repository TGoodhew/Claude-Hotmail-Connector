/**
 * User profile (User.Read) — used by the whoami diagnostic to confirm which
 * account is connected.
 */

import type { GraphClient } from "./client.js";

export interface WhoAmIResult {
  id: string | null;
  displayName: string | null;
  userPrincipalName: string | null;
  mail: string | null;
}

interface GraphUser {
  id?: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
}

/** Return the signed-in user's basic profile via GET /me. */
export async function whoami(graph: GraphClient): Promise<WhoAmIResult> {
  const me = await graph.get<GraphUser>("/me", {
    query: { $select: "id,displayName,userPrincipalName,mail" },
  });
  return {
    id: me.id ?? null,
    displayName: me.displayName ?? null,
    userPrincipalName: me.userPrincipalName ?? null,
    mail: me.mail ?? null,
  };
}
