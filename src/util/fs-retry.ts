/**
 * Retrying file operations through Windows' brief sharing violations. Windows
 * refuses to open, rename or delete a file for a moment while another process
 * has it open, or is replacing or deleting it (EPERM / EACCES / EBUSY). Shared
 * by the cross-process coordination (#126) and the token store (#142).
 */

/** Attempts before giving up, and the pause between them. */
const ATTEMPTS = 50;
const PAUSE_MS = 10;

export function isErrno(e: unknown, code: string): boolean {
  return typeof e === "object" && e !== null && (e as NodeJS.ErrnoException).code === code;
}

/** Whether `e` is one of the transient sharing errors above. */
export function isTransientFsError(e: unknown): boolean {
  return isErrno(e, "EPERM") || isErrno(e, "EACCES") || isErrno(e, "EBUSY");
}

/**
 * Run `op`, retrying it for up to about half a second while it fails
 * transiently (see {@link isTransientFsError}). Any other error, or the last
 * transient one, is thrown.
 */
export async function retrying<T>(op: () => Promise<T>): Promise<T> {
  for (let attempt = 1; ; attempt++) {
    try {
      return await op();
    } catch (e) {
      if (!isTransientFsError(e) || attempt >= ATTEMPTS) throw e;
      await new Promise((r) => setTimeout(r, PAUSE_MS));
    }
  }
}
