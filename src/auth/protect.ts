/**
 * OS-level protection for the token cache's AES key (spec v2 §8.3).
 *
 * The token cache itself is AES-256-GCM encrypted (see `tokens.ts`); this
 * module protects the *key* so that copying the cache folder isn't enough to
 * read it. On Windows the key is wrapped with **DPAPI** (`CurrentUser` scope):
 * only the same Windows user on the same machine can unwrap it. DPAPI is
 * reached through the built-in PowerShell `ProtectedData` API, so no native
 * module is needed and the Node SEA exe keeps working. The key travels over
 * stdin, never on the command line.
 *
 * macOS Keychain / Linux libsecret backends are not implemented yet; on those
 * platforms {@link defaultKeyProtector} returns `null` and the token store
 * falls back to a plain key file.
 */

import { spawn } from "node:child_process";
import { win32 } from "node:path";

/** Wraps/unwraps a small secret (the AES key) with an OS-protected mechanism. */
export interface KeyProtector {
  /** Stable identifier recorded in the key file, e.g. `"dpapi"`. */
  readonly name: string;
  protect(data: Buffer): Promise<Buffer>;
  unprotect(data: Buffer): Promise<Buffer>;
}

/** Runs a PowerShell command with `stdin` and resolves with its trimmed stdout. */
export type PowerShellRunner = (command: string, stdin: string) => Promise<string>;

const POWERSHELL_TIMEOUT_MS = 20_000;

/** App-specific DPAPI entropy: another app can't unwrap the blob by calling DPAPI alone. */
const ENTROPY_B64 = Buffer.from("claude-hotmail-connector/token-cache-key/v1").toString("base64");

function dpapiCommand(op: "Protect" | "Unprotect"): string {
  return [
    "$ErrorActionPreference = 'Stop'",
    "Add-Type -AssemblyName System.Security",
    "$data = [Convert]::FromBase64String([Console]::In.ReadToEnd().Trim())",
    `$entropy = [Convert]::FromBase64String('${ENTROPY_B64}')`,
    `$out = [Security.Cryptography.ProtectedData]::${op}($data, $entropy, ` +
      "[Security.Cryptography.DataProtectionScope]::CurrentUser)",
    "[Console]::Out.Write([Convert]::ToBase64String($out))",
  ].join("; ");
}

/**
 * Absolute path to Windows PowerShell 5.1 (always present on Windows). Never
 * spawn it by bare name: Windows would search the current directory (e.g. the
 * project folder Claude Code launched us in) before PATH, and a planted
 * `powershell.exe` there would receive the key on stdin.
 */
export function powershellPath(env: NodeJS.ProcessEnv = process.env): string {
  const systemRoot = env.SystemRoot || env.windir || "C:\\Windows";
  return win32.join(systemRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}

/** Default runner: Windows PowerShell, by absolute path. */
export const runPowerShell: PowerShellRunner = (command, stdin) =>
  new Promise((resolve, reject) => {
    const child = spawn(
      powershellPath(),
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", command],
      { windowsHide: true, stdio: ["pipe", "pipe", "pipe"] },
    );
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("PowerShell DPAPI call timed out."));
    }, POWERSHELL_TIMEOUT_MS);
    child.stdout.on("data", (d: Buffer) => (stdout += d.toString("utf8")));
    child.stderr.on("data", (d: Buffer) => (stderr += d.toString("utf8")));
    child.on("error", (e) => {
      clearTimeout(timer);
      reject(e);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve(stdout.trim());
      // stderr holds PowerShell's own error text, never the key (sent on stdin).
      else reject(new Error(`PowerShell exited with ${code}: ${stderr.trim().slice(0, 300)}`));
    });
    child.stdin.end(stdin);
  });

/** DPAPI (`CurrentUser`) protector. `run` is injectable for tests. */
export function createDpapiProtector(run: PowerShellRunner = runPowerShell): KeyProtector {
  async function call(op: "Protect" | "Unprotect", data: Buffer): Promise<Buffer> {
    const out = Buffer.from(await run(dpapiCommand(op), data.toString("base64")), "base64");
    if (out.length === 0) throw new Error(`DPAPI ${op} returned no data.`);
    return out;
  }
  return {
    name: "dpapi",
    protect: (data) => call("Protect", data),
    unprotect: (data) => call("Unprotect", data),
  };
}

/** The OS protector for this platform, or `null` if none is implemented. */
export function defaultKeyProtector(
  platform: NodeJS.Platform = process.platform,
): KeyProtector | null {
  return platform === "win32" ? createDpapiProtector() : null;
}
