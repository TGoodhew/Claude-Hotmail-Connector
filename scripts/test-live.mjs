#!/usr/bin/env node
/**
 * Run the opt-in live Graph test (src/live.test.ts, #48) with HOTMAIL_LIVE_TESTS=1
 * set, on any OS. Sign in first with `node dist/index.js login`.
 *
 * Under Vitest the server's own `.env` lookup (next to `dist/index.js`) doesn't
 * apply, so the checkout's `.env` is loaded here. Variables already set in the
 * environment win, as they do for the server.
 */
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { config as loadDotenv } from "dotenv";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
loadDotenv({ path: join(root, ".env"), quiet: true });

const result = spawnSync("npx", ["vitest", "run", "src/live.test.ts"], {
  cwd: root,
  stdio: "inherit",
  shell: true, // resolves npx (a .cmd on Windows); the arguments contain no spaces
  env: { ...process.env, HOTMAIL_LIVE_TESTS: "1" },
});
process.exit(result.status ?? 1);
