#!/usr/bin/env node
/**
 * Run the opt-in live Graph test (src/live.test.ts, #48) with HOTMAIL_LIVE_TESTS=1
 * set, on any OS. Sign in first with `node dist/index.js login`.
 */
import { spawnSync } from "node:child_process";

const result = spawnSync("npx", ["vitest", "run", "src/live.test.ts"], {
  stdio: "inherit",
  shell: true, // resolves npx (a .cmd on Windows); the arguments contain no spaces
  env: { ...process.env, HOTMAIL_LIVE_TESTS: "1" },
});
process.exit(result.status ?? 1);
