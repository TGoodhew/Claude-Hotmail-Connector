/**
 * Build a self-contained Windows single-executable (the "portable" download and
 * the payload the Inno installer wraps) using Node's official Single Executable
 * Applications (SEA) feature — no external runtime, no extra runtime deps.
 *
 * Prereqoms: Node 20+, and `postject` available via npx.
 * Run on Windows:  npm run build:exe
 *
 * Steps:
 *   1. Bundle the app to a single CommonJS file (dist-sea/bundle.cjs) — SEA needs
 *      one self-contained script (tsup does this).
 *   2. Generate the SEA blob from sea-config.json.
 *   3. Copy the running node.exe to build/hotmail-connector.exe.
 *   4. Inject the blob into that copy with postject.
 *
 * The produced build/hotmail-connector.exe accepts the same subcommands as
 * `node dist/index.js` (default = server, plus setup/login/logout/whoami).
 */
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, rmSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const outDir = join(root, "build");
const exePath = join(outDir, "hotmail-connector.exe");
const blobPath = join(outDir, "sea-prep.blob");

function run(cmd, args) {
  console.log(`$ ${cmd} ${args.join(" ")}`);
  execFileSync(cmd, args, { stdio: "inherit", shell: process.platform === "win32" });
}

if (process.platform !== "win32") {
  console.error("build:exe currently targets Windows. Run it on Windows.");
  process.exit(1);
}

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

// 1. Bundle to a single CJS file for SEA.
run("npx", [
  "tsup",
  "src/index.ts",
  "--format",
  "cjs",
  "--bundle",
  "--platform",
  "node",
  "--target",
  "node20",
  "--out-dir",
  "dist-sea",
  "--clean",
  "--no-splitting",
]);
// tsup emits dist-sea/index.cjs; SEA config expects bundle.cjs.
copyFileSync(join(root, "dist-sea", "index.cjs"), join(root, "dist-sea", "bundle.cjs"));

// 2. Generate the SEA blob.
run(process.execPath, ["--experimental-sea-config", "sea-config.json"]);

// 3. Copy node.exe as our exe base.
copyFileSync(process.execPath, exePath);

// 4. Inject the blob (Windows fuse sentinel).
run("npx", [
  "postject",
  exePath,
  "NODE_SEA_BLOB",
  blobPath,
  "--sentinel-fuse",
  "NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2",
]);

console.log(`\n✓ Built ${exePath}`);
console.log("  Test it:  build\\hotmail-connector.exe whoami");
