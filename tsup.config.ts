import { defineConfig } from "tsup";

export default defineConfig({
  // The server, and the PDF worker it starts beside it (#156).
  entry: { index: "src/index.ts", "pdf-worker": "src/pdf-worker.ts" },
  format: ["esm"],
  target: "node20",
  platform: "node",
  outDir: "dist",
  clean: true,
  sourcemap: true,
  dts: false,
  splitting: false,
  // Prepend a shebang so the built entrypoint can be launched directly by an
  // MCP client (Claude Desktop / VS Code) as a stdio subprocess.
  banner: {
    js: "#!/usr/bin/env node",
  },
});
