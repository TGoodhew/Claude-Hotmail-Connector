import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
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
