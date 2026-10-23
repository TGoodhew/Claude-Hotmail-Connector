import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    globals: true,
    include: ["src/**/*.test.ts", "test/**/*.test.ts"],
    passWithNoTests: true,
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      // The entry points run as their own processes or threads (the PDF worker is
      // tested through its built file); test-support only builds test inputs.
      exclude: ["src/**/*.test.ts", "src/index.ts", "src/pdf-worker.ts", "src/test-support/**"],
      // CI runs `npm run test:coverage` and fails below these (#47). They sit just
      // under the levels on CI (Linux, where the Windows-only DPAPI paths don't run)
      // on 2026-10-21; raise them as tests grow.
      thresholds: { statements: 91, branches: 86, functions: 88, lines: 92 },
    },
  },
});
