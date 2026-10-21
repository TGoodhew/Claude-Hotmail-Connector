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
      exclude: ["src/**/*.test.ts", "src/index.ts"],
      // CI runs `npm run test:coverage` and fails below these (#47). They sit
      // just under the levels reached on 2026-10-21; raise them as tests grow.
      thresholds: { statements: 92, branches: 86, functions: 90, lines: 93 },
    },
  },
});
