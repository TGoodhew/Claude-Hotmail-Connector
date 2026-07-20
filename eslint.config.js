import js from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    ignores: ["dist/**", "node_modules/**", "coverage/**"],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: "module",
    },
    rules: {
      // stderr logging is intentional (stdout is reserved for MCP JSON-RPC).
      "no-console": "off",
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
    },
  },
  {
    // Node build scripts (plain JS): declare the Node globals they use.
    files: ["scripts/**/*.{js,mjs,cjs}"],
    languageOptions: {
      globals: { process: "readonly", console: "readonly" },
    },
  },
);
