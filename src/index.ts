/**
 * Entry point for the Claude Hotmail Connector — a local stdio MCP server.
 *
 * This is a placeholder scaffold. The MCP server wiring (StdioServerTransport,
 * tool registration, and startup sign-in) is implemented in a later issue
 * (see "MCP server assembly over stdio"). Keeping this minimal lets the build
 * and CI pipeline go green from the first scaffold commit.
 */

export const NAME = "claude-hotmail-connector";

async function main(): Promise<void> {
  // Intentionally does nothing yet — replaced by the stdio MCP server.
  console.error(`[${NAME}] scaffold entrypoint — not yet wired up.`);
}

// Only run when executed directly (not when imported by tests).
if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
