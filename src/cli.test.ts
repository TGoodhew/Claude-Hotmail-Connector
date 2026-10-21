import { describe, expect, it } from "vitest";
import { parseCommand, USAGE } from "./cli.js";
import { SERVER_VERSION } from "./mcp.js";

describe("parseCommand (#39)", () => {
  it("runs the server only with no arguments, as MCP clients launch it", () => {
    expect(parseCommand([])).toEqual({ kind: "server" });
  });

  it("recognises the account commands", () => {
    for (const c of ["login", "logout", "whoami"]) expect(parseCommand([c])).toEqual({ kind: c });
  });

  it("recognises help and version in their usual spellings", () => {
    for (const c of ["help", "-h", "--help"]) expect(parseCommand([c])).toEqual({ kind: "help" });
    for (const c of ["version", "-v", "--version"]) {
      expect(parseCommand([c])).toEqual({ kind: "version" });
    }
  });

  it("rejects a typo instead of silently starting the server", () => {
    expect(parseCommand(["logni"])).toEqual({ kind: "invalid", message: "Unknown command: logni" });
    expect(parseCommand(["--verbose"])).toMatchObject({ kind: "invalid" });
    expect(parseCommand(["Login"])).toMatchObject({ kind: "invalid" }); // case-sensitive
  });

  it("rejects extra arguments", () => {
    expect(parseCommand(["login", "now"])).toEqual({
      kind: "invalid",
      message: "Unexpected argument: now",
    });
  });
});

describe("USAGE", () => {
  it("names every command and the version", () => {
    for (const c of ["login", "logout", "whoami", "help", "version"]) expect(USAGE).toContain(c);
    expect(USAGE).toContain(SERVER_VERSION);
  });
});
