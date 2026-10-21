import { describe, expect, it } from "vitest";
import { allTools } from "../mcp.js";

describe("tool descriptions", () => {
  it("identifies the personal Outlook.com/Hotmail account scope", () => {
    const tools = allTools();
    expect(tools).toHaveLength(6);
    for (const tool of tools) {
      expect(tool.description).toMatch(/personal Microsoft Outlook\.com\/Hotmail/);
    }
  });
});
