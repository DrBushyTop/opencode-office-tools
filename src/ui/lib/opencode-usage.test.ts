import { describe, expect, it } from "vitest";
import { formatTokenUsage, getLatestSessionUsage } from "./opencode-usage";

describe("opencode usage", () => {
  it("formats token usage with context percentage when available", () => {
    expect(formatTokenUsage({ total: 77123, providerID: "anthropic", modelID: "claude-sonnet-4-5" }, [
      { providerID: "anthropic", modelID: "claude-sonnet-4-5", limitContext: 140000 },
    ])).toBe("77.1K (55%)");
  });

  it("omits percentage when context limit is unavailable", () => {
    expect(formatTokenUsage({ total: 77123, providerID: "anthropic", modelID: "claude-sonnet-4-5" }, [])).toBe("77.1K");
  });

  it("reads the latest assistant usage from opencode messages", () => {
    expect(getLatestSessionUsage([
      {
        info: {
          id: "user-1",
          role: "user",
          time: { created: 1 },
        },
      },
      {
        info: {
          id: "assistant-1",
          role: "assistant",
          time: { created: 2 },
          providerID: "anthropic",
          modelID: "claude-sonnet-4-5",
          tokens: {
            input: 70000,
            output: 4000,
            reasoning: 100,
            cache: { read: 3000, write: 23 },
          },
        },
      },
    ])).toEqual({
      total: 77123,
      providerID: "anthropic",
      modelID: "claude-sonnet-4-5",
    });
  });
});
