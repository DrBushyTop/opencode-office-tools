import { describe, expect, it } from "vitest";
import { filterModels } from "./model-search";

const models = [
  {
    key: "anthropic/claude-sonnet-4-5",
    label: "Anthropic / Claude Sonnet 4.5",
    providerID: "anthropic",
    modelID: "claude-sonnet-4-5",
  },
  {
    key: "openai/gpt-5.4",
    label: "OpenAI / GPT-5.4",
    providerID: "openai",
    modelID: "gpt-5.4",
  },
  {
    key: "google/gemini-2.5-pro",
    label: "Google / Gemini 2.5 Pro",
    providerID: "google",
    modelID: "gemini-2.5-pro",
  },
];

describe("filterModels", () => {
  it("keeps exact and fuzzy matches", () => {
    expect(filterModels(models, "cldsn45")[0]?.key).toBe("anthropic/claude-sonnet-4-5");
    expect(filterModels(models, "gm25p")[0]?.key).toBe("google/gemini-2.5-pro");
  });

  it("matches provider and model ids", () => {
    expect(filterModels(models, "openaigpt54")[0]?.key).toBe("openai/gpt-5.4");
  });

  it("omits non-matches", () => {
    expect(filterModels(models, "zzz")).toEqual([]);
  });
});
