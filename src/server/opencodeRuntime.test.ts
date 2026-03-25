import { describe, expect, it } from "vitest";

const { configuredModels, readResponseBody } = require("./opencodeRuntime.js");

describe("opencode runtime helpers", () => {
  it("returns only configured default models", () => {
    const models = configuredModels(
      {
        providers: [
          {
            id: "anthropic",
            name: "Anthropic",
            models: {
              "claude-sonnet-4-5": { id: "claude-sonnet-4-5", name: "Claude Sonnet 4.5" },
              "claude-opus-4": { id: "claude-opus-4", name: "Claude Opus 4" },
            },
          },
          {
            id: "openai",
            name: "OpenAI",
            models: {
              "gpt-5.4": { id: "gpt-5.4", name: "GPT-5.4" },
            },
          },
        ],
        default: {
          chat: "openai/gpt-5.4",
          fast: "anthropic/claude-sonnet-4-5",
        },
      },
      { model: "openai/gpt-5.4" },
    );

    expect(models).toEqual([
      {
        key: "openai/gpt-5.4",
        label: "OpenAI / GPT-5.4",
        providerID: "openai",
        modelID: "gpt-5.4",
      },
      {
        key: "anthropic/claude-sonnet-4-5",
        label: "Anthropic / Claude Sonnet 4.5",
        providerID: "anthropic",
        modelID: "claude-sonnet-4-5",
      },
    ]);
  });

  it("returns null for empty successful responses", async () => {
    const value = await readResponseBody({
      status: 202,
      text: async () => "",
      headers: { get: () => "application/json" },
    });

    expect(value).toBeNull();
  });

  it("parses json responses with body text", async () => {
    const value = await readResponseBody({
      status: 200,
      text: async () => '{"ok":true}',
      headers: { get: () => "application/json; charset=utf-8" },
    });

    expect(value).toEqual({ ok: true });
  });
});
