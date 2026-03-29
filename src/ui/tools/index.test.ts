import { describe, expect, it } from "vitest";
import { normalizeToolExecutionResult } from "./index";

describe("normalizeToolExecutionResult", () => {
  it("preserves successful structured tool results", () => {
    const result = {
      textResultForLlm: "Captured image of slide 1 of 3 (800px wide)",
      binaryResultsForLlm: [
        {
          data: "AAA",
          mimeType: "image/png",
          type: "image",
          description: "Slide 1 of 3",
        },
      ],
      resultType: "success",
      toolTelemetry: {},
    };

    expect(normalizeToolExecutionResult(result)).toEqual(result);
  });

  it("throws on failure results", () => {
    expect(() => normalizeToolExecutionResult({
      textResultForLlm: "boom",
      resultType: "failure",
      error: "boom",
      toolTelemetry: {},
    })).toThrow("boom");
  });
});
