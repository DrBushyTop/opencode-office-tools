import { describe, expect, it } from "vitest";
import { getToolNamesForHost, getToolsForHost, normalizeToolExecutionResult } from "./index";

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

  it("resolves every registered PowerPoint tool handler", () => {
    const hostType = {
      Word: "Word",
      PowerPoint: "PowerPoint",
      Excel: "Excel",
      OneNote: "OneNote",
    };

    Object.assign(globalThis, {
      Office: {
        HostType: hostType,
      },
    });

    expect(getToolsForHost(hostType.PowerPoint as never).map((tool) => tool.name).sort()).toEqual(
      [...getToolNamesForHost("powerpoint")].sort(),
    );
  });
});
