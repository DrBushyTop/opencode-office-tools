import { describe, expect, it } from "vitest";
import {
  formatAvailableShapeTargets,
  invalidSlideIndexMessage,
  roundTripRefreshHint,
  roundTripSlideRefreshHint,
  shouldAddRoundTripRefreshHint,
  toolFailure,
} from "./powerpointShared";

describe("powerpointShared", () => {
  it("formats slide index hints with available values", () => {
    expect(invalidSlideIndexMessage(4, 3)).toBe("Invalid slideIndex 4. Available slideIndex values: 0, 1, 2.");
  });

  it("formats shape target hints with ids and names", () => {
    expect(formatAvailableShapeTargets(0, [
      { id: "shape-1", name: "Title 1" },
      { id: "shape-2", name: "Subtitle 2" },
    ])).toContain("Available shapes on slide 1: shapeIndex 0: id=shape-1, name=\"Title 1\"; shapeIndex 1: id=shape-2, name=\"Subtitle 2\"");
  });

  it("handles unloaded shape names in target hints", () => {
    const shape = {
      id: "shape-1",
      get name() {
        throw new Error("name not loaded");
      },
    } as unknown as { id?: string | null; name?: string | null };

    expect(formatAvailableShapeTargets(0, [shape])).toContain("shapeIndex 0: id=shape-1, name=\"\"");
  });

  it("appends optional hints to failures", () => {
    expect(toolFailure("The object can not be found here.", "Hint: refresh the slide map.").error)
      .toBe("The object can not be found here. Hint: refresh the slide map.");
  });

  it("only suggests refresh hints for stale-object style failures", () => {
    expect(shouldAddRoundTripRefreshHint("The object can not be found here.")).toBe(true);
    expect(shouldAddRoundTripRefreshHint("Invalid slideIndex 4. Available slideIndex values: 0, 1, 2.")).toBe(false);
    expect(shouldAddRoundTripRefreshHint("Shape abc was not found on slide 2.")).toBe(false);
  });

  it("separates shape and slide refresh hints", () => {
    expect(roundTripRefreshHint()).toContain("shapeId values");
    expect(roundTripSlideRefreshHint()).not.toContain("shapeId values");
  });
});
