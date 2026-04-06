import { describe, expect, it } from "vitest";
import {
  describeError,
  formatAvailableShapeTargets,
  invalidSlideIndexMessage,
  roundTripRefreshHint,
  roundTripSlideRefreshHint,
  shouldAddRoundTripRefreshHint,
  shouldAddRoundTripShapeTargetRefreshHint,
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

  it("adds animation refresh hints for stale shape-id misses without catching plain index validation", () => {
    expect(shouldAddRoundTripShapeTargetRefreshHint("The object can not be found here.")).toBe(true);
    expect(shouldAddRoundTripShapeTargetRefreshHint("Shape abc was not found on slide 2. Available shapes on slide 2: shapeIndex 0: id=def, name=\"Title\"")).toBe(true);
    expect(shouldAddRoundTripShapeTargetRefreshHint("Invalid slideIndex 4. Available slideIndex values: 0, 1, 2.")).toBe(false);
    expect(shouldAddRoundTripShapeTargetRefreshHint("Provide a valid shapeId or shapeIndex for slide 2. Available shapes on slide 2: shapeIndex 0: id=def, name=\"Title\"")).toBe(false);
  });

  it("separates shape and slide refresh hints", () => {
    expect(roundTripRefreshHint()).toContain("shapeId values");
    expect(roundTripSlideRefreshHint()).not.toContain("shapeId values");
  });

  describe("describeError with Office.js debugInfo", () => {
    it("skips redundant [code] suffix when code equals message", () => {
      const error = Object.assign(new Error("InvalidArgument"), { code: "InvalidArgument" });
      expect(describeError(error)).toBe("InvalidArgument");
    });

    it("includes [code] suffix when code differs from message", () => {
      const error = Object.assign(new Error("Something went wrong"), { code: "GeneralException" });
      expect(describeError(error)).toBe("Something went wrong [GeneralException]");
    });

    it("includes debugInfo.errorLocation when available", () => {
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: { errorLocation: "Shapes.addGeometricShape" },
      });
      expect(describeError(error)).toContain("Shapes.addGeometricShape");
    });

    it("includes debugInfo.message when it differs from both message and code", () => {
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: { message: "The argument is invalid or missing", errorLocation: "Fill.setSolidColor" },
      });
      const result = describeError(error);
      expect(result).toContain("The argument is invalid or missing");
      expect(result).toContain("Fill.setSolidColor");
    });

    it("does not duplicate debugInfo.message when it matches the error message", () => {
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: { message: "InvalidArgument" },
      });
      expect(describeError(error)).toBe("InvalidArgument");
    });
  });
});
