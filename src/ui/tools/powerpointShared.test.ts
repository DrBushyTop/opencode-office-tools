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

    it("includes debugInfo.statement from extended error logging", () => {
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: {
          errorLocation: "Shape.geometricShapeType",
          statement: 'shape.geometricShapeType = "RoundRectangle"',
        },
      });
      const result = describeError(error);
      expect(result).toContain("Shape.geometricShapeType");
      expect(result).toContain('statement: shape.geometricShapeType = "RoundRectangle"');
    });

    it("includes surrounding statements when provided", () => {
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: {
          statement: "shape.lineFormat.weight = 0.75",
          surroundingStatements: [
            "shape = shapes.addGeometricShape(\"Rectangle\")",
            "shape.fill.setSolidColor(\"#0B1E36\")",
            "shape.lineFormat.weight = 0.75",
          ],
        },
      });
      const result = describeError(error);
      expect(result).toContain("surrounding statements:");
      expect(result).toContain("shape = shapes.addGeometricShape");
      expect(result).toContain("shape.lineFormat.weight = 0.75");
    });

    it("truncates very long surrounding statement lists", () => {
      const lines = Array.from({ length: 12 }, (_, index) => `statement_${index}()`);
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: { surroundingStatements: lines },
      });
      const result = describeError(error);
      expect(result).toContain("statement_0");
      expect(result).toContain("(6 more)");
      expect(result).not.toContain("statement_11");
    });

    it("unwraps a single level of innerError", () => {
      const inner = Object.assign(new Error("GenericException"), {
        code: "GenericException",
        debugInfo: { message: "The underlying shape could not be created", errorLocation: "Fill.setSolidColor" },
      });
      const error = Object.assign(new Error("InvalidArgument"), {
        code: "InvalidArgument",
        debugInfo: { innerError: inner },
      });
      const result = describeError(error);
      expect(result).toContain("caused by:");
      expect(result).toContain("The underlying shape could not be created");
      expect(result).toContain("Fill.setSolidColor");
    });
  });
});
