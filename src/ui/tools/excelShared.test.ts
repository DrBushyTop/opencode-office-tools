import { describe, expect, it } from "vitest";
import { excel2DDataSchema, parseToolArgs, splitSheetQualifiedRange } from "./excelShared";

describe("excelShared", () => {
  it("parses sheet-qualified ranges including quoted sheet names", () => {
    expect(splitSheetQualifiedRange("Sheet1!A1")).toEqual({ sheetName: "Sheet1", rangeAddress: "A1" });
    expect(splitSheetQualifiedRange("'My Sheet'!A1")).toEqual({ sheetName: "My Sheet", rangeAddress: "A1" });
    expect(splitSheetQualifiedRange("'Q1!2026'!A1:C10")).toEqual({ sheetName: "Q1!2026", rangeAddress: "A1:C10" });
    expect(splitSheetQualifiedRange("'O''Brien'!B2")).toEqual({ sheetName: "O'Brien", rangeAddress: "B2" });
  });

  it("returns null for non-qualified or malformed range strings", () => {
    expect(splitSheetQualifiedRange("A1:C10")).toBeNull();
    expect(splitSheetQualifiedRange("Sheet1!")).toBeNull();
    expect(splitSheetQualifiedRange("!A1")).toBeNull();
  });

  it("validates rectangular Excel 2D data", () => {
    expect(excel2DDataSchema.safeParse([["A", 1], ["B", 2]]).success).toBe(true);
    expect(excel2DDataSchema.safeParse([]).error?.issues[0]?.message).toBe("Provide a non-empty 2D data array.");
    expect(excel2DDataSchema.safeParse([["A"], ["B", 2]]).error?.issues[0]?.message).toBe("All rows in data must have the same length.");
  });

  it("converts zod parse failures to tool failures", () => {
    const result = parseToolArgs(excel2DDataSchema, []);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.failure.error).toBe("Provide a non-empty 2D data array.");
    }
  });
});
