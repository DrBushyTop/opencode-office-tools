import { describe, expect, it } from "vitest";
import { splitSheetQualifiedRange } from "./excelShared";

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
});
