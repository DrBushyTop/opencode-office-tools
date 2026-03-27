import { describe, expect, it } from "vitest";
import { isValidWorkbookNamedRangeName } from "./manageNamedRange";

describe("manageNamedRange helpers", () => {
  it("accepts valid workbook-scoped Excel names", () => {
    expect(isValidWorkbookNamedRangeName("Budget")).toBe(true);
    expect(isValidWorkbookNamedRangeName("_Budget")).toBe(true);
    expect(isValidWorkbookNamedRangeName("\\Budget")).toBe(true);
    expect(isValidWorkbookNamedRangeName("Sales.Q1")).toBe(true);
  });

  it("rejects invalid or reference-like workbook-scoped Excel names", () => {
    expect(isValidWorkbookNamedRangeName("A1")).toBe(false);
    expect(isValidWorkbookNamedRangeName("R1C1")).toBe(false);
    expect(isValidWorkbookNamedRangeName("Quarter 1")).toBe(false);
    expect(isValidWorkbookNamedRangeName("1Budget")).toBe(false);
  });
});
