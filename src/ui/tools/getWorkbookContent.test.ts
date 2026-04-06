import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("./excelShared", async () => {
  const actual = await vi.importActual<any>("./excelShared");
  return {
    ...actual,
    getWorksheet: vi.fn(),
    describeRange: vi.fn(),
  };
});

describe("getWorkbookContent", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.resetModules();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("describes an explicit range with detail options forwarded", async () => {
    const excelShared = await import("./excelShared");
    const worksheet = { name: "Budget", getRange: vi.fn(() => ({ address: "Budget!A1:B2" })) };
    vi.mocked(excelShared.getWorksheet).mockResolvedValue(worksheet as never);
    vi.mocked(excelShared.describeRange).mockResolvedValue("described range");
    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: { sync: ReturnType<typeof vi.fn> }) => Promise<unknown>) => callback({ sync: vi.fn() })),
    };

    const { getWorkbookContent } = await import("./getWorkbookContent");
    const result = await getWorkbookContent.handler({ sheetName: "Budget", range: "A1:B2", detail: true });

    expect(excelShared.getWorksheet).toHaveBeenCalled();
    expect(worksheet.getRange).toHaveBeenCalledWith("A1:B2");
    expect(excelShared.describeRange).toHaveBeenCalledWith(expect.any(Object), { address: "Budget!A1:B2" }, "Budget", {
      detail: true,
      includeNumberFormats: true,
      includeTables: true,
      includeValidation: true,
      includeMergedAreas: true,
    });
    expect(result).toBe("described range");
  });

  it("returns the empty-range message for worksheets with no used range", async () => {
    const excelShared = await import("./excelShared");
    const usedRange = { isNullObject: true, load: vi.fn() };
    const worksheet = { name: "Budget", getUsedRangeOrNullObject: vi.fn(() => usedRange) };
    vi.mocked(excelShared.getWorksheet).mockResolvedValue(worksheet as never);
    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: { sync: ReturnType<typeof vi.fn> }) => Promise<unknown>) => callback({ sync: vi.fn() })),
    };

    const { getWorkbookContent } = await import("./getWorkbookContent");
    const result = await getWorkbookContent.handler({ sheetName: "Budget" });

    expect(result).toBe("Worksheet: Budget\nRange: (empty used range)\n\n(empty range)");
  });
});
