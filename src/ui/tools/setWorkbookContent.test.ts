import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("./excelShared", async () => {
  const actual = await vi.importActual<any>("./excelShared");
  return {
    ...actual,
    getWorksheet: vi.fn(),
  };
});

function loadable<T extends object>(value: T): T & { load: ReturnType<typeof vi.fn> } {
  return Object.assign(value, { load: vi.fn() });
}

describe("setWorkbookContent", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.resetModules();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("clears, writes, and creates a table over the destination range", async () => {
    const excelShared = await import("./excelShared");
    const targetRange = loadable({ address: "Budget!B2:C3", clear: vi.fn(), formulas: undefined as unknown, values: undefined as unknown });
    const startRange = {
      getResizedRange: vi.fn(() => targetRange),
    };
    const table = loadable({ name: "SalesTable", style: "" });
    const worksheet = {
      name: "Budget",
      getRange: vi.fn(() => startRange),
      tables: { add: vi.fn(() => table) },
    };
    vi.mocked(excelShared.getWorksheet).mockResolvedValue(worksheet as never);

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: { sync: ReturnType<typeof vi.fn> }) => Promise<unknown>) => callback({ sync: vi.fn() })),
      ClearApplyTo: { contents: "contents", all: "all" },
    };

    const { setWorkbookContent } = await import("./setWorkbookContent");
    const result = await setWorkbookContent.handler({
      sheetName: "Budget",
      startCell: "B2",
      data: [["Region", "Sales"], ["North", 42]],
      clearMode: "contents",
      createTable: true,
      tableName: "SalesTable",
      hasHeaders: true,
      tableStyle: "TableStyleMedium2",
    });

    expect(worksheet.getRange).toHaveBeenCalledWith("B2");
    expect(startRange.getResizedRange).toHaveBeenCalledWith(1, 1);
    expect(targetRange.clear).toHaveBeenCalledWith("contents");
    expect(targetRange.values).toEqual([["Region", "Sales"], ["North", 42]]);
    expect(worksheet.tables.add).toHaveBeenCalledWith(targetRange, true);
    expect(table.name).toBe("SalesTable");
    expect(table.style).toBe("TableStyleMedium2");
    expect(result).toBe("Updated Budget!B2:C3 on Budget with 2 rows and 2 columns. Promoted the written range to table SalesTable.");
  });
});
