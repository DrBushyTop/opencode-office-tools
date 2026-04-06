import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("./excelShared", async () => {
  const actual = await vi.importActual<any>("./excelShared");
  return {
    ...actual,
    isExcelRequirementSetSupported: vi.fn(() => true),
  };
});

function loadable<T extends object>(value: T): T & { load: ReturnType<typeof vi.fn> } {
  return Object.assign(value, { load: vi.fn() });
}

describe("getWorkbookOverview", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("summarizes workbook structure and totals", async () => {
    const activeSheet = loadable({ name: "Sheet1" });
    const usedRange = loadable({ isNullObject: false, address: "Sheet1!A1:B2", rowCount: 2, columnCount: 2 });
    const tableCollection = loadable({ items: [{ name: "Sales", style: "TableStyleMedium2", showTotals: false }] });
    const worksheetNames = loadable({ items: [{ name: "LocalTotal" }] });
    const protection = loadable({ protected: false });
    const freezeLocation = loadable({ isNullObject: true, address: "" });
    const autoFilter = loadable({ enabled: false, isDataFiltered: false });
    const sheet = loadable({
      id: "sheet-1",
      name: "Sheet1",
      position: 0,
      visibility: "Visible",
      getUsedRangeOrNullObject: vi.fn(() => usedRange),
      charts: { getCount: vi.fn(() => ({ value: 1 })) },
      tables: Object.assign(tableCollection, { getCount: vi.fn(() => ({ value: 1 })) }),
      pivotTables: { getCount: vi.fn(() => ({ value: 0 })) },
      names: worksheetNames,
      protection,
      freezePanes: { getLocationOrNullObject: vi.fn(() => freezeLocation) },
      autoFilter,
    });
    const workbookNames = loadable({ items: [loadable({ name: "GrandTotal", value: "=Sheet1!$B$2", scope: "Workbook" })] });
    const worksheets = loadable({
      items: [sheet],
      getActiveWorksheet: vi.fn(() => activeSheet),
    });
    const context = {
      workbook: {
        worksheets,
        names: workbookNames,
      },
      sync: vi.fn(),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const { getWorkbookOverview } = await import("./getWorkbookOverview");
    const result = await getWorkbookOverview.handler();

    expect(result).toContain("Workbook inventory");
    expect(result).toContain("Sheet count: 1");
    expect(result).toContain("Active sheet: Sheet1");
    expect(result).toContain("footprint: Sheet1!A1:B2 (2 rows x 2 cols)");
    expect(result).toContain("objects: charts=1, tables=1, pivotTables=0");
    expect(result).toContain("table catalog: Sales (TableStyleMedium2)");
    expect(result).toContain("local names: LocalTotal");
    expect(result).toContain("Workbook totals:");
    expect(result).toContain("Workbook names (1):");
    expect(result).toContain("- GrandTotal [Workbook]: =Sheet1!$B$2");
  });
});
