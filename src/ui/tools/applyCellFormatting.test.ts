import { afterEach, describe, expect, it, vi } from "vitest";

function loadable<T extends object>(value: T): T & { load: ReturnType<typeof vi.fn> } {
  return Object.assign(value, { load: vi.fn() });
}

describe("applyCellFormatting", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("applies unmerge before sizing operations", async () => {
    const operations: string[] = [];
    const font = {
      set bold(value: boolean) {
        operations.push(`bold:${String(value)}`);
      },
      set italic(value: boolean) {
        operations.push(`italic:${String(value)}`);
      },
      set underline(value: unknown) {
        operations.push(`underline:${String(value)}`);
      },
      set size(value: number) {
        operations.push(`fontSize:${String(value)}`);
      },
      set color(value: string) {
        operations.push(`fontColor:${value}`);
      },
    };
    const fill = {
      set color(value: string) {
        operations.push(`fill:${value}`);
      },
    };
    const borders = {
      getItem: vi.fn(() => ({
        set style(value: unknown) {
          operations.push(`borderStyle:${String(value)}`);
        },
        set color(value: string) {
          operations.push(`borderColor:${value}`);
        },
        set weight(value: unknown) {
          operations.push(`borderWeight:${String(value)}`);
        },
      })),
    };
    const format = {
      font,
      fill,
      borders,
      set horizontalAlignment(value: unknown) {
        operations.push(`horizontal:${String(value)}`);
      },
      set verticalAlignment(value: unknown) {
        operations.push(`vertical:${String(value)}`);
      },
      set wrapText(value: boolean) {
        operations.push(`wrap:${String(value)}`);
      },
      set rowHeight(value: number) {
        operations.push(`rowHeight:${String(value)}`);
      },
      set columnWidth(value: number) {
        operations.push(`columnWidth:${String(value)}`);
      },
      autofitRows: vi.fn(() => {
        operations.push("autoFitRows");
      }),
      autofitColumns: vi.fn(() => {
        operations.push("autoFitColumns");
      }),
    };
    const tableCollection = loadable({ items: [] as Array<{ name: string }> });
    const range = loadable({
      address: "Sheet1!A1:F7",
      rowCount: 7,
      columnCount: 6,
      format,
      getTables: vi.fn(() => tableCollection),
      set numberFormat(value: string[][]) {
        operations.push(`numberFormat:${value[0]?.[0] ?? ""}`);
      },
      merge: vi.fn((across?: boolean) => {
        operations.push(`merge:${String(Boolean(across))}`);
      }),
      unmerge: vi.fn(() => {
        operations.push("unmerge");
      }),
    });
    const sheet = loadable({
      name: "Sheet1",
      getRange: vi.fn(() => range),
    });
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: vi.fn(() => sheet),
        },
      },
      sync: vi.fn(async () => undefined),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
      RangeUnderlineStyle: { single: "Single", none: "None" },
      HorizontalAlignment: {
        left: "Left",
        center: "Center",
        right: "Right",
        general: "General",
        fill: "Fill",
        justify: "Justify",
        centerAcrossSelection: "CenterAcrossSelection",
        distributed: "Distributed",
      },
      VerticalAlignment: {
        top: "Top",
        center: "Center",
        bottom: "Bottom",
        justify: "Justify",
        distributed: "Distributed",
      },
      BorderLineStyle: {
        continuous: "Continuous",
        none: "None",
        double: "Double",
        dash: "Dash",
        dot: "Dot",
      },
      BorderWeight: {
        thin: "Thin",
        medium: "Medium",
        thick: "Thick",
      },
      BorderIndex: {
        edgeTop: "EdgeTop",
        edgeBottom: "EdgeBottom",
        edgeLeft: "EdgeLeft",
        edgeRight: "EdgeRight",
        insideHorizontal: "InsideHorizontal",
        insideVertical: "InsideVertical",
      },
    };

    const { applyCellFormatting } = await import("./applyCellFormatting");
    const result = await applyCellFormatting.handler({
      range: "A1:F7",
      sheetName: "Sheet1",
      bold: false,
      italic: false,
      underline: false,
      fontSize: 11,
      fontColor: "#000000",
      backgroundColor: "#FFFFFF",
      numberFormat: "@",
      horizontalAlignment: "left",
      verticalAlignment: "center",
      wrapText: true,
      merge: false,
      mergeAcross: false,
      borderStyle: "thin",
      borderColor: "#D9D9D9",
      interiorBorders: true,
      rowHeight: 22,
      columnWidth: 22,
      autoFitRows: true,
      autoFitColumns: true,
    });

    expect(typeof result).toBe("string");
    expect(operations.indexOf("unmerge")).toBeGreaterThanOrEqual(0);
    expect(operations.indexOf("rowHeight:22")).toBeGreaterThan(operations.indexOf("unmerge"));
    expect(operations.indexOf("columnWidth:22")).toBeGreaterThan(operations.indexOf("unmerge"));
    expect(operations.indexOf("autoFitRows")).toBeGreaterThan(operations.indexOf("unmerge"));
    expect(operations.indexOf("autoFitColumns")).toBeGreaterThan(operations.indexOf("unmerge"));
  });

  it("includes Excel debug information in failures", async () => {
    const excelError = Object.assign(new Error("The requested operation is invalid."), {
      code: "InvalidOperation",
      debugInfo: {
        message: "Cannot change part of a merged cell.",
        errorLocation: "RangeFormat.rowHeight",
      },
    });

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async () => {
        throw excelError;
      }),
    };

    const { applyCellFormatting } = await import("./applyCellFormatting");
    const result = await applyCellFormatting.handler({ range: "A1", rowHeight: 22 });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "The requested operation is invalid.: Cannot change part of a merged cell. (at RangeFormat.rowHeight) [InvalidOperation]",
    });
  });

  it("skips unmerge for ranges that overlap an Excel table", async () => {
    const operations: string[] = [];
    const format = {
      font: {
        bold: false,
        italic: false,
        underline: "None",
        size: 11,
        color: "#000000",
      },
      fill: {
        color: "#FFFFFF",
      },
      borders: {
        getItem: vi.fn(() => ({ style: "", color: "", weight: "" })),
      },
      horizontalAlignment: "Left",
      verticalAlignment: "Center",
      wrapText: false,
      rowHeight: 0,
      columnWidth: 0,
      autofitRows: vi.fn(() => {
        operations.push("autoFitRows");
      }),
      autofitColumns: vi.fn(() => {
        operations.push("autoFitColumns");
      }),
    };
    const tableCollection = loadable({ items: [{ name: "SurveyTable" }] });
    const range = loadable({
      address: "Sheet1!A1:F7",
      rowCount: 7,
      columnCount: 6,
      format,
      getTables: vi.fn(() => tableCollection),
      set numberFormat(_value: string[][]) {},
      merge: vi.fn(() => {
        operations.push("merge");
      }),
      unmerge: vi.fn(() => {
        operations.push("unmerge");
      }),
    });
    const sheet = loadable({
      name: "Sheet1",
      getRange: vi.fn(() => range),
    });
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: vi.fn(() => sheet),
        },
      },
      sync: vi.fn(async () => undefined),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
      RangeUnderlineStyle: { single: "Single", none: "None" },
      HorizontalAlignment: {
        left: "Left",
        center: "Center",
        right: "Right",
        general: "General",
        fill: "Fill",
        justify: "Justify",
        centerAcrossSelection: "CenterAcrossSelection",
        distributed: "Distributed",
      },
      VerticalAlignment: {
        top: "Top",
        center: "Center",
        bottom: "Bottom",
        justify: "Justify",
        distributed: "Distributed",
      },
      BorderLineStyle: {
        continuous: "Continuous",
        none: "None",
        double: "Double",
        dash: "Dash",
        dot: "Dot",
      },
      BorderWeight: {
        thin: "Thin",
        medium: "Medium",
        thick: "Thick",
      },
      BorderIndex: {
        edgeTop: "EdgeTop",
        edgeBottom: "EdgeBottom",
        edgeLeft: "EdgeLeft",
        edgeRight: "EdgeRight",
        insideHorizontal: "InsideHorizontal",
        insideVertical: "InsideVertical",
      },
    };

    const { applyCellFormatting } = await import("./applyCellFormatting");
    const result = await applyCellFormatting.handler({ range: "A1:F7", sheetName: "Sheet1", merge: false, rowHeight: 22 });

    expect(typeof result).toBe("string");
    expect(operations).not.toContain("unmerge");
    expect(result).toContain("merge unchanged (table cells cannot be merged or unmerged)");
  });

  it("returns a clear failure when merge is requested for a table range", async () => {
    const tableCollection = loadable({ items: [{ name: "SurveyTable" }] });
    const range = loadable({
      address: "Sheet1!A1:F7",
      rowCount: 7,
      columnCount: 6,
      format: {
        font: {},
        fill: {},
        borders: { getItem: vi.fn() },
      },
      getTables: vi.fn(() => tableCollection),
    });
    const sheet = loadable({
      name: "Sheet1",
      getRange: vi.fn(() => range),
    });
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: vi.fn(() => sheet),
        },
      },
      sync: vi.fn(async () => undefined),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const { applyCellFormatting } = await import("./applyCellFormatting");
    const result = await applyCellFormatting.handler({ range: "A1:F7", sheetName: "Sheet1", merge: true });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Cannot merge Sheet1!A1:F7 because it overlaps an Excel table. Convert the table to a normal range first or omit merge.",
    });
  });
});
