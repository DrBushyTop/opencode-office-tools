import { afterEach, describe, expect, it, vi } from "vitest";
import { insertTable } from "./insertTable";

function loadable<T extends object>(value: T): T & { load: ReturnType<typeof vi.fn> } {
  return Object.assign(value, { load: vi.fn() });
}

describe("insertTable", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { Word?: unknown }).Word;
  });

  it("rejects tables without any columns", async () => {
    const result = await insertTable.handler({ data: [[]] });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Table must have at least one column.",
    });
  });

  it("inserts a table and applies header and striped row styling", async () => {
    const headerParagraph = loadable({ font: { bold: false, color: "#000000" } });
    const makeCell = () => ({
      shadingColor: "",
      body: loadable({ paragraphs: loadable({ items: [headerParagraph] }) }),
    });
    const headerRow = loadable({ cells: loadable({ items: [makeCell(), makeCell()] }) });
    const firstDataRow = loadable({ cells: loadable({ items: [makeCell(), makeCell()] }) });
    const secondDataRow = loadable({ cells: loadable({ items: [makeCell(), makeCell()] }) });
    const table = loadable({
      rows: loadable({ items: [headerRow, firstDataRow, secondDataRow] }),
      style: "",
    });
    const selection = {
      insertTable: vi.fn(() => table),
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
      },
      sync: vi.fn(),
    };

    (globalThis as { Word?: unknown }).Word = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
      InsertLocation: { after: "after" },
    };

    const result = await insertTable.handler({
      data: [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]],
      hasHeader: true,
      style: "striped",
    });

    expect(selection.insertTable).toHaveBeenCalledWith(3, 2, "after", [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    expect(table.style).toBe("TableStyleLight3");
    for (const cell of headerRow.cells.items) {
      expect(cell.shadingColor).toBe("#4472C4");
      expect(cell.body.paragraphs.items[0].font).toMatchObject({ bold: true, color: "#FFFFFF" });
    }
    for (const cell of firstDataRow.cells.items) {
      expect(cell.shadingColor).toBe("#E8E8E8");
    }
    expect(result).toBe("Inserted 3x2 table with striped style and header row.");
  });
});
