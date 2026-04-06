import { afterEach, describe, expect, it, vi } from "vitest";
import { setSelectedRange } from "./setSelectedRange";

function loadable<T extends object>(value: T): T & { load: ReturnType<typeof vi.fn> } {
  return Object.assign(value, { load: vi.fn() });
}

describe("setSelectedRange", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("rejects dimension mismatches for multi-cell selections", async () => {
    const worksheet = loadable({ name: "Sheet1" });
    const selectedRange = loadable({
      address: "Sheet1!A1:B2",
      rowCount: 2,
      columnCount: 2,
      worksheet,
      getResizedRange: vi.fn(),
    });
    const context = {
      workbook: {
        getSelectedRange: vi.fn(() => selectedRange),
      },
      sync: vi.fn(),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const result = await setSelectedRange.handler({ data: [[1, 2, 3]] });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("Data dimensions (1x3) do not match selection dimensions (2x2)"),
    });
  });

  it("expands a single-cell selection and writes formulas when requested", async () => {
    const worksheet = loadable({ name: "Sheet1" });
    const targetRange = { formulas: undefined as unknown, values: undefined as unknown };
    const selectedRange = loadable({
      address: "Sheet1!A1",
      rowCount: 1,
      columnCount: 1,
      worksheet,
      getResizedRange: vi.fn(() => targetRange),
    });
    const context = {
      workbook: {
        getSelectedRange: vi.fn(() => selectedRange),
      },
      sync: vi.fn(),
    };

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const data = [["=SUM(A1:A2)", 3]];
    const result = await setSelectedRange.handler({ data, useFormulas: true });

    expect(selectedRange.getResizedRange).toHaveBeenCalledWith(0, 1);
    expect(targetRange.formulas).toEqual(data);
    expect(targetRange.values).toBeUndefined();
    expect(result).toBe("Successfully wrote 1 rows and 2 columns to the selected range in Sheet1.");
  });
});
