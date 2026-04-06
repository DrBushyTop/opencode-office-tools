import { z } from "zod";
import type { Tool } from "./types";
import { countDataColumns, excel2DDataSchema, parseToolArgs, toolFailure, writeExcelData } from "./excelShared";

const setSelectedRangeArgsSchema = z.object({
  data: excel2DDataSchema,
  useFormulas: z.boolean().default(true),
});

function resolveWriteTarget(selectedRange: Excel.Range, dataRowCount: number, dataColCount: number) {
  const canExpandSelection = selectedRange.rowCount === 1 && selectedRange.columnCount === 1;
  if (canExpandSelection) {
    return { ok: true as const, range: selectedRange.getResizedRange(dataRowCount - 1, dataColCount - 1) };
  }

  const hasExactMatch = dataRowCount === selectedRange.rowCount && dataColCount === selectedRange.columnCount;
  if (hasExactMatch) {
    return { ok: true as const, range: selectedRange };
  }

  return {
    ok: false as const,
    failure: toolFailure(`Data dimensions (${dataRowCount}x${dataColCount}) do not match selection dimensions (${selectedRange.rowCount}x${selectedRange.columnCount}). Either select a single cell to auto-expand, or provide data matching the selection size.`),
  };
}

export const setSelectedRange: Tool = {
  name: "set_selected_range",
  description: "Fill the active Excel selection with a rectangular 2D array. Single-cell selections automatically expand to the incoming data size.",
  parameters: {
    type: "object",
    properties: {
      data: {
        type: "array",
        description: "Rectangular 2D array of values to write to the selected range.",
        items: {
          type: "array",
          items: {
            type: ["string", "number", "boolean"],
          },
        },
      },
      useFormulas: {
        type: "boolean",
        description: "If true, treat string values starting with '=' as formulas. Default is true.",
      },
    },
    required: ["data"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(setSelectedRangeArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const { data, useFormulas } = parsedArgs.data;
    const columnCount = countDataColumns(data);

    try {
      return await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(["address", "rowCount", "columnCount"]);

        const worksheet = selectedRange.worksheet;
        worksheet.load("name");

        await context.sync();

        const dataRowCount = data.length;
        const dataColCount = columnCount;

        const target = resolveWriteTarget(selectedRange, dataRowCount, dataColCount);
        if (!target.ok) return target.failure;

        writeExcelData(target.range, data, useFormulas);

        await context.sync();

        return `Successfully wrote ${dataRowCount} rows and ${dataColCount} columns to the selected range in ${worksheet.name}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
