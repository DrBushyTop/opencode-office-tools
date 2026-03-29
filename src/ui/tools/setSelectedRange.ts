import { z } from "zod";
import type { Tool } from "./types";
import { excel2DDataSchema, parseToolArgs, toolFailure } from "./excelShared";

const setSelectedRangeArgsSchema = z.object({
  data: excel2DDataSchema,
  useFormulas: z.boolean().default(true),
});

export const setSelectedRange: Tool = {
  name: "set_selected_range",
  description: "Write values or formulas to the currently selected range in Excel. A single-cell selection expands to fit the provided rectangular 2D array.",
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
    const columnCount = data[0].length;

    try {
      return await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(["address", "rowCount", "columnCount"]);

        const worksheet = selectedRange.worksheet;
        worksheet.load("name");

        await context.sync();

        const dataRowCount = data.length;
        const dataColCount = columnCount;

        let targetRange: Excel.Range;
        if (selectedRange.rowCount === 1 && selectedRange.columnCount === 1) {
          targetRange = selectedRange.getResizedRange(dataRowCount - 1, dataColCount - 1);
        } else {
          if (dataRowCount !== selectedRange.rowCount || dataColCount !== selectedRange.columnCount) {
            return toolFailure(`Data dimensions (${dataRowCount}x${dataColCount}) do not match selection dimensions (${selectedRange.rowCount}x${selectedRange.columnCount}). Either select a single cell to auto-expand, or provide data matching the selection size.`);
          }
          targetRange = selectedRange;
        }

        if (useFormulas) {
          const hasFormulas = data.some((row) => row.some((cell) => typeof cell === "string" && cell.startsWith("=")));
          if (hasFormulas) {
            targetRange.formulas = data;
          } else {
            targetRange.values = data;
          }
        } else {
          targetRange.values = data;
        }

        await context.sync();

        return `Successfully wrote ${dataRowCount} rows and ${dataColCount} columns to the selected range in ${worksheet.name}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
