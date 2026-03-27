import type { Tool } from "./types";
import { getWorksheet, toolFailure } from "./excelShared";

export const setWorkbookContent: Tool = {
  name: "set_workbook_content",
  description: "Write a 2D array to a worksheet starting at a specific cell. Can write formulas, clear the destination first, and optionally turn the written range into a table.",
  parameters: {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: "Optional worksheet name. Defaults to the active sheet.",
      },
      startCell: {
        type: "string",
        description: "Starting cell address such as 'A1' or 'B5'.",
      },
      data: {
        type: "array",
        description: "2D array of values to write.",
        items: {
          type: "array",
          items: {
            type: ["string", "number", "boolean"],
          },
        },
      },
      useFormulas: {
        type: "boolean",
        description: "If true, string values starting with '=' are written as formulas. Default false.",
      },
      clearMode: {
        type: "string",
        enum: ["none", "contents", "all"],
        description: "Optionally clear the target range before writing. Default none.",
      },
      createTable: {
        type: "boolean",
        description: "Create an Excel table over the written range after writing. Default false.",
      },
      tableName: {
        type: "string",
        description: "Optional table name to assign when createTable is true.",
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether the written range includes a header row when createTable is true. Default true.",
      },
      tableStyle: {
        type: "string",
        description: "Optional Excel table style to apply when createTable is true.",
      },
    },
    required: ["startCell", "data"],
  },
  handler: async (args) => {
    const {
      sheetName,
      startCell,
      data,
      useFormulas = false,
      clearMode = "none",
      createTable = false,
      tableName,
      hasHeaders = true,
      tableStyle,
    } = args as {
      sheetName?: string;
      startCell: string;
      data: Array<Array<string | number | boolean>>;
      useFormulas?: boolean;
      clearMode?: "none" | "contents" | "all";
      createTable?: boolean;
      tableName?: string;
      hasHeaders?: boolean;
      tableStyle?: string;
    };

    if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0]) || data[0].length === 0) {
      return toolFailure("Provide a non-empty 2D data array.");
    }

    const columnCount = data[0].length;
    if (!data.every((row) => Array.isArray(row) && row.length === columnCount)) {
      return toolFailure("All rows in data must have the same length.");
    }

    try {
      return await Excel.run(async (context) => {
        const worksheet = await getWorksheet(context, sheetName);
        const rowCount = data.length;
        const startRange = worksheet.getRange(startCell);
        const targetRange = startRange.getResizedRange(rowCount - 1, columnCount - 1);
        targetRange.load("address");
        await context.sync();

        if (clearMode === "contents") {
          targetRange.clear(Excel.ClearApplyTo.contents);
        } else if (clearMode === "all") {
          targetRange.clear(Excel.ClearApplyTo.all);
        }

        const hasFormulaStrings = useFormulas && data.some((row) => row.some((cell) => typeof cell === "string" && cell.startsWith("=")));
        if (hasFormulaStrings) {
          targetRange.formulas = data;
        } else {
          targetRange.values = data;
        }

        let tableResult = "";
        if (createTable) {
          const table = worksheet.tables.add(targetRange, hasHeaders);
          if (tableName) table.name = tableName;
          if (tableStyle) table.style = tableStyle;
          table.load("name");
          await context.sync();
          tableResult = ` Created table ${table.name}.`;
        } else {
          await context.sync();
        }

        return `Wrote ${rowCount} rows and ${columnCount} columns to ${targetRange.address} in ${worksheet.name}.${tableResult}`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
