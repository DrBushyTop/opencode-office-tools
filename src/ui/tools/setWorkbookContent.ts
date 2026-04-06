import { z } from "zod";
import type { Tool } from "./types";
import { countDataColumns, excel2DDataSchema, getWorksheet, parseToolArgs, toolFailure, writeExcelData } from "./excelShared";

const setWorkbookContentArgsSchema = z.object({
  sheetName: z.string().optional(),
  startCell: z.string(),
  data: excel2DDataSchema,
  useFormulas: z.boolean().default(false),
  clearMode: z.enum(["none", "contents", "all"]).default("none"),
  createTable: z.boolean().default(false),
  tableName: z.string().optional(),
  hasHeaders: z.boolean().default(true),
  tableStyle: z.string().optional(),
});

function clearTargetRange(range: Excel.Range, clearMode: "none" | "contents" | "all") {
  if (clearMode === "contents") {
    range.clear(Excel.ClearApplyTo.contents);
  }
  if (clearMode === "all") {
    range.clear(Excel.ClearApplyTo.all);
  }
}

async function createTableIfRequested(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  targetRange: Excel.Range,
  options: { createTable: boolean; tableName?: string; hasHeaders: boolean; tableStyle?: string },
) {
  if (!options.createTable) {
    await context.sync();
    return "";
  }

  const table = worksheet.tables.add(targetRange, options.hasHeaders);
  if (options.tableName) table.name = options.tableName;
  if (options.tableStyle) table.style = options.tableStyle;
  table.load("name");
  await context.sync();
  return ` Created table ${table.name}.`;
}

export const setWorkbookContent: Tool = {
  name: "set_workbook_content",
  description: "Write a rectangular block of Excel data beginning at a chosen start cell. The write can clear the target first, preserve formulas, and optionally promote the block into a table.",
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
    const parsedArgs = parseToolArgs(setWorkbookContentArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const { sheetName, startCell, data, useFormulas, clearMode, createTable, tableName, hasHeaders, tableStyle } = parsedArgs.data;
    const columnCount = countDataColumns(data);

    try {
      return await Excel.run(async (context) => {
        const worksheet = await getWorksheet(context, sheetName);
        const rowCount = data.length;
        const startRange = worksheet.getRange(startCell);
        const targetRange = startRange.getResizedRange(rowCount - 1, columnCount - 1);
        targetRange.load("address");
        await context.sync();

        clearTargetRange(targetRange, clearMode);
        writeExcelData(targetRange, data, useFormulas);

        const tableResult = await createTableIfRequested(context, worksheet, targetRange, {
          createTable,
          tableName,
          hasHeaders,
          tableStyle,
        });

        return `Wrote ${rowCount} rows and ${columnCount} columns to ${targetRange.address} in ${worksheet.name}.${tableResult}`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
