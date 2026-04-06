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

type WorkbookWritePlan = {
  rowCount: number;
  columnCount: number;
  clearMode: "none" | "contents" | "all";
  createTable: boolean;
  tableName?: string;
  hasHeaders: boolean;
  tableStyle?: string;
};

function createWritePlan(args: z.infer<typeof setWorkbookContentArgsSchema>): WorkbookWritePlan {
  return {
    rowCount: args.data.length,
    columnCount: countDataColumns(args.data),
    clearMode: args.clearMode,
    createTable: args.createTable,
    tableName: args.tableName,
    hasHeaders: args.hasHeaders,
    tableStyle: args.tableStyle,
  };
}

function applyClearMode(range: Excel.Range, clearMode: WorkbookWritePlan["clearMode"]) {
  if (clearMode === "contents") {
    range.clear(Excel.ClearApplyTo.contents);
  } else if (clearMode === "all") {
    range.clear(Excel.ClearApplyTo.all);
  }
}

async function finalizeTablePromotion(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  targetRange: Excel.Range,
  plan: WorkbookWritePlan,
) {
  if (!plan.createTable) {
    await context.sync();
    return "";
  }

  const table = worksheet.tables.add(targetRange, plan.hasHeaders);
  if (plan.tableName) table.name = plan.tableName;
  if (plan.tableStyle) table.style = plan.tableStyle;
  table.load("name");
  await context.sync();
  return ` Table promotion complete: ${table.name}.`;
}

function formatWriteSummary(worksheetName: string, rangeAddress: string, plan: WorkbookWritePlan, tableSummary: string) {
  return `Write plan applied to ${worksheetName} at ${rangeAddress}: ${plan.rowCount} row(s) x ${plan.columnCount} column(s).${tableSummary}`;
}

export const setWorkbookContent: Tool = {
  name: "set_workbook_content",
  description: "Build a workbook write plan from a start cell and a 2D data block, then apply clearing, writing, and optional table promotion in one Excel operation.",
  parameters: {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: "Optional worksheet name. Defaults to the active sheet.",
      },
      startCell: {
        type: "string",
        description: "Top-left cell for the write plan, such as 'A1' or 'B5'.",
      },
      data: {
        type: "array",
        description: "Rectangular 2D array of values to write.",
        items: {
          type: "array",
          items: {
            type: ["string", "number", "boolean"],
          },
        },
      },
      useFormulas: {
        type: "boolean",
        description: "Treat strings beginning with '=' as formulas when true.",
      },
      clearMode: {
        type: "string",
        enum: ["none", "contents", "all"],
        description: "Optional clear behavior applied before writing.",
      },
      createTable: {
        type: "boolean",
        description: "Promote the written range into an Excel table after writing.",
      },
      tableName: {
        type: "string",
        description: "Optional table name used when createTable is true.",
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether the written range already includes a header row for table promotion.",
      },
      tableStyle: {
        type: "string",
        description: "Optional Excel table style to apply during table promotion.",
      },
    },
    required: ["startCell", "data"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(setWorkbookContentArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const plan = createWritePlan(parsedArgs.data);

    try {
      return await Excel.run(async (context) => {
        const worksheet = await getWorksheet(context, parsedArgs.data.sheetName);
        const anchorRange = worksheet.getRange(parsedArgs.data.startCell);
        const writeRange = anchorRange.getResizedRange(plan.rowCount - 1, plan.columnCount - 1);
        writeRange.load("address");
        await context.sync();

        applyClearMode(writeRange, plan.clearMode);
        writeExcelData(writeRange, parsedArgs.data.data, parsedArgs.data.useFormulas);
        const tableSummary = await finalizeTablePromotion(context, worksheet, writeRange, plan);
        return formatWriteSummary(worksheet.name, writeRange.address, plan, tableSummary);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
