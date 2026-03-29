import { z } from "zod";
import type { Tool } from "./types";
import { excelCellValueSchema, getWorksheet, isExcelRequirementSetSupported, parseToolArgs, toolFailure } from "./excelShared";

const tableRowValuesSchema = z.array(z.array(excelCellValueSchema)).superRefine((value, context) => {
  if (value.length <= 1) return;
  const columnCount = value[0].length;
  if (!value.every((row) => row.length === columnCount)) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "All rows in data must have the same length.",
      path: ["values"],
    });
  }
});

const manageTableArgsSchema = z.object({
  action: z.enum(["create", "rename", "resize", "setProperties", "addRows", "clearFilters", "reapplyFilters", "convertToRange", "delete"]),
  tableName: z.string().optional(),
  sheetName: z.string().optional(),
  sourceRange: z.string().optional(),
  hasHeaders: z.boolean().default(true),
  newName: z.string().optional(),
  style: z.string().optional(),
  showHeaders: z.boolean().optional(),
  showTotals: z.boolean().optional(),
  showBandedRows: z.boolean().optional(),
  showBandedColumns: z.boolean().optional(),
  showFilterButton: z.boolean().optional(),
  highlightFirstColumn: z.boolean().optional(),
  highlightLastColumn: z.boolean().optional(),
  values: tableRowValuesSchema.optional(),
  rowIndex: z.number().refine((value) => Number.isInteger(value) && value >= -1, "rowIndex must be an integer greater than or equal to -1.").optional(),
  alwaysInsert: z.boolean().default(true),
});

export const manageTable: Tool = {
  name: "manage_table",
  description: "Create or update Excel tables. Supports creation, rename, resize, style changes, header or totals visibility, row appends/inserts, filter reset, conversion back to range, and deletion.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "rename", "resize", "setProperties", "addRows", "clearFilters", "reapplyFilters", "convertToRange", "delete"],
        description: "Table operation to perform.",
      },
      tableName: {
        type: "string",
        description: "Existing table name or id. Required for all actions except create.",
      },
      sheetName: {
        type: "string",
        description: "Worksheet name for create when the range is sheet-local.",
      },
      sourceRange: {
        type: "string",
        description: "Source range for create or resize, such as 'A1:D20'.",
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether the source range has headers when creating a table. Default true.",
      },
      newName: {
        type: "string",
        description: "New table name for rename or create.",
      },
      style: {
        type: "string",
        description: "Excel table style to apply.",
      },
      showHeaders: { type: "boolean" },
      showTotals: { type: "boolean" },
      showBandedRows: { type: "boolean" },
      showBandedColumns: { type: "boolean" },
      showFilterButton: { type: "boolean" },
      highlightFirstColumn: { type: "boolean" },
      highlightLastColumn: { type: "boolean" },
      values: {
        type: "array",
        items: {
          type: "array",
          items: { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }] },
        },
        description: "Row data to append or insert for addRows.",
      },
      rowIndex: {
        type: "number",
        description: "Zero-based row index for addRows. Omit or use -1 to append.",
      },
      alwaysInsert: {
        type: "boolean",
        description: "Whether addRows should insert into the table when adding rows. Default true.",
      },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(manageTableArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const {
      action,
      tableName,
      sheetName,
      sourceRange,
      hasHeaders,
      newName,
      style,
      showHeaders,
      showTotals,
      showBandedRows,
      showBandedColumns,
      showFilterButton,
      highlightFirstColumn,
      highlightLastColumn,
      values,
      rowIndex,
      alwaysInsert,
    } = parsedArgs.data;

    if ((action === "create" && hasHeaders === false && showFilterButton) || (action === "setProperties" && showHeaders === false && showFilterButton)) {
      return toolFailure("showFilterButton can only be enabled when the table shows headers.");
    }

    try {
      return await Excel.run(async (context) => {
        if (action === "create") {
          if (!sourceRange) return toolFailure("sourceRange is required for create.");
          const sheet = await getWorksheet(context, sheetName);
          const range = sheet.getRange(sourceRange);
          const table = sheet.tables.add(range, hasHeaders);
          if (newName) table.name = newName;
          if (style) table.style = style;
          if (showHeaders !== undefined) table.showHeaders = showHeaders;
          if (showTotals !== undefined) table.showTotals = showTotals;
          if (showBandedRows !== undefined) table.showBandedRows = showBandedRows;
          if (showBandedColumns !== undefined) table.showBandedColumns = showBandedColumns;
          if (showFilterButton !== undefined) table.showFilterButton = showFilterButton;
          if (highlightFirstColumn !== undefined) table.highlightFirstColumn = highlightFirstColumn;
          if (highlightLastColumn !== undefined) table.highlightLastColumn = highlightLastColumn;
          table.load(["name", "style"]);
          await context.sync();
          return `Created table ${table.name} from ${sourceRange} on ${sheet.name}${style ? ` with style ${table.style}` : ""}.`;
        }

        if (!tableName) return toolFailure("tableName is required for this action.");
        const table = context.workbook.tables.getItemOrNullObject(tableName);
        table.load(["isNullObject", "name", "style"]);
        await context.sync();
        if ((table as Excel.Table & { isNullObject?: boolean }).isNullObject) {
          return toolFailure(`Table ${tableName} was not found.`);
        }

        switch (action) {
          case "rename":
            if (!newName) return toolFailure("newName is required for rename.");
            table.name = newName;
            await context.sync();
            return `Renamed table ${tableName} to ${newName}.`;
          case "resize":
            if (!sourceRange) return toolFailure("sourceRange is required for resize.");
            if (!isExcelRequirementSetSupported("1.13")) {
              return toolFailure("Resizing tables requires ExcelApi 1.13.");
            }
            table.resize(sourceRange);
            await context.sync();
            return `Resized table ${table.name} to ${sourceRange}.`;
          case "setProperties":
            if (style !== undefined) table.style = style;
            if (showHeaders !== undefined) table.showHeaders = showHeaders;
            if (showTotals !== undefined) table.showTotals = showTotals;
            if (showBandedRows !== undefined) table.showBandedRows = showBandedRows;
            if (showBandedColumns !== undefined) table.showBandedColumns = showBandedColumns;
            if (showFilterButton !== undefined) table.showFilterButton = showFilterButton;
            if (highlightFirstColumn !== undefined) table.highlightFirstColumn = highlightFirstColumn;
            if (highlightLastColumn !== undefined) table.highlightLastColumn = highlightLastColumn;
            await context.sync();
            return `Updated table ${table.name}.`;
          case "addRows": {
            if (!values?.length) return toolFailure("values is required for addRows.");
            if (alwaysInsert !== true && !isExcelRequirementSetSupported("1.15")) {
              return toolFailure("alwaysInsert=false for addRows requires ExcelApi 1.15.");
            }
            if (isExcelRequirementSetSupported("1.15")) {
              table.rows.add(rowIndex ?? -1, values, alwaysInsert);
            } else {
              table.rows.add(rowIndex ?? -1, values);
            }
            await context.sync();
            return `Added ${values.length} row(s) to table ${table.name}${rowIndex !== undefined && rowIndex >= 0 ? ` at index ${rowIndex}` : ""}.`;
          }
          case "clearFilters":
            table.clearFilters();
            await context.sync();
            return `Cleared filters on table ${table.name}.`;
          case "reapplyFilters":
            table.reapplyFilters();
            await context.sync();
            return `Reapplied filters on table ${table.name}.`;
          case "convertToRange": {
            const range = table.convertToRange();
            range.load("address");
            await context.sync();
            return `Converted table ${table.name} to range ${range.address}.`;
          }
          case "delete":
            table.delete();
            await context.sync();
            return `Deleted table ${table.name}.`;
          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
