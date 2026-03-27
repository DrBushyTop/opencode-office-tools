import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("manage_table", "Create or update Excel tables, including style, totals, resizing, row appends/inserts, filter reset, conversion, and deletion.", {
  action: tool.schema.enum(["create", "rename", "resize", "setProperties", "addRows", "clearFilters", "reapplyFilters", "convertToRange", "delete"]).describe("Table operation to perform."),
  tableName: tool.schema.string().optional().describe("Existing table name or id. Required except for create."),
  sheetName: tool.schema.string().optional().describe("Worksheet name for create when sourceRange is sheet-local."),
  sourceRange: tool.schema.string().optional().describe("Source range for create or resize."),
  hasHeaders: tool.schema.boolean().optional().describe("Whether the source range has headers when creating a table."),
  newName: tool.schema.string().optional().describe("New table name for rename or create."),
  style: tool.schema.string().optional().describe("Excel table style to apply."),
  showHeaders: tool.schema.boolean().optional(),
  showTotals: tool.schema.boolean().optional(),
  showBandedRows: tool.schema.boolean().optional(),
  showBandedColumns: tool.schema.boolean().optional(),
  showFilterButton: tool.schema.boolean().optional(),
  highlightFirstColumn: tool.schema.boolean().optional(),
  highlightLastColumn: tool.schema.boolean().optional(),
  values: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))).optional().describe("Row data to append or insert for addRows."),
  rowIndex: tool.schema.number().optional().describe("Zero-based row index for addRows. Omit or use -1 to append."),
  alwaysInsert: tool.schema.boolean().optional().describe("Whether addRows should insert into the table when adding rows. Default true."),
})
