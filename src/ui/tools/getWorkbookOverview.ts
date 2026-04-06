import type { Tool } from "./types";
import { isExcelRequirementSetSupported, toolFailure } from "./excelShared";

type WorksheetOverview = {
  id: string;
  name: string;
  position: number;
  visibility: string;
  protection: string;
  usedRangeText: string;
  usedCellCount: number;
  chartCount: number;
  tableCount: number;
  pivotCount: number;
  autoFilterText?: string;
  frozenPaneText?: string;
  tableDetails?: string;
  worksheetNames?: string;
};

function formatWorksheetLines(sheet: WorksheetOverview, activeSheetName: string) {
  const lines = [
    `- ${sheet.name}${sheet.name === activeSheetName ? " <- active" : ""}`,
    `  id=${sheet.id}, position=${sheet.position}, visibility=${sheet.visibility}, ${sheet.protection}`,
    `  usedRange=${sheet.usedRangeText}`,
    `  charts=${sheet.chartCount}, tables=${sheet.tableCount}, pivotTables=${sheet.pivotCount}`,
  ];

  if (sheet.autoFilterText) lines.push(`  autoFilter=${sheet.autoFilterText}`);
  if (sheet.frozenPaneText) lines.push(`  frozenPanes=${sheet.frozenPaneText}`);
  if (sheet.tableDetails) lines.push(`  tableDetails=${sheet.tableDetails}`);
  if (sheet.worksheetNames) lines.push(`  worksheetNames=${sheet.worksheetNames}`);
  return lines;
}

export const getWorkbookOverview: Tool = {
  name: "get_workbook_overview",
  description: "Summarize the current Excel workbook structure, including sheets, used ranges, charts, tables, names, filters, and frozen panes when the host supports them.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const sheets = workbook.worksheets;
        const names = workbook.names;
        const supportsFreeze = isExcelRequirementSetSupported("1.7");
        const supportsAutoFilter = isExcelRequirementSetSupported("1.9");
        const supportsPivotTables = isExcelRequirementSetSupported("1.3");

        sheets.load("items");
        names.load("items");
        await context.sync();

        const activeSheet = workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");

        const usedRanges = sheets.items.map((sheet) => sheet.getUsedRangeOrNullObject());
        const chartCounts = sheets.items.map((sheet) => sheet.charts.getCount());
        const tableCounts = sheets.items.map((sheet) => sheet.tables.getCount());
        const pivotCounts = supportsPivotTables ? sheets.items.map((sheet) => sheet.pivotTables.getCount()) : [];
        const tableCollections = sheets.items.map((sheet) => sheet.tables);
        const sheetNames = sheets.items.map((sheet) => sheet.names);
        const protections = sheets.items.map((sheet) => sheet.protection);
        const freezeRanges = supportsFreeze ? sheets.items.map((sheet) => sheet.freezePanes.getLocationOrNullObject()) : [];
        const autoFilters = supportsAutoFilter ? sheets.items.map((sheet) => sheet.autoFilter) : [];

        for (const [index, sheet] of sheets.items.entries()) {
          sheet.load(["id", "name", "position", "visibility"]);
          usedRanges[index].load(["isNullObject", "address", "rowCount", "columnCount"]);
          tableCollections[index].load("items/name,items/style,items/showTotals");
          sheetNames[index].load("items/name");
          protections[index].load("protected");
          if (supportsFreeze) {
            freezeRanges[index].load(["isNullObject", "address"]);
          }
          if (supportsAutoFilter) {
            autoFilters[index].load(["enabled", "isDataFiltered"]);
          }
        }

        for (const name of names.items) {
          name.load(["name", "value", "scope"]);
        }

        await context.sync();

        const worksheetSummaries: WorksheetOverview[] = sheets.items.map((sheet, index) => {
          const usedRange = usedRanges[index];
          const chartCount = chartCounts[index].value;
          const tableCount = tableCounts[index].value;
          const pivotCount = supportsPivotTables ? pivotCounts[index].value : 0;
          const usedRangeText = usedRange.isNullObject
            ? "(empty)"
            : `${usedRange.address} (${usedRange.rowCount} rows x ${usedRange.columnCount} cols)`;

          return {
            id: sheet.id,
            name: sheet.name,
            position: sheet.position,
            visibility: sheet.visibility || "Visible",
            protection: protections[index].protected ? "protected" : "unprotected",
            usedRangeText,
            usedCellCount: usedRange.isNullObject ? 0 : usedRange.rowCount * usedRange.columnCount,
            chartCount,
            tableCount,
            pivotCount,
            autoFilterText: supportsAutoFilter
              ? `${autoFilters[index].enabled ? "enabled" : "disabled"}, filtered=${autoFilters[index].isDataFiltered ? "yes" : "no"}`
              : undefined,
            frozenPaneText: supportsFreeze
              ? (freezeRanges[index].isNullObject ? "(none)" : freezeRanges[index].address)
              : undefined,
            tableDetails: tableCollections[index].items.length
              ? tableCollections[index].items.map((table) => `${table.name} (${table.style}${table.showTotals ? ", totals" : ""})`).join(", ")
              : undefined,
            worksheetNames: sheetNames[index].items.length
              ? sheetNames[index].items.map((name) => name.name).join(", ")
              : undefined,
          };
        });

        const totalCells = worksheetSummaries.reduce((sum, sheet) => sum + sheet.usedCellCount, 0);
        const totalCharts = worksheetSummaries.reduce((sum, sheet) => sum + sheet.chartCount, 0);
        const totalTables = worksheetSummaries.reduce((sum, sheet) => sum + sheet.tableCount, 0);
        const totalPivotTables = worksheetSummaries.reduce((sum, sheet) => sum + sheet.pivotCount, 0);

        const lines: string[] = [
          `Workbook overview`,
          `${"━".repeat(40)}`,
          `Worksheets: ${sheets.items.length}`,
          `Active sheet: ${activeSheet.name}`,
        ];

        for (const summary of worksheetSummaries) {
          lines.push("", ...formatWorksheetLines(summary, activeSheet.name));
        }

        lines.push(
          "",
          `Total cells with data: ${totalCells.toLocaleString()}`,
          `Total charts: ${totalCharts}`,
          `Total tables: ${totalTables}`,
          `Total PivotTables: ${totalPivotTables}`,
        );

        if (names.items.length) {
          lines.push("", `Workbook names (${names.items.length}):`);
          for (const name of names.items) {
            lines.push(`- ${name.name} [${name.scope}]: ${name.value}`);
          }
        }

        return lines.join("\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
