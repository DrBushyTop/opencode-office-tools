import type { Tool } from "./types";
import { isExcelRequirementSetSupported, toolFailure } from "./excelShared";

export const getWorkbookOverview: Tool = {
  name: "get_workbook_overview",
  description: "Get a structural overview of the Excel workbook, including worksheets, used ranges, visibility, protection, tables, PivotTables, charts, named ranges, filters, and frozen panes.",
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

        const lines: string[] = [
          `Workbook overview`,
          `${"━".repeat(40)}`,
          `Worksheets: ${sheets.items.length}`,
          `Active sheet: ${activeSheet.name}`,
        ];

        let totalCells = 0;
        let totalCharts = 0;
        let totalTables = 0;
        let totalPivotTables = 0;

        for (const [index, sheet] of sheets.items.entries()) {
          const usedRange = usedRanges[index];
          const chartCount = chartCounts[index].value;
          const tableCount = tableCounts[index].value;
          const pivotCount = supportsPivotTables ? pivotCounts[index].value : 0;
          const visibility = sheet.visibility || "Visible";
          const protection = protections[index].protected ? "protected" : "unprotected";
          const usedRangeText = usedRange.isNullObject
            ? "(empty)"
            : `${usedRange.address} (${usedRange.rowCount} rows x ${usedRange.columnCount} cols)`;

          if (!usedRange.isNullObject) {
            totalCells += usedRange.rowCount * usedRange.columnCount;
          }
          totalCharts += chartCount;
          totalTables += tableCount;
          totalPivotTables += pivotCount;

          lines.push(
            "",
            `- ${sheet.name}${sheet.name === activeSheet.name ? " <- active" : ""}`,
            `  id=${sheet.id}, position=${sheet.position}, visibility=${visibility}, ${protection}`,
            `  usedRange=${usedRangeText}`,
            `  charts=${chartCount}, tables=${tableCount}, pivotTables=${pivotCount}`,
          );

          if (supportsAutoFilter) {
            lines.push(`  autoFilter=${autoFilters[index].enabled ? "enabled" : "disabled"}, filtered=${autoFilters[index].isDataFiltered ? "yes" : "no"}`);
          }

          if (supportsFreeze) {
            lines.push(`  frozenPanes=${freezeRanges[index].isNullObject ? "(none)" : freezeRanges[index].address}`);
          }

          if (tableCollections[index].items.length) {
            lines.push(`  tableDetails=${tableCollections[index].items.map((table) => `${table.name} (${table.style}${table.showTotals ? ", totals" : ""})`).join(", ")}`);
          }

          if (sheetNames[index].items.length) {
            lines.push(`  worksheetNames=${sheetNames[index].items.map((name) => name.name).join(", ")}`);
          }
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
