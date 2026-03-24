import type { Tool } from "@github/copilot-sdk";

export const getWorkbookOverview: Tool = {
  name: "get_workbook_overview",
  description: `Get a structural overview of the Excel workbook. Use this first to understand the workbook before reading or editing specific sheets.

Returns:
- List of all worksheets with their used range dimensions
- Active worksheet name
- Named ranges defined in the workbook
- Chart count per sheet
- Total cell count with data

This is faster than reading entire sheets and helps you understand what to target for edits.`,
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
        
        sheets.load("items");
        names.load("items");
        workbook.load("name");
        
        await context.sync();

        // Get active sheet
        const activeSheet = workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        await context.sync();

        const sheetInfos: string[] = [];
        let totalCells = 0;
        let totalCharts = 0;

        // Load details for each sheet
        for (const sheet of sheets.items) {
          sheet.load("name");
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load(["address", "rowCount", "columnCount"]);
          const charts = sheet.charts;
          charts.load("count");
        }
        await context.sync();

        for (const sheet of sheets.items) {
          const usedRange = sheet.getUsedRangeOrNullObject();
          const charts = sheet.charts;
          
          let rangeInfo = "(empty)";
          let cellCount = 0;
          
          if (!usedRange.isNullObject) {
            const rows = usedRange.rowCount;
            const cols = usedRange.columnCount;
            cellCount = rows * cols;
            rangeInfo = `${usedRange.address} (${rows} rows × ${cols} cols)`;
          }
          
          totalCells += cellCount;
          totalCharts += charts.count;
          
          const isActive = sheet.name === activeSheet.name ? " ← active" : "";
          const chartInfo = charts.count > 0 ? `, ${charts.count} chart(s)` : "";
          
          sheetInfos.push(`  • ${sheet.name}: ${rangeInfo}${chartInfo}${isActive}`);
        }

        // Get named ranges
        const namedRanges: string[] = [];
        for (const name of names.items) {
          name.load(["name", "value"]);
        }
        await context.sync();
        
        for (const name of names.items) {
          namedRanges.push(`  • ${name.name}: ${name.value}`);
        }

        // Build output
        let output = `Workbook Overview:\n`;
        output += `${"━".repeat(40)}\n`;
        output += `Worksheets (${sheets.items.length}):\n`;
        output += sheetInfos.join("\n");
        output += `\n\nTotal cells with data: ${totalCells.toLocaleString()}`;
        
        if (totalCharts > 0) {
          output += `\nTotal charts: ${totalCharts}`;
        }
        
        if (namedRanges.length > 0) {
          output += `\n\nNamed Ranges (${namedRanges.length}):\n`;
          output += namedRanges.join("\n");
        }

        return output;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
