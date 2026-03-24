import type { Tool } from "@github/copilot-sdk";

export const getWorkbookInfo: Tool = {
  name: "get_workbook_info",
  description: "Get metadata about the Excel workbook including all worksheet/tab names, the active sheet, and basic workbook properties. Use this to understand the structure of the workbook before reading or writing data.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        const activeWorksheet = workbook.worksheets.getActiveWorksheet();

        worksheets.load(["items/name", "items/position"]);
        activeWorksheet.load("name");

        await context.sync();

        const sheetInfo = worksheets.items
          .sort((a, b) => a.position - b.position)
          .map((sheet, index) => `${index + 1}. ${sheet.name}`);

        let output = `Workbook Structure\n`;
        output += `==================\n\n`;
        output += `Active Sheet: ${activeWorksheet.name}\n\n`;
        output += `All Sheets (${worksheets.items.length}):\n`;
        output += sheetInfo.join("\n");

        return output;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
