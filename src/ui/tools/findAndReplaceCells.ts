import type { Tool } from "@github/copilot-sdk";

export const findAndReplaceCells: Tool = {
  name: "find_and_replace_cells",
  description: `Find and replace text in Excel cells.

Parameters:
- find: The text to search for
- replace: The text to replace it with
- sheetName: Optional worksheet name. If omitted, searches the active sheet.
- matchCase: If true, search is case-sensitive (default: false)
- matchEntireCell: If true, only match cells where the entire content matches (default: false)

Returns the number of replacements made.

Examples:
- Replace all "TBD" with "Complete": find="TBD", replace="Complete"
- Replace in specific sheet: find="2023", replace="2024", sheetName="Sales Data"
- Match exact cell content: find="Yes", replace="Approved", matchEntireCell=true`,
  parameters: {
    type: "object",
    properties: {
      find: {
        type: "string",
        description: "The text to search for.",
      },
      replace: {
        type: "string",
        description: "The text to replace matches with.",
      },
      sheetName: {
        type: "string",
        description: "Optional worksheet name. Defaults to active sheet.",
      },
      matchCase: {
        type: "boolean",
        description: "If true, the search is case-sensitive. Default is false.",
      },
      matchEntireCell: {
        type: "boolean",
        description: "If true, only matches cells where entire content matches. Default is false.",
      },
    },
    required: ["find", "replace"],
  },
  handler: async (args) => {
    const { find, replace, sheetName, matchCase = false, matchEntireCell = false } = args as {
      find: string;
      replace: string;
      sheetName?: string;
      matchCase?: boolean;
      matchEntireCell?: boolean;
    };

    if (!find || find.length === 0) {
      return { textResultForLlm: "Search text cannot be empty.", resultType: "failure", error: "Empty search", toolTelemetry: {} };
    }

    try {
      return await Excel.run(async (context) => {
        // Get the target worksheet
        let sheet: Excel.Worksheet;
        if (sheetName) {
          sheet = context.workbook.worksheets.getItem(sheetName);
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        sheet.load("name");

        // Get used range
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["values", "address", "rowCount", "columnCount"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return `No data found in worksheet "${sheet.name}".`;
        }

        const values = usedRange.values;
        let replacementCount = 0;

        // Search and replace in values
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            if (cellValue === null || cellValue === undefined) continue;

            const cellStr = String(cellValue);
            const searchStr = matchCase ? find : find.toLowerCase();
            const compareStr = matchCase ? cellStr : cellStr.toLowerCase();

            if (matchEntireCell) {
              if (compareStr === searchStr) {
                values[row][col] = replace;
                replacementCount++;
              }
            } else {
              if (compareStr.includes(searchStr)) {
                // Replace all occurrences in the cell
                const regex = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
                values[row][col] = cellStr.replace(regex, replace);
                replacementCount++;
              }
            }
          }
        }

        if (replacementCount === 0) {
          return `No matches found for "${find}" in worksheet "${sheet.name}".`;
        }

        // Write back the modified values
        usedRange.values = values;
        await context.sync();

        return `Replaced ${replacementCount} cell(s) containing "${find}" with "${replace}" in worksheet "${sheet.name}".`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
