import type { Tool } from "./types";
import { getWorksheet, isExcelRequirementSetSupported, toolFailure } from "./excelShared";

export const findAndReplaceCells: Tool = {
  name: "find_and_replace_cells",
  description: "Find and replace text in Excel using Excel's native replace API so formulas and search semantics are preserved.",
  parameters: {
    type: "object",
    properties: {
      find: { type: "string", description: "The text to search for." },
      replace: { type: "string", description: "The replacement text." },
      sheetName: { type: "string", description: "Optional worksheet name. Defaults to the active sheet." },
      matchCase: { type: "boolean", description: "Whether the search is case-sensitive." },
      matchEntireCell: { type: "boolean", description: "Whether the whole cell must match." },
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

    if (!find) {
      return toolFailure("Search text cannot be empty.");
    }
    if (!isExcelRequirementSetSupported("1.9")) {
      return toolFailure("Native Excel find and replace requires ExcelApi 1.9.");
    }

    try {
      return await Excel.run(async (context) => {
        const sheet = await getWorksheet(context, sheetName);
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["isNullObject", "address"]);
        await context.sync();

        if ((usedRange as Excel.Range & { isNullObject?: boolean }).isNullObject) {
          return `No data found in worksheet ${sheet.name}.`;
        }

        const count = (usedRange as Excel.Range).replaceAll(find, replace, {
          matchCase,
          completeMatch: matchEntireCell,
        });
        await context.sync();

        if (count.value === 0) {
          return `No matches found for ${JSON.stringify(find)} in ${sheet.name}.`;
        }

        return `Replaced ${count.value} occurrence(s) of ${JSON.stringify(find)} with ${JSON.stringify(replace)} in ${usedRange.address} on ${sheet.name}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
