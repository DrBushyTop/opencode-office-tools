import type { Tool } from "./types";
import { describeRange, getWorksheet, toolFailure } from "./excelShared";

export const getWorkbookContent: Tool = {
  name: "get_workbook_content",
  description: "Read values and formulas from a worksheet or range. Optional detail mode also includes display text, number formats, data validation, merged areas, and overlapping tables or PivotTables.",
  parameters: {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: "Optional name of the worksheet to read. Defaults to the active sheet.",
      },
      range: {
        type: "string",
        description: "Optional cell range to read (for example, 'A1:D10'). Defaults to the used range of the sheet.",
      },
      detail: {
        type: "boolean",
        description: "Include display text, number formats, data validation, merged areas, and table/PivotTable overlap details.",
      },
    },
  },
  handler: async (args) => {
    const { sheetName, range, detail = false } = (args as { sheetName?: string; range?: string; detail?: boolean }) || {};

    try {
      return await Excel.run(async (context) => {
        const worksheet = await getWorksheet(context, sheetName);
        if (range) {
          const targetRange = worksheet.getRange(range);
          return await describeRange(context, targetRange, worksheet.name, {
            detail,
            includeNumberFormats: detail,
            includeTables: detail,
            includeValidation: detail,
            includeMergedAreas: detail,
          });
        }

        const targetRange = worksheet.getUsedRangeOrNullObject();
        targetRange.load(["isNullObject"]);
        await context.sync();

        if ((targetRange as Excel.Range & { isNullObject?: boolean }).isNullObject) {
          return `Worksheet: ${worksheet.name}\nRange: (empty used range)\n\n(empty range)`;
        }

        return await describeRange(context, targetRange as Excel.Range, worksheet.name, {
          detail,
          includeNumberFormats: detail,
          includeTables: detail,
          includeValidation: detail,
          includeMergedAreas: detail,
        });
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
