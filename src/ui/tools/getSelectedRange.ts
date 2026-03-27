import type { Tool } from "./types";
import { describeRange, toolFailure } from "./excelShared";

export const getSelectedRange: Tool = {
  name: "get_selected_range",
  description: "Read the currently selected Excel range. Optional detail mode also includes display text, number formats, data validation, merged areas, and overlapping tables or PivotTables.",
  parameters: {
    type: "object",
    properties: {
      detail: {
        type: "boolean",
        description: "Include display text, number formats, data validation, merged areas, and table/PivotTable overlap details.",
      },
    },
  },
  handler: async (args) => {
    const { detail = false } = (args as { detail?: boolean }) || {};

    try {
      return await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const worksheet = selectedRange.worksheet;
        worksheet.load("name");
        await context.sync();

        return await describeRange(context, selectedRange, worksheet.name, {
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
