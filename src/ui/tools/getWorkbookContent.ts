import { z } from "zod";
import type { Tool } from "./types";
import { buildRangeDescribeOptions, describeRange, getWorksheet, parseToolArgs, toolFailure } from "./excelShared";

const getWorkbookContentArgsSchema = z.object({
  sheetName: z.string().optional(),
  range: z.string().optional(),
  detail: z.boolean().default(false),
});

export const getWorkbookContent: Tool = {
  name: "get_workbook_content",
  description: "Read worksheet cells from Excel. You can target a named sheet and optional range, or inspect the active sheet's used range with optional metadata.",
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
    const parsedArgs = parseToolArgs(getWorkbookContentArgsSchema, args ?? {});
    if (!parsedArgs.success) return parsedArgs.failure;

    const { sheetName, range, detail } = parsedArgs.data;
    const describeOptions = buildRangeDescribeOptions(detail);

    try {
      return await Excel.run(async (context) => {
        const worksheet = await getWorksheet(context, sheetName);
        if (range) {
          const targetRange = worksheet.getRange(range);
          return await describeRange(context, targetRange, worksheet.name, describeOptions);
        }

        const targetRange = worksheet.getUsedRangeOrNullObject();
        targetRange.load(["isNullObject"]);
        await context.sync();

        if ((targetRange as Excel.Range & { isNullObject?: boolean }).isNullObject) {
          return `Worksheet: ${worksheet.name}\nRange: (empty used range)\n\nNo populated cells were found.`;
        }

        return await describeRange(context, targetRange as Excel.Range, worksheet.name, describeOptions);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
