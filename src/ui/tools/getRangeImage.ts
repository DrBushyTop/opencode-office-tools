import { z } from "zod";
import type { Tool } from "./types";
import { getWorksheet, isExcelRequirementSetSupported, parseToolArgs } from "./excelShared";
import { createToolFailure, describeErrorWithCode } from "./toolShared";

const getRangeImageArgsSchema = z.object({
  range: z.string(),
  sheetName: z.string().optional(),
});

export const getRangeImage: Tool = {
  name: "get_range_image",
  description: "Render an Excel range as a PNG snapshot. Useful for checking layout, truncation, spacing, wrapping, and readability after formatting changes.",
  parameters: {
    type: "object",
    properties: {
      range: { type: "string", description: "Target range such as A1:F12." },
      sheetName: { type: "string", description: "Optional worksheet name. Defaults to the active sheet." },
    },
    required: ["range"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(getRangeImageArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    if (!isExcelRequirementSetSupported("1.7")) {
      return createToolFailure("This Excel host cannot export range images. Use a host with ExcelApi 1.7+ and try again.");
    }

    try {
      return await Excel.run(async (context) => {
        const sheet = await getWorksheet(context, parsedArgs.data.sheetName);
        const range = sheet.getRange(parsedArgs.data.range);
        range.load("address");
        const image = range.getImage();
        await context.sync();

        return {
          textResultForLlm: `Rendered ${range.address} in ${sheet.name} as a PNG snapshot.`,
          binaryResultsForLlm: [
            {
              data: image.value,
              mimeType: "image/png",
              type: "image",
              description: `${sheet.name} ${range.address}`,
            },
          ],
          resultType: "success" as const,
          toolTelemetry: {},
        };
      });
    } catch (error: unknown) {
      return createToolFailure(error, { describe: describeErrorWithCode });
    }
  },
};
