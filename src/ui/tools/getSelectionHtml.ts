import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const getSelectionHtmlArgsSchema = z.object({});

export type GetSelectionHtmlArgs = z.infer<typeof getSelectionHtmlArgsSchema>;

export const getSelectionHtml: Tool = {
  name: "get_selection_html",
  description: "Read the current Word selection as HTML.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    const parsedArgs = getSelectionHtmlArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const html = selection.getHtml();
        await context.sync();
        return html.value || "(empty selection)";
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
