import type { Tool } from "./types";
import { toolFailure } from "./wordShared";

export const getSelectionHtml: Tool = {
  name: "get_selection_html",
  description: "Read the current Word selection as HTML.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
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
