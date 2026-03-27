import type { Tool } from "./types";
import { toolFailure } from "./wordShared";

export const getSelection: Tool = {
  name: "get_selection",
  description: "Read the current Word selection as OOXML.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const ooxml = selection.getOoxml();
        await context.sync();

        return ooxml.value || "(no selection)";
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
