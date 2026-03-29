import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const getSelectionArgsSchema = z.object({});

export type GetSelectionArgs = z.infer<typeof getSelectionArgsSchema>;

export const getSelection: Tool = {
  name: "get_selection",
  description: "Read the current Word selection as OOXML.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    const parsedArgs = getSelectionArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

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
