import type { Tool } from "./types";
import {
  formatZodError,
  getNoteSelectionArgsSchema,
  getSelectedDataAsync,
  toolFailure,
} from "./onenoteShared";

export const getNoteSelection: Tool = {
  name: "get_note_selection",
  description: "Read the current OneNote selection as plain text or a matrix of values.",
  parameters: {
    type: "object",
    properties: {
      format: {
        type: "string",
        enum: ["text", "matrix"],
        description: "Selection format to read. Default is text.",
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = getNoteSelectionArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(formatZodError(parsedArgs.error));
    }

    const { format = "text" } = parsedArgs.data;

    try {
      if (format === "matrix") {
        const value = await getSelectedDataAsync<unknown[][]>(Office.CoercionType.Matrix);
        return JSON.stringify(value, null, 2);
      }

      const value = await getSelectedDataAsync<string>(Office.CoercionType.Text);
      return value || "(empty selection)";
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
