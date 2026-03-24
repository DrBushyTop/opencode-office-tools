import type { Tool } from "@github/copilot-sdk";

export const getSelectionText: Tool = {
  name: "get_selection_text",
  description: `Get the currently selected text in the Word document as plain readable text.

Returns the selected text content. If nothing is selected, returns the text at the cursor position (which may be empty).

Use this to understand what the user has highlighted before making changes to it.`,
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        const text = selection.text || "";
        if (text.trim().length === 0) {
          return "(No text selected - cursor is at an empty position)";
        }
        return text;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
