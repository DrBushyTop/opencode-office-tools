import type { Tool } from "@github/copilot-sdk";

export const getSelection: Tool = {
  name: "get_selection",
  description: "Get the currently selected OOXML content (w:document element only).",
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
        
        // Extract just the w:document element
        const match = ooxml.value.match(/<w:document[^>]*>[\s\S]*<\/w:document>/);
        return match ? match[0] : "(no selection)";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
