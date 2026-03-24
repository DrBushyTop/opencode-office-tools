import type { Tool } from "@github/copilot-sdk";

export const getDocumentContent: Tool = {
  name: "get_document_content",
  description: "Get the HTML content of the Word document.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const html = body.getHtml();
        await context.sync();
        return html.value || "(empty document)";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
