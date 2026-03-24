import type { Tool } from "@github/copilot-sdk";

export const setDocumentContent: Tool = {
  name: "set_document_content",
  description: "Replace the entire document body with new HTML content. Supports standard HTML tags like <p>, <h1>-<h6>, <ul>, <ol>, <li>, <table>, <b>, <i>, <u>, <a>, etc.",
  parameters: {
    type: "object",
    properties: {
      html: {
        type: "string",
        description: "The HTML content to set as the document body.",
      },
    },
    required: ["html"],
  },
  handler: async (args) => {
    const { html } = args as { html: string };
    
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertHtml(html, Word.InsertLocation.start);
        await context.sync();
        return "Document content replaced successfully.";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
