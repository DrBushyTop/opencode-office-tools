import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const setDocumentContentArgsSchema = z.object({
  html: z.string(),
});

export type SetDocumentContentArgs = z.infer<typeof setDocumentContentArgsSchema>;

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
    const parsedArgs = setDocumentContentArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { html } = parsedArgs.data;
    
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertHtml(html, Word.InsertLocation.start);
        await context.sync();
        return "Document content replaced successfully.";
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
