import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const getDocumentContentArgsSchema = z.object({});

export type GetDocumentContentArgs = z.infer<typeof getDocumentContentArgsSchema>;

export const getDocumentContent: Tool = {
  name: "get_document_content",
  description: "Get the HTML content of the Word document.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    const parsedArgs = getDocumentContentArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const html = body.getHtml();
        await context.sync();
        return html.value || "(empty document)";
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
