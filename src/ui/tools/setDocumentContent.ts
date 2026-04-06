import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const argsSchema = z.object({
  html: z.string().min(1, "html must not be empty"),
});

/**
 * Overwrites the active document body with the supplied HTML markup.
 * Standard inline/block HTML elements are accepted (headings, lists,
 * tables, anchors, emphasis, etc.).
 */
export const setDocumentContent: Tool = {
  name: "set_document_content",
  description: "Replace the current Word document with new HTML content.",
  parameters: {
    type: "object",
    properties: {
      html: {
        type: "string",
        description: "HTML markup that will become the new document body.",
      },
    },
    required: ["html"],
  },

  handler: async (args) => {
    const parsed = argsSchema.safeParse(args ?? {});
    if (!parsed.success) return toolFailure(getZodErrorMessage(parsed.error));

    try {
      return await Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.clear();
        body.insertHtml(parsed.data.html, Word.InsertLocation.replace);
        await ctx.sync();
        return "Document body replaced.";
      });
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
