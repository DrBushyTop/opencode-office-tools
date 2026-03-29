import type { Tool } from "./types";
import { z } from "zod";
import {
  DocumentWriteLocationSchema,
  getZodErrorMessage,
  resolveDocumentRangeTarget,
  toolFailure,
  writeResolvedWordTarget,
} from "./wordShared";

const insertContentAtSelectionArgsSchema = z.object({
  html: z.string(),
  location: DocumentWriteLocationSchema.optional().default("replace"),
});

export type InsertContentAtSelectionArgs = z.infer<typeof insertContentAtSelectionArgsSchema>;

export const insertContentAtSelection: Tool = {
  name: "insert_content_at_selection",
  description: `Insert HTML content at the current cursor position or selection in Word.

This is a surgical edit - it only affects the selected area, not the entire document.

Parameters:
- html: The HTML content to insert. Supports tags like <p>, <h1>-<h6>, <ul>, <ol>, <li>, <table>, <b>, <i>, <u>, <a>, <br>, etc.
- location: Where to insert relative to the selection:
  - "replace" (default): Replace the selected text with the new content
  - "before": Insert before the selection, keeping the selection intact
  - "after": Insert after the selection, keeping the selection intact
  - "start": Insert at the start of the selection
  - "end": Insert at the end of the selection

Examples:
- Insert a paragraph: html = "<p>New paragraph text</p>"
- Insert a heading: html = "<h2>Section Title</h2>"
- Insert a list: html = "<ul><li>Item 1</li><li>Item 2</li></ul>"
- Insert bold text: html = "<b>Important note</b>"`,
  parameters: {
    type: "object",
    properties: {
      html: {
        type: "string",
        description: "The HTML content to insert at the selection.",
      },
      location: {
        type: "string",
        enum: ["replace", "before", "after", "start", "end"],
        description: "Where to insert the content relative to the selection. Default is 'replace'.",
      },
    },
    required: ["html"],
  },
  handler: async (args) => {
    const parsedArgs = insertContentAtSelectionArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { html, location } = parsedArgs.data;
    
    try {
      return await Word.run(async (context) => {
        const selection = await resolveDocumentRangeTarget(context, { kind: "selection" });
        writeResolvedWordTarget(selection, "insert", "html", html, location);
        await context.sync();
        
        return `Content inserted successfully (location: ${location}).`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
