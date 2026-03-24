import type { Tool } from "@github/copilot-sdk";

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
    const { html, location = "replace" } = args as { html: string; location?: string };
    
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        
        // Map location string to Word.InsertLocation
        let insertLocation: Word.InsertLocation;
        switch (location) {
          case "before":
            insertLocation = Word.InsertLocation.before;
            break;
          case "after":
            insertLocation = Word.InsertLocation.after;
            break;
          case "start":
            insertLocation = Word.InsertLocation.start;
            break;
          case "end":
            insertLocation = Word.InsertLocation.end;
            break;
          case "replace":
          default:
            insertLocation = Word.InsertLocation.replace;
            break;
        }
        
        selection.insertHtml(html, insertLocation);
        await context.sync();
        
        return `Content inserted successfully (location: ${location}).`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
