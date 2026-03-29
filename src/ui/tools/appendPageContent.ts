import type { Tool } from "./types";
import {
  appendPageContentArgsSchema,
  ensureNonEmptyHtml,
  formatZodError,
  isOneNoteRequirementSetSupported,
  loadActivePageOrThrow,
  toolFailure,
} from "./onenoteShared";

export const appendPageContent: Tool = {
  name: "append_page_content",
  description: "Append HTML content to the active OneNote page. Appends to the last outline when possible, or creates a new outline if needed.",
  parameters: {
    type: "object",
    properties: {
      html: {
        type: "string",
        description: "HTML content to append to the active page.",
      },
    },
    required: ["html"],
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const parsedArgs = appendPageContentArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(formatZodError(parsedArgs.error));
    }

    const html = ensureNonEmptyHtml(parsedArgs.data.html);
    if (!html) {
      return toolFailure("HTML content cannot be empty.");
    }

    try {
      return await OneNote.run(async (context) => {
        const page = loadActivePageOrThrow(context);
        const contents = page.contents;
        page.load(["id", "title"]);
        contents.load("items/id,type");
        await context.sync();

        const outlines = contents.items.filter((item) => String(item.type) === "Outline");
        if (outlines.length > 0) {
          outlines[outlines.length - 1].outline.appendHtml(html);
        } else {
          page.addOutline(40, 90, html);
        }

        await context.sync();
        return `Appended content to ${JSON.stringify(page.title || "Untitled page")} (${page.id}).`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
