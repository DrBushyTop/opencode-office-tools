import type { Tool } from "./types";
import {
  createPageArgsSchema,
  ensureNonEmptyHtml,
  formatZodError,
  isOneNoteRequirementSetSupported,
  loadActivePageOrThrow,
  loadActiveSectionOrThrow,
  toolFailure,
} from "./onenoteShared";

export const createPage: Tool = {
  name: "create_page",
  description: "Create a new OneNote page in the active section or before/after the active page, with optional initial HTML content.",
  parameters: {
    type: "object",
    properties: {
      title: {
        type: "string",
        description: "New page title. Default is New Page.",
      },
      html: {
        type: "string",
        description: "Optional initial HTML content to place on the new page as an outline.",
      },
      location: {
        type: "string",
        enum: ["sectionEnd", "before", "after"],
        description: "Where to create the page. Default is sectionEnd.",
      },
    },
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const parsedArgs = createPageArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(formatZodError(parsedArgs.error));
    }

    const { title = "New Page", html, location = "sectionEnd" } = parsedArgs.data;
    const normalizedTitle = String(title || "").trim() || "New Page";
    const normalizedHtml = html === undefined ? "" : ensureNonEmptyHtml(html);

    try {
      return await OneNote.run(async (context) => {
        const page = location === "sectionEnd"
          ? loadActiveSectionOrThrow(context).addPage(normalizedTitle)
          : loadActivePageOrThrow(context).insertPageAsSibling(location === "before" ? "Before" : "After", normalizedTitle);

        page.load(["id", "title"]);
        if (normalizedHtml) {
          page.addOutline(40, 90, normalizedHtml);
        }
        await context.sync();

        return `Created page ${JSON.stringify(page.title || normalizedTitle)} (${page.id}).`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
