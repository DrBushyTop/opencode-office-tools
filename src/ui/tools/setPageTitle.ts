import type { Tool } from "./types";
import {
  isOneNoteRequirementSetSupported,
  loadActivePageOrThrow,
  toolFailure,
} from "./onenoteShared";

export const setPageTitle: Tool = {
  name: "set_page_title",
  description: "Update the title of the active OneNote page.",
  parameters: {
    type: "object",
    properties: {
      title: {
        type: "string",
        description: "New title for the active page.",
      },
    },
    required: ["title"],
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const title = String((args as { title: string }).title || "").trim();
    if (!title) {
      return toolFailure("Title cannot be empty.");
    }

    try {
      return await OneNote.run(async (context) => {
        const page = loadActivePageOrThrow(context);
        page.title = title;
        page.load(["id", "title"]);
        await context.sync();
        return `Renamed the active page to ${JSON.stringify(page.title)} (${page.id}).`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
