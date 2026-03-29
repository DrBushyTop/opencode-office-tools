import type { Tool } from "./types";
import {
  appendSectionGroupOverview,
  appendSectionOverview,
  formatZodError,
  getNotebookOverviewArgsSchema,
  isOneNoteRequirementSetSupported,
  toolFailure,
} from "./onenoteShared";

export const getNotebookOverview: Tool = {
  name: "get_notebook_overview",
  description: `Get a structural overview of the active OneNote notebook.

Returns:
- active notebook, section, and page metadata
- section and section-group hierarchy
- page titles, ids, and client urls across the notebook

Use this first before navigating or editing pages.`,
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const parsedArgs = getNotebookOverviewArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(formatZodError(parsedArgs.error));
    }

    try {
      return await OneNote.run(async (context) => {
        const app = context.application;
        const notebook = app.getActiveNotebook();
        const section = app.getActiveSection();
        const page = app.getActivePage();

        notebook.load(["id", "name", "baseUrl", "clientUrl"]);
        section.load(["id", "name"]);
        page.load(["id", "title", "pageLevel", "clientUrl", "webUrl"]);
        notebook.sections.load("items/id,name");
        notebook.sectionGroups.load("items/id,name");
        await context.sync();

        const lines = [
          `Notebook ${JSON.stringify(notebook.name)} (${notebook.id})`,
          `Base URL: ${notebook.baseUrl || "(none)"}`,
          `Client URL: ${notebook.clientUrl || "(none)"}`,
          `Active section: ${JSON.stringify(section.name)} (${section.id})`,
          `Active page: ${JSON.stringify(page.title || "Untitled")} (${page.id}), level=${page.pageLevel}`,
          "",
          "Structure:",
        ];

        const activeIds = { sectionId: section.id, pageId: page.id };
        for (const notebookSection of notebook.sections.items) {
          await appendSectionOverview(context, notebookSection, lines, "", activeIds);
        }
        for (const group of notebook.sectionGroups.items) {
          await appendSectionGroupOverview(context, group, lines, "", activeIds);
        }

        return lines.join("\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
