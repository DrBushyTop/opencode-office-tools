import type { Tool } from "./types";
import {
  findPageById,
  isOneNoteRequirementSetSupported,
  toolFailure,
} from "./onenoteShared";

export const navigateToPage: Tool = {
  name: "navigate_to_page",
  description: "Navigate OneNote to a specific page by page id or client URL.",
  parameters: {
    type: "object",
    properties: {
      pageId: {
        type: "string",
        description: "Page id from get_notebook_overview. Provide exactly one of pageId or clientUrl.",
      },
      clientUrl: {
        type: "string",
        description: "Client URL of the page to open. Provide exactly one of pageId or clientUrl.",
      },
    },
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const { pageId, clientUrl } = (args as { pageId?: string; clientUrl?: string }) || {};
    if ((!pageId && !clientUrl) || (pageId && clientUrl)) {
      return toolFailure("Provide exactly one of pageId or clientUrl.");
    }

    try {
      return await OneNote.run(async (context) => {
        const app = context.application;

        if (clientUrl) {
          const page = app.navigateToPageWithClientUrl(clientUrl);
          const activePage = app.getActivePage();
          page.load(["id", "title", "clientUrl"]);
          activePage.load(["id", "title", "clientUrl"]);
          await context.sync();

          if (activePage.id !== page.id) {
            return toolFailure("OneNote did not navigate to the requested clientUrl.");
          }

          return `Navigated to page ${JSON.stringify(activePage.title || "Untitled")} (${activePage.id}).`;
        }

        const notebook = app.getActiveNotebook();
        notebook.load(["id", "name"]);
        await context.sync();
        const page = await findPageById(context, notebook, String(pageId));
        if (!page) {
          return toolFailure(`Page ${pageId} was not found in the active notebook.`);
        }

        app.navigateToPage(page);
        page.load(["id", "title"]);
        await context.sync();
        return `Navigated to page ${JSON.stringify(page.title || "Untitled")} (${page.id}).`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
