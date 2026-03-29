import type { Tool } from "./types";
import {
  asJsonString,
  formatPageSummary,
  formatPageText,
  formatZodError,
  getPageContentArgsSchema,
  isOneNoteRequirementSetSupported,
  loadActivePageOrThrow,
  loadPageContentSummaries,
  toolFailure,
} from "./onenoteShared";

export const getPageContent: Tool = {
  name: "get_page_content",
  description: `Read the active OneNote page.

Formats:
- summary: structural summary and preview
- text: extracted plain text with placeholders for non-text items
- json: page analysis JSON when available, otherwise structured extracted JSON

OneNote only exposes full page content for the active page.`,
  parameters: {
    type: "object",
    properties: {
      format: {
        type: "string",
        enum: ["summary", "text", "json"],
        description: "Preferred response format. Default is summary.",
      },
    },
  },
  handler: async (args) => {
    if (!isOneNoteRequirementSetSupported("1.1")) {
      return toolFailure("OneNoteApi 1.1 is required.");
    }

    const parsedArgs = getPageContentArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(formatZodError(parsedArgs.error));
    }

    const { format = "summary" } = parsedArgs.data;

    try {
      return await OneNote.run(async (context) => {
        const page = loadActivePageOrThrow(context);
        page.load(["id", "title", "pageLevel", "clientUrl", "webUrl"]);
        const analysis = format === "json" ? page.analyzePage() : null;
        await context.sync();

        if (format === "json" && analysis?.value) {
          try {
            return asJsonString(JSON.parse(analysis.value));
          } catch {
            return analysis.value;
          }
        }

        const contentItems = await loadPageContentSummaries(context, page, format);
        if (format === "text") {
          return formatPageText(contentItems);
        }

        if (format === "json") {
          return asJsonString({
            page: {
              id: page.id,
              title: page.title,
              pageLevel: page.pageLevel,
              clientUrl: page.clientUrl,
              webUrl: page.webUrl,
            },
            contents: contentItems,
          });
        }

        return formatPageSummary({ title: page.title, id: page.id, pageLevel: page.pageLevel }, contentItems);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
