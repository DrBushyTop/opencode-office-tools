import type { Tool } from "./types";
import {
  parseDocumentRangeAddress,
  resolveDocumentRangeTarget,
  summarizePlainText,
  toolFailure,
} from "./wordShared";

export const findDocumentText: Tool = {
  name: "find_document_text",
  description: `Locate text in the active Word document without modifying it.

Optionally scope the search to a generic address such as selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3].`,
  parameters: {
    type: "object",
    properties: {
      find: {
        type: "string",
        description: "Text to find.",
      },
      address: {
        type: "string",
        description: "Optional scope address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3].",
      },
      matchCase: {
        type: "boolean",
        description: "Match case exactly.",
      },
      matchWholeWord: {
        type: "boolean",
        description: "Match whole words only.",
      },
      maxResults: {
        type: "number",
        description: "Maximum number of preview matches to return. Default is 20.",
      },
    },
    required: ["find"],
  },
  handler: async (args) => {
    const {
      find,
      address,
      matchCase = false,
      matchWholeWord = false,
      maxResults = 20,
    } = args as {
      find: string;
      address?: string;
      matchCase?: boolean;
      matchWholeWord?: boolean;
      maxResults?: number;
    };

    if (!find.trim()) {
      return toolFailure("Search text cannot be empty.");
    }

    if (!Number.isFinite(maxResults) || maxResults < 1) {
      return toolFailure("maxResults must be a positive number.");
    }

    const parsed = address ? parseDocumentRangeAddress(address) : null;
    if (address && !parsed) {
      return toolFailure("Unsupported scope address. Try selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3].");
    }

    try {
      return await Word.run(async (context) => {
        const resolved = parsed
          ? await resolveDocumentRangeTarget(context, parsed)
          : { kind: "body" as const, label: "document", target: context.document.body };

        const matches = resolved.target.search(find, {
          ignorePunct: false,
          ignoreSpace: false,
          matchCase,
          matchWholeWord,
        });
        matches.load("items");
        await context.sync();

        const previewItems = matches.items.slice(0, Math.floor(maxResults));
        for (const item of previewItems) {
          item.load("text");
        }
        await context.sync();

        return {
          scope: resolved.label,
          find,
          count: matches.items.length,
          truncated: matches.items.length > previewItems.length,
          matches: previewItems.map((item, index) => ({
            index: index + 1,
            preview: summarizePlainText(item.text || "", 120),
          })),
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
