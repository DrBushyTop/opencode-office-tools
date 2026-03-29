import type { Tool } from "./types";
import { z } from "zod";
import {
  getZodErrorMessage,
  isWordDesktopRequirementSetSupported,
  summarizePlainText,
  toolFailure,
} from "./wordShared";

const getDocumentTargetsArgsSchema = z.object({
  kind: z.enum(["all", "tables", "contentControls", "bookmarks"]).optional().default("all"),
  maxItems: z.number().int().positive().optional().default(25),
  includeText: z.boolean().optional().default(true),
});

export type GetDocumentTargetsArgs = z.infer<typeof getDocumentTargetsArgsSchema>;

export const getDocumentTargets: Tool = {
  name: "get_document_targets",
  description: `Inspect discoverable Word targets for later generic addressing.

Lists tables, content controls, and bookmarks so later reads and writes can target addresses like table[1], table[1].cell[2,3], bookmark[Name], or content_control[id=12].`,
  parameters: {
    type: "object",
    properties: {
      kind: {
        type: "string",
        enum: ["all", "tables", "contentControls", "bookmarks"],
        description: "Which target family to inspect. Default is all.",
      },
      maxItems: {
        type: "number",
        description: "Maximum items to include per family. Default is 25.",
      },
      includeText: {
        type: "boolean",
        description: "Include short text previews when available.",
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = getDocumentTargetsArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { kind, maxItems, includeText } = parsedArgs.data;

    const limit = Math.floor(maxItems);

    try {
      return await Word.run(async (context) => {
        const result: Record<string, unknown> = {};

        if (kind === "all" || kind === "tables") {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();

          const items = tables.items.slice(0, limit);
          for (const table of items) {
            table.load(["rowCount", "values", "styleBandedRows", "styleFirstColumn", "styleLastColumn"]);
          }
          await context.sync();

          result.tables = {
            total: tables.items.length,
            items: items.map((table, index) => {
              const values = Array.isArray(table.values) ? table.values : [];
              const previewRows = values.slice(0, 2).map((row) => row.slice(0, 3).join(" | "));
              return {
                address: `table[${index + 1}]`,
                rowCount: table.rowCount,
                columnCount: values.reduce((max, row) => Math.max(max, row.length), 0),
                bandedRows: table.styleBandedRows,
                firstColumnStyled: table.styleFirstColumn,
                lastColumnStyled: table.styleLastColumn,
                preview: includeText ? summarizePlainText(previewRows.join(" / "), 120) : undefined,
              };
            }),
          };
        }

        if (kind === "all" || kind === "contentControls") {
          const contentControls = context.document.contentControls;
          contentControls.load("items");
          await context.sync();

          const items = contentControls.items.slice(0, limit);
          for (const contentControl of items) {
            contentControl.load(["id", "title", "tag", "type", "appearance", "cannotEdit", "cannotDelete", "text"]);
          }
          await context.sync();

          result.contentControls = {
            total: contentControls.items.length,
            items: items.map((contentControl) => ({
              address: `content_control[id=${contentControl.id}]`,
              id: contentControl.id,
              title: contentControl.title || "",
              tag: contentControl.tag || "",
              type: contentControl.type,
              appearance: contentControl.appearance,
              cannotEdit: contentControl.cannotEdit,
              cannotDelete: contentControl.cannotDelete,
              preview: includeText ? summarizePlainText(contentControl.text || "", 120) : undefined,
            })),
          };
        }

        if (kind === "all" || kind === "bookmarks") {
          if (!isWordDesktopRequirementSetSupported("1.4")) {
            result.bookmarks = {
              available: false,
              message: "Bookmark inspection requires WordApiDesktop 1.4.",
            };
          } else {
            const bookmarks = context.document.bookmarks;
            bookmarks.load("items");
            await context.sync();

            const items = bookmarks.items.slice(0, limit);
            for (const bookmark of items) {
              bookmark.load(["name", "isEmpty", "storyType"]);
              if (includeText) {
                bookmark.range.load("text");
              }
            }
            await context.sync();

            result.bookmarks = {
              available: true,
              total: bookmarks.items.length,
              items: items.map((bookmark) => ({
                address: `bookmark[${bookmark.name}]`,
                name: bookmark.name,
                isEmpty: bookmark.isEmpty,
                storyType: bookmark.storyType,
                preview: includeText ? summarizePlainText(bookmark.range.text || "", 120) : undefined,
              })),
            };
          }
        }

        return result;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
