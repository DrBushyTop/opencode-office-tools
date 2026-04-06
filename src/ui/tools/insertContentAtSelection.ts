import type { Tool } from "./types";
import { z } from "zod";
import {
  DocumentWriteLocationSchema,
  getZodErrorMessage,
  resolveDocumentRangeTarget,
  toolFailure,
  writeResolvedWordTarget,
} from "./wordShared";

const argsSchema = z.object({
  html: z.string().min(1, "html is required"),
  location: DocumentWriteLocationSchema.optional().default("replace"),
});

/**
 * Writes HTML into the active Word selection. The `location` parameter
 * controls placement relative to the selection (replace, before, after,
 * start, end).
 */
export const insertContentAtSelection: Tool = {
  name: "insert_content_at_selection",
  description:
    "Insert HTML content at the current Word selection. " +
    "Use location to control placement: replace (default), before, after, start, end.",
  parameters: {
    type: "object",
    properties: {
      html: {
        type: "string",
        description: "HTML markup to insert.",
      },
      location: {
        type: "string",
        enum: ["replace", "before", "after", "start", "end"],
        description:
          "Placement relative to the current selection (defaults to replace).",
      },
    },
    required: ["html"],
  },

  handler: async (args) => {
    const parsed = argsSchema.safeParse(args ?? {});
    if (!parsed.success) return toolFailure(getZodErrorMessage(parsed.error));

    const { html, location } = parsed.data;

    try {
      return await Word.run(async (ctx) => {
        const target = await resolveDocumentRangeTarget(ctx, {
          kind: "selection",
        });
        writeResolvedWordTarget(target, "insert", "html", html, location);
        await ctx.sync();

        return `Inserted at selection (${location}).`;
      });
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
