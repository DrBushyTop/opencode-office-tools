import type { Tool } from "./types";
import { z } from "zod";
import {
  DocumentContentFormatSchema,
  getZodErrorMessage,
  parseDocumentRangeAddress,
  readResolvedWordTarget,
  resolveDocumentRangeTarget,
  toolFailure,
} from "./wordShared";

const getDocumentRangeArgsSchema = z.object({
  address: z.string(),
  format: DocumentContentFormatSchema.optional().default("text"),
});

export type GetDocumentRangeArgs = z.infer<typeof getDocumentRangeArgsSchema>;

export const getDocumentRange: Tool = {
  name: "get_document_range",
  description: `Read a generic Word target by address.

Supported addresses:
- document or body
- selection
- bookmark[Name]
- content_control[id=12]
- content_control[index=1]
- table[1]
- table[1].cell[2,3]

Use format="html" for HTML symmetry with insertion, format="ooxml" for markup, format="text" for plain text, or format="summary" for a short preview.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: "Target address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3].",
      },
      format: {
        type: "string",
        enum: ["summary", "text", "html", "ooxml"],
        description: "Preferred response format. Default is text.",
      },
    },
    required: ["address"],
  },
  handler: async (args) => {
    const parsedArgs = getDocumentRangeArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { address, format } = parsedArgs.data;
    const parsed = parseDocumentRangeAddress(address);
    if (!parsed) {
      return toolFailure("Unsupported document range address. Try selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3].");
    }

    try {
      return await Word.run(async (context) => {
        const resolved = await resolveDocumentRangeTarget(context, parsed);
        return readResolvedWordTarget(context, resolved, format);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
