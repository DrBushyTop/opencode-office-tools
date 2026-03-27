import type { Tool } from "./types";
import {
  parseDocumentRangeAddress,
  resolveDocumentRangeTarget,
  toolFailure,
  writeResolvedWordTarget,
  type DocumentWriteFormat,
  type DocumentWriteLocation,
  type DocumentWriteOperation,
} from "./wordShared";

export const setDocumentRange: Tool = {
  name: "set_document_range",
  description: `Update a generic Word target by address.

Supported addresses:
- document or body
- selection
- bookmark[Name]
- content_control[id=12]
- content_control[index=1]
- table[1]
- table[1].cell[2,3]

Use this for generic content edits. Keep set_document_part for section headers, footers, page setup, and native tables of contents.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: "Target address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3].",
      },
      operation: {
        type: "string",
        enum: ["replace", "insert", "clear"],
        description: "Operation to perform. Default is replace.",
      },
      format: {
        type: "string",
        enum: ["html", "text", "ooxml"],
        description: "Content format for replace or insert. Default is html.",
      },
      content: {
        type: "string",
        description: "Content to write for replace or insert operations. Required unless operation is clear.",
      },
      location: {
        type: "string",
        enum: ["replace", "before", "after", "start", "end"],
        description: "Insertion location for insert operations. Default is replace.",
      },
    },
    required: ["address"],
  },
  handler: async (args) => {
    const {
      address,
      operation = "replace",
      format = "html",
      content,
      location = "replace",
    } = args as {
      address: string;
      operation?: DocumentWriteOperation;
      format?: DocumentWriteFormat;
      content?: string;
      location?: DocumentWriteLocation;
    };

    const parsed = parseDocumentRangeAddress(address);
    if (!parsed) {
      return toolFailure("Unsupported document range address. Try selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3].");
    }

    if ((operation === "replace" || operation === "insert") && content === undefined) {
      return toolFailure("content is required for replace or insert operations.");
    }

    try {
      return await Word.run(async (context) => {
        const resolved = await resolveDocumentRangeTarget(context, parsed);
        writeResolvedWordTarget(resolved, operation, format, content, location);
        await context.sync();

        if (operation === "clear") {
          return `Cleared ${resolved.label}.`;
        }

        return `${operation === "replace" ? "Updated" : "Inserted into"} ${resolved.label} using ${format}${operation === "insert" ? ` at ${location}` : ""}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
