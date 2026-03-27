import type { Tool } from "./types";
import {
  normalizeImagePayload,
  parseSelectionWriteFormat,
  setSelectedDataAsync,
  toolFailure,
} from "./onenoteShared";

export const setNoteSelection: Tool = {
  name: "set_note_selection",
  description: "Write text, HTML, or an image to the current OneNote selection.",
  parameters: {
    type: "object",
    properties: {
      content: {
        type: "string",
        description: "Content to insert into the current selection.",
      },
      coercionType: {
        type: "string",
        enum: ["text", "html", "image"],
        description: "How to treat the provided content. Default is text.",
      },
    },
    required: ["content"],
  },
  handler: async (args) => {
    const { content } = args as { content: string; coercionType?: string };
    const coercionType = parseSelectionWriteFormat((args as { coercionType?: string }).coercionType);
    const trimmed = String(content || "").trim();

    if (!trimmed) {
      return toolFailure("Content cannot be empty.");
    }

    try {
      if (coercionType === "html") {
        await setSelectedDataAsync(content, Office.CoercionType.Html);
        return "Updated the current OneNote selection with HTML content.";
      }

      if (coercionType === "image") {
        await setSelectedDataAsync(normalizeImagePayload(content), Office.CoercionType.Image);
        return "Inserted an image into the current OneNote selection.";
      }

      await setSelectedDataAsync(content, Office.CoercionType.Text);
      return "Updated the current OneNote selection with text.";
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
