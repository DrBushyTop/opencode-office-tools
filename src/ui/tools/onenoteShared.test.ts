import { describe, expect, it } from "vitest";
import {
  formatZodError,
  formatPageSummary,
  formatPageText,
  navigateToPageArgsSchema,
  normalizeImagePayload,
  oneNoteContentSummarySchema,
  oneNoteParagraphSummarySchema,
  parsePagePlacement,
  parseSelectionFormat,
  parseSelectionWriteFormat,
} from "./onenoteShared";

describe("onenoteShared", () => {
  it("normalizes image payloads from data urls", () => {
    expect(normalizeImagePayload("data:image/png;base64,QUJDRA==")).toBe("QUJDRA==");
    expect(normalizeImagePayload("QUJDRA==")).toBe("QUJDRA==");
  });

  it("parses selection and placement defaults", () => {
    expect(parseSelectionFormat(undefined)).toBe("text");
    expect(parseSelectionFormat("matrix")).toBe("matrix");
    expect(parseSelectionWriteFormat(undefined)).toBe("text");
    expect(parseSelectionWriteFormat("html")).toBe("html");
    expect(parsePagePlacement(undefined)).toBe("sectionEnd");
    expect(parsePagePlacement("before")).toBe("before");
  });

  it("formats page text and summary", () => {
    const content = [{
      id: "content-1",
      type: "Outline",
      paragraphs: [
        { type: "RichText", text: "First paragraph" },
        { type: "Table", rowCount: 2, columnCount: 3 },
      ],
    }];

    expect(formatPageText(content)).toContain("First paragraph");
    expect(formatPageText(content)).toContain("[Table 2x3]");
    expect(formatPageSummary({ title: "Notes", id: "page-1", pageLevel: 0 }, content)).toContain("Notes");
  });

  it("defines zod schemas for OneNote-facing data", () => {
    expect(oneNoteParagraphSummarySchema.parse({ type: "RichText", text: "Hello" })).toEqual({ type: "RichText", text: "Hello" });
    expect(oneNoteContentSummarySchema.parse({
      id: "content-1",
      type: "Outline",
      paragraphs: [{ type: "RichText", text: "Hello" }],
    })).toEqual({
      id: "content-1",
      type: "Outline",
      paragraphs: [{ type: "RichText", text: "Hello" }],
    });
  });

  it("validates navigate_to_page args with zod", () => {
    const result = navigateToPageArgsSchema.safeParse({});
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(formatZodError(result.error)).toContain("Provide exactly one of pageId or clientUrl.");
    }
  });
});
