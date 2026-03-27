import { describe, expect, it } from "vitest";
import {
  formatPageSummary,
  formatPageText,
  normalizeImagePayload,
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
});
