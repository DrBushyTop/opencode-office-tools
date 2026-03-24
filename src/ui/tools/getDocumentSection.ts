import type { Tool } from "@github/copilot-sdk";

export const getDocumentSection: Tool = {
  name: "get_document_section",
  description: `Get the content of a specific section of the Word document by heading.

Use this after get_document_overview to read a specific section without loading the entire document.

Parameters:
- headingText: The text of the heading to find (partial match supported)
- includeSubsections: If true, includes all content until the next heading of same or higher level (default: true)

Returns the HTML content of that section.

Examples:
- Get "Introduction" section: headingText="Introduction"
- Get "Chapter 2" section: headingText="Chapter 2"
- Get just heading content without subsections: headingText="Methods", includeSubsections=false`,
  parameters: {
    type: "object",
    properties: {
      headingText: {
        type: "string",
        description: "The heading text to search for (case-insensitive partial match).",
      },
      includeSubsections: {
        type: "boolean",
        description: "If true, includes content until the next heading of same or higher level. Default is true.",
      },
    },
    required: ["headingText"],
  },
  handler: async (args) => {
    const { headingText, includeSubsections = true } = args as {
      headingText: string;
      includeSubsections?: boolean;
    };

    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        // Load paragraph details
        for (const para of paragraphs.items) {
          para.load(["text", "style"]);
        }
        await context.sync();

        // Find the heading
        let startIndex = -1;
        let startLevel = 0;
        const searchLower = headingText.toLowerCase();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          const style = para.style || "";
          const text = (para.text || "").toLowerCase();

          // Check if this is a heading that matches
          const headingMatch = style.match(/Heading\s*(\d)/i);
          if (headingMatch && text.includes(searchLower)) {
            startIndex = i;
            startLevel = parseInt(headingMatch[1], 10);
            break;
          }
          // Also check Title style
          if ((style === "Title" || style === "Subtitle") && text.includes(searchLower)) {
            startIndex = i;
            startLevel = style === "Title" ? 1 : 2;
            break;
          }
        }

        if (startIndex === -1) {
          return `No heading found matching "${headingText}". Use get_document_overview to see available headings.`;
        }

        // Find the end of the section
        let endIndex = paragraphs.items.length;
        if (includeSubsections) {
          for (let i = startIndex + 1; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            const style = para.style || "";
            const headingMatch = style.match(/Heading\s*(\d)/i);
            if (headingMatch) {
              const level = parseInt(headingMatch[1], 10);
              if (level <= startLevel) {
                endIndex = i;
                break;
              }
            }
            if (style === "Title") {
              endIndex = i;
              break;
            }
          }
        } else {
          // Just get content until next heading of any level
          for (let i = startIndex + 1; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            const style = para.style || "";
            if (style.match(/Heading\s*\d/i) || style === "Title" || style === "Subtitle") {
              endIndex = i;
              break;
            }
          }
        }

        // Get the range from start to end
        const startPara = paragraphs.items[startIndex];
        const endPara = paragraphs.items[Math.min(endIndex, paragraphs.items.length) - 1];
        
        const range = startPara.getRange(Word.RangeLocation.whole);
        range.expandTo(endPara.getRange(Word.RangeLocation.whole));
        
        const html = range.getHtml();
        await context.sync();

        return html.value || "(empty section)";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
