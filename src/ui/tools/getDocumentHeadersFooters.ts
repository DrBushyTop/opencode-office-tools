import type { Tool } from "./types";
import {
  getHeaderFooterBody,
  isWordDesktopRequirementSetSupported,
  summarizePlainText,
  toolFailure,
  type HeaderFooterTypeName,
} from "./wordShared";

const headerFooterTypes: HeaderFooterTypeName[] = ["primary", "firstPage", "evenPages"];

export const getDocumentHeadersFooters: Tool = {
  name: "get_document_headers_footers",
  description: `Inspect Word section header/footer configuration.

Returns:
- section count
- first-page and even-page header/footer settings
- header/footer distances
- text previews for each section's header and footer bodies`,
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        const canReadTocs = isWordDesktopRequirementSetSupported("1.4");
        const canReadPageSetup = isWordDesktopRequirementSetSupported("1.3");
        const tablesOfContents = canReadTocs ? context.document.tablesOfContents : null;

        sections.load("items");
        if (tablesOfContents) {
          tablesOfContents.load("items");
        }
        await context.sync();

        const bodies: Array<{ label: string; body: Word.Body }> = [];
        for (let i = 0; i < sections.items.length; i += 1) {
          const section = sections.items[i];
          if (canReadPageSetup) {
            section.pageSetup.load([
              "differentFirstPageHeaderFooter",
              "oddAndEvenPagesHeaderFooter",
              "headerDistance",
              "footerDistance",
            ]);
          }

          for (const type of headerFooterTypes) {
            bodies.push({ label: `Section ${i + 1} ${type} header`, body: getHeaderFooterBody(section, "header", type) });
            bodies.push({ label: `Section ${i + 1} ${type} footer`, body: getHeaderFooterBody(section, "footer", type) });
          }
        }

        for (const entry of bodies) {
          entry.body.load("text");
        }
        await context.sync();

        const lines: string[] = [
          "Document Header/Footer Overview:",
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
          `Sections: ${sections.items.length}`,
        ];

        if (tablesOfContents) {
          lines.push(`Tables of contents: ${tablesOfContents.items.length}`);
        }

        for (let i = 0; i < sections.items.length; i += 1) {
          const section = sections.items[i];

          lines.push("", `Section ${i + 1}:`);
          if (canReadPageSetup) {
            const pageSetup = section.pageSetup;
            lines.push(`- Different first page: ${pageSetup.differentFirstPageHeaderFooter ? "on" : "off"}`);
            lines.push(`- Different odd/even pages: ${pageSetup.oddAndEvenPagesHeaderFooter ? "on" : "off"}`);
            lines.push(`- Header distance: ${pageSetup.headerDistance} pt`);
            lines.push(`- Footer distance: ${pageSetup.footerDistance} pt`);
          } else {
            lines.push("- Page setup details unavailable on this Word host");
          }

          for (const type of headerFooterTypes) {
            const header = bodies.find((entry) => entry.label === `Section ${i + 1} ${type} header`);
            const footer = bodies.find((entry) => entry.label === `Section ${i + 1} ${type} footer`);
            lines.push(`- ${type} header: ${summarizePlainText(header?.body.text || "")}`);
            lines.push(`- ${type} footer: ${summarizePlainText(footer?.body.text || "")}`);
          }
        }

        return lines.join("\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
