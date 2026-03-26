import type { Tool } from "./types";
import {
  getHeaderFooterBody,
  isValidSectionSelector,
  isWordDesktopRequirementSetSupported,
  parseDocumentPartAddress,
  summarizePlainText,
  toolFailure,
  type HeaderFooterTypeName,
} from "./wordShared";

const headerFooterTypes: HeaderFooterTypeName[] = ["primary", "firstPage", "evenPages"];

export const getDocumentPart: Tool = {
  name: "get_document_part",
  description: `Read a specific Word document part using an address.

Supported addresses:
- headers_footers
- section[1]
- section[*]
- section[1].header.primary
- section[2].footer.firstPage
- table_of_contents

Use format="text" or format="html" for specific header/footer bodies. Use format="summary" for overviews.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: "Document part address such as section[1].header.primary or table_of_contents.",
      },
      format: {
        type: "string",
        enum: ["summary", "text", "html"],
        description: "Preferred response format. Default is 'summary'.",
      },
    },
    required: ["address"],
  },
  handler: async (args) => {
    const { address, format = "summary" } = args as { address: string; format?: "summary" | "text" | "html" };
    const parsed = parseDocumentPartAddress(address);

    if (!parsed) {
      return toolFailure("Unsupported document part address. Try headers_footers, section[1].header.primary, section[*], or table_of_contents.");
    }

    try {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        if (parsed.kind === "headersFootersOverview") {
          const canReadPageSetup = isWordDesktopRequirementSetSupported("1.3");

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
        }

        if (parsed.kind === "tableOfContents") {
          if (!isWordDesktopRequirementSetSupported("1.4")) {
            return "Native table-of-contents inspection is unavailable on this Word host.";
          }

          const tablesOfContents = context.document.tablesOfContents;
          tablesOfContents.load("items");
          await context.sync();

          for (const item of tablesOfContents.items) {
            item.load(["upperHeadingLevel", "lowerHeadingLevel", "arePageNumbersIncluded", "arePageNumbersRightAligned"]);
          }
          await context.sync();

          if (tablesOfContents.items.length === 0) {
            return "No native tables of contents found in the document.";
          }

          return [
            "Tables Of Contents:",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            ...tablesOfContents.items.map((item, index) => `- TOC ${index + 1}: headings ${item.upperHeadingLevel}-${item.lowerHeadingLevel}, page numbers ${item.arePageNumbersIncluded ? "on" : "off"}, right aligned ${item.arePageNumbersRightAligned ? "on" : "off"}`),
          ].join("\n");
        }

        if (!isValidSectionSelector(parsed.section)) {
          return toolFailure("Section selector must be a positive integer or *.");
        }

        const targetSections = parsed.section === "*"
          ? sections.items.map((section, index) => ({ section, index }))
          : [{ section: sections.items[parsed.section - 1], index: parsed.section - 1 }].filter((item) => Boolean(item.section));

        if (targetSections.length === 0) {
          return toolFailure(`Section ${parsed.section} does not exist.`);
        }

        const canReadPageSetup = isWordDesktopRequirementSetSupported("1.3");
        if (!parsed.target) {
          for (const { section } of targetSections) {
            if (canReadPageSetup) {
              section.pageSetup.load([
                "differentFirstPageHeaderFooter",
                "oddAndEvenPagesHeaderFooter",
                "headerDistance",
                "footerDistance",
              ]);
            }
          }
          await context.sync();

          return targetSections.map(({ section, index }) => {
            if (!canReadPageSetup) {
              return `Section ${index + 1}: page setup details unavailable on this Word host.`;
            }

            const pageSetup = section.pageSetup;
            return [
              `Section ${index + 1}:`,
              `- Different first page: ${pageSetup.differentFirstPageHeaderFooter ? "on" : "off"}`,
              `- Different odd/even pages: ${pageSetup.oddAndEvenPagesHeaderFooter ? "on" : "off"}`,
              `- Header distance: ${pageSetup.headerDistance} pt`,
              `- Footer distance: ${pageSetup.footerDistance} pt`,
            ].join("\n");
          }).join("\n\n");
        }

        const bodies = targetSections.map(({ section, index }) => {
          const body = getHeaderFooterBody(section, parsed.target!, parsed.type || "primary");
          return { body, index };
        });

        if (format === "html") {
          const htmlResults = bodies.map(({ body }) => body.getHtml());
          await context.sync();
          return bodies.map(({ index }, i) => `Section ${index + 1} ${parsed.target}.${parsed.type || "primary"}:\n${htmlResults[i].value || "(empty)"}`).join("\n\n");
        }

        for (const entry of bodies) {
          entry.body.load("text");
        }
        await context.sync();

        if (format === "text" && bodies.length === 1) {
          return bodies[0].body.text || "(empty)";
        }

        return bodies.map(({ body, index }) => `Section ${index + 1} ${parsed.target}.${parsed.type || "primary"}: ${format === "text" ? (body.text || "(empty)") : summarizePlainText(body.text || "")}`).join("\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
