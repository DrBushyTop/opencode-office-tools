import type { Tool } from "./types";
import {
  getHeaderFooterBody,
  isWordDesktopRequirementSetSupported,
  toolFailure,
  type HeaderFooterTarget,
  type HeaderFooterTypeName,
} from "./wordShared";

type Mode = "replace" | "append";

export const setSectionHeaderFooter: Tool = {
  name: "set_section_header_footer",
  description: `Configure Word headers or footers for one section or all sections.

Supports:
- replacing or appending HTML content in a header/footer body
- toggling different first-page headers/footers
- toggling odd/even page headers/footers
- updating header/footer distance in points

If sectionIndex is omitted, section 1 is used unless applyToAllSections is true.`,
  parameters: {
    type: "object",
    properties: {
      target: {
        type: "string",
        enum: ["header", "footer"],
        description: "Which body to update.",
      },
      type: {
        type: "string",
        enum: ["primary", "firstPage", "evenPages"],
        description: "Which header/footer variant to edit. Default is 'primary'.",
      },
      html: {
        type: "string",
        description: "HTML content to insert into the chosen header/footer.",
      },
      mode: {
        type: "string",
        enum: ["replace", "append"],
        description: "Whether to replace existing content or append to it. Default is 'replace'.",
      },
      sectionIndex: {
        type: "number",
        description: "1-based section number to target.",
      },
      applyToAllSections: {
        type: "boolean",
        description: "Apply the same change to every section.",
      },
      differentFirstPage: {
        type: "boolean",
        description: "Whether the section should use a different first-page header/footer.",
      },
      oddAndEvenPages: {
        type: "boolean",
        description: "Whether odd and even pages should use different headers/footers.",
      },
      headerDistance: {
        type: "number",
        description: "Distance from the top of the page to the header, in points.",
      },
      footerDistance: {
        type: "number",
        description: "Distance from the bottom of the page to the footer, in points.",
      },
    },
    required: ["target"],
  },
  handler: async (args) => {
    const {
      target,
      type = "primary",
      html,
      mode = "replace",
      sectionIndex,
      applyToAllSections = false,
      differentFirstPage,
      oddAndEvenPages,
      headerDistance,
      footerDistance,
    } = args as {
      target: HeaderFooterTarget;
      type?: HeaderFooterTypeName;
      html?: string;
      mode?: Mode;
      sectionIndex?: number;
      applyToAllSections?: boolean;
      differentFirstPage?: boolean;
      oddAndEvenPages?: boolean;
      headerDistance?: number;
      footerDistance?: number;
    };

    const hasHtmlChange = html !== undefined;
    const hasConfigurationChange = [differentFirstPage, oddAndEvenPages, headerDistance, footerDistance].some((value) => value !== undefined);
    if (!hasHtmlChange && !hasConfigurationChange) {
      return toolFailure("Provide html content and/or header/footer configuration values.");
    }

    if ((headerDistance !== undefined && (!Number.isFinite(headerDistance) || headerDistance < 0))
      || (footerDistance !== undefined && (!Number.isFinite(footerDistance) || footerDistance < 0))) {
      return toolFailure("headerDistance and footerDistance must be non-negative finite numbers.");
    }

    try {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        if (sections.items.length === 0) {
          return toolFailure("The document does not contain any sections.");
        }

        if (hasConfigurationChange && !isWordDesktopRequirementSetSupported("1.3")) {
          return toolFailure("This Word host can edit header/footer content, but page setup configuration requires WordApiDesktop 1.3.");
        }

        let targetSections: Word.Section[];
        if (applyToAllSections) {
          targetSections = sections.items;
        } else {
          const resolvedIndex = Math.max(1, sectionIndex || 1) - 1;
          const section = sections.items[resolvedIndex];
          if (!section) {
            return toolFailure(`Section ${resolvedIndex + 1} does not exist.`);
          }
          targetSections = [section];
        }

        for (const section of targetSections) {
          const pageSetup = section.pageSetup;

          if (differentFirstPage !== undefined) {
            pageSetup.differentFirstPageHeaderFooter = differentFirstPage;
          }
          if (oddAndEvenPages !== undefined) {
            pageSetup.oddAndEvenPagesHeaderFooter = oddAndEvenPages;
          }
          if (type === "firstPage") {
            pageSetup.differentFirstPageHeaderFooter = true;
          }
          if (type === "evenPages") {
            pageSetup.oddAndEvenPagesHeaderFooter = true;
          }
          if (headerDistance !== undefined) {
            pageSetup.headerDistance = headerDistance;
          }
          if (footerDistance !== undefined) {
            pageSetup.footerDistance = footerDistance;
          }

          if (hasHtmlChange) {
            const body = getHeaderFooterBody(section, target, type);
            if (mode === "replace") {
              body.clear();
              if (html) {
                body.insertHtml(html, Word.InsertLocation.start);
              }
            } else if (html) {
              body.insertHtml(html, Word.InsertLocation.end);
            }
          }
        }

        await context.sync();

        const scope = applyToAllSections ? `all ${targetSections.length} sections` : `section ${sections.items.indexOf(targetSections[0]) + 1}`;
        const action = hasHtmlChange
          ? `${mode === "replace" ? (html ? "updated" : "cleared") : "appended to"} the ${type} ${target}`
          : `updated ${target} configuration`;
        return `Successfully ${action} for ${scope}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
