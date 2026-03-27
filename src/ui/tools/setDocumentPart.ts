import type { Tool } from "./types";
import {
  getHeaderFooterBody,
  isValidSectionSelector,
  isWordDesktopRequirementSetSupported,
  isWordRequirementSetSupported,
  parseDocumentPartAddress,
  toolFailure,
} from "./wordShared";

type Operation = "replace" | "append" | "clear" | "insert" | "configure";
type TocInsertLocation = "replace" | "before" | "after" | "start" | "end";

function resolveInsertLocation(location: TocInsertLocation) {
  switch (location) {
    case "before":
      return Word.InsertLocation.before;
    case "after":
      return Word.InsertLocation.after;
    case "start":
      return Word.InsertLocation.start;
    case "end":
      return Word.InsertLocation.end;
    case "replace":
    default:
      return Word.InsertLocation.replace;
  }
}

function resolveNativeInsertionRange(range: Word.Range, location: TocInsertLocation) {
  switch (location) {
    case "before":
    case "start":
      return range.getRange(Word.RangeLocation.start);
    case "after":
    case "end":
      return range.getRange(Word.RangeLocation.end);
    case "replace":
    default:
      return range;
  }
}

function describePlacement(location: TocInsertLocation) {
  switch (location) {
    case "before":
      return "before the current selection boundary";
    case "after":
      return "after the current selection boundary";
    case "start":
      return "at the start of the current selection";
    case "end":
      return "at the end of the current selection";
    case "replace":
    default:
      return "replacing the current selection";
  }
}

export const setDocumentPart: Tool = {
  name: "set_document_part",
  description: `Update a structural Word document part using an address.

Supported patterns:
- section[1].header.primary with replace, append, or clear
- section[2] or section[*] with configure page setup options
- table_of_contents with insert

Use set_document_range for generic body, selection, bookmark, content-control, and table edits.
Use flat options for section configuration and TOC insertion.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: "Document part address such as section[1].header.primary or table_of_contents.",
      },
      operation: {
        type: "string",
        enum: ["replace", "append", "clear", "insert", "configure"],
        description: "Operation to perform. Default is 'replace'.",
      },
      html: {
        type: "string",
        description: "HTML content to write when targeting a body-like part.",
      },
      differentFirstPage: {
        type: "boolean",
      },
      oddAndEvenPages: {
        type: "boolean",
      },
      headerDistance: {
        type: "number",
      },
      footerDistance: {
        type: "number",
      },
      location: {
        type: "string",
        enum: ["replace", "before", "after", "start", "end"],
      },
      upperHeadingLevel: {
        type: "number",
      },
      lowerHeadingLevel: {
        type: "number",
      },
      includePageNumbers: {
        type: "boolean",
      },
      rightAlignPageNumbers: {
        type: "boolean",
      },
      useHyperlinksOnWeb: {
        type: "boolean",
      },
    },
    required: ["address"],
  },
  handler: async (args) => {
    const {
      address,
      operation = "replace",
      html,
      differentFirstPage,
      oddAndEvenPages,
      headerDistance,
      footerDistance,
      location = "replace",
      upperHeadingLevel = 1,
      lowerHeadingLevel = 3,
      includePageNumbers = true,
      rightAlignPageNumbers = true,
      useHyperlinksOnWeb = true,
    } = args as {
      address: string;
      operation?: Operation;
      html?: string;
      differentFirstPage?: boolean;
      oddAndEvenPages?: boolean;
      headerDistance?: number;
      footerDistance?: number;
      location?: TocInsertLocation;
      upperHeadingLevel?: number;
      lowerHeadingLevel?: number;
      includePageNumbers?: boolean;
      rightAlignPageNumbers?: boolean;
      useHyperlinksOnWeb?: boolean;
    };

    const parsed = parseDocumentPartAddress(address);
    if (!parsed) {
      return toolFailure("Unsupported document part address. Try section[1].header.primary, section[*], or table_of_contents.");
    }

    if ((headerDistance !== undefined && (!Number.isFinite(headerDistance) || headerDistance < 0))
      || (footerDistance !== undefined && (!Number.isFinite(footerDistance) || footerDistance < 0))) {
      return toolFailure("headerDistance and footerDistance must be non-negative finite numbers.");
    }

    if (parsed.kind === "headersFootersOverview") {
      return toolFailure("headers_footers is read-only. Use section[...] addresses when writing.");
    }

    try {
      return await Word.run(async (context) => {
        if (parsed.kind === "tableOfContents") {
          if (operation !== "insert") {
            return toolFailure("table_of_contents currently supports only the insert operation.");
          }

          if (upperHeadingLevel < 1 || lowerHeadingLevel > 9 || upperHeadingLevel > lowerHeadingLevel) {
            return toolFailure("Heading levels must be between 1 and 9, and upperHeadingLevel must be less than or equal to lowerHeadingLevel.");
          }

          const selectionRange = context.document.getSelection().getRange();
          const requiresAdvancedOptions = upperHeadingLevel !== 1
            || lowerHeadingLevel !== 3
            || includePageNumbers !== true
            || rightAlignPageNumbers !== true
            || useHyperlinksOnWeb !== true;

          if (isWordDesktopRequirementSetSupported("1.4")) {
            const toc = context.document.tablesOfContents.add(resolveNativeInsertionRange(selectionRange, location), {
              upperHeadingLevel,
              lowerHeadingLevel,
              includePageNumbers,
              rightAlignPageNumbers,
              useBuiltInHeadingStyles: true,
              useHyperlinksOnWeb,
            });
            toc.load(["upperHeadingLevel", "lowerHeadingLevel"]);
            await context.sync();
            return `Inserted native table of contents for heading levels ${toc.upperHeadingLevel}-${toc.lowerHeadingLevel}, ${describePlacement(location)}.`;
          }

          if (!isWordRequirementSetSupported("1.5")) {
            return toolFailure("This Word host does not support native tables of contents or TOC fields via Office.js.");
          }

          if (requiresAdvancedOptions) {
            return toolFailure("Custom TOC options require a Word host with WordApiDesktop 1.4 native table-of-contents support.");
          }

          selectionRange.insertField(resolveInsertLocation(location), Word.FieldType.toc);
          await context.sync();
          return `Requested a basic TOC field, ${describePlacement(location)}. Word field support varies by host, especially on web clients.`;
        }

        if (!isValidSectionSelector(parsed.section)) {
          return toolFailure("Section selector must be a positive integer or *.");
        }

        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        const targetSections = parsed.section === "*"
          ? sections.items
          : [sections.items[parsed.section - 1]].filter(Boolean);

        if (targetSections.length === 0) {
          return toolFailure(`Section ${parsed.section} does not exist.`);
        }

        const hasConfigurationChange = [differentFirstPage, oddAndEvenPages, headerDistance, footerDistance].some((value) => value !== undefined);

        if (parsed.target) {
          if (operation === "insert" || operation === "configure") {
            return toolFailure(`The ${address} target supports replace, append, or clear.`);
          }

          if (hasConfigurationChange) {
            return toolFailure(`Section configuration must target ${parsed.section === "*" ? "section[*]" : `section[${parsed.section}]`} directly, not ${address}.`);
          }

          if ((operation === "replace" || operation === "append") && html === undefined) {
            return toolFailure("html is required for replace or append operations on header/footer targets.");
          }
        } else {
          if (operation !== "configure") {
            return toolFailure(`The ${address} target only supports configure.`);
          }
          if (!hasConfigurationChange) {
            return toolFailure("Provide at least one section configuration value to update.");
          }
        }

        if (operation === "configure") {
          if (!isWordDesktopRequirementSetSupported("1.3")) {
            return toolFailure("Section page setup configuration requires WordApiDesktop 1.3.");
          }

          for (const section of targetSections) {
            const pageSetup = section.pageSetup;
            if (differentFirstPage !== undefined) pageSetup.differentFirstPageHeaderFooter = differentFirstPage;
            if (oddAndEvenPages !== undefined) pageSetup.oddAndEvenPagesHeaderFooter = oddAndEvenPages;
            if (headerDistance !== undefined) pageSetup.headerDistance = headerDistance;
            if (footerDistance !== undefined) pageSetup.footerDistance = footerDistance;
          }
        }

        if (parsed.target) {
          for (const section of targetSections) {
            const body = getHeaderFooterBody(section, parsed.target, parsed.type || "primary");
            if (operation === "clear") {
              body.clear();
            } else if (operation === "replace") {
              body.clear();
              if (html) body.insertHtml(html, Word.InsertLocation.start);
            } else if (operation === "append" && html) {
              body.insertHtml(html, Word.InsertLocation.end);
            }
          }
        }

        await context.sync();

        if (parsed.target) {
          return `Updated ${address} with ${operation} across ${targetSections.length} section${targetSections.length === 1 ? "" : "s"}.`;
        }
        return `Updated section configuration for ${address}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
