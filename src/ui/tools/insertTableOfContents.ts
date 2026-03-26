import type { Tool } from "./types";
import { isWordDesktopRequirementSetSupported, isWordRequirementSetSupported, toolFailure } from "./wordShared";

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

export const insertTableOfContents: Tool = {
  name: "insert_table_of_contents",
  description: `Insert a real Word table of contents at the current selection.

Uses the native tables-of-contents API when available and falls back to a TOC field on hosts that don't expose the newer collection API.`,
  parameters: {
    type: "object",
    properties: {
      location: {
        type: "string",
        enum: ["replace", "before", "after", "start", "end"],
        description: "Where to insert the table of contents relative to the current selection. Default is 'replace'.",
      },
      upperHeadingLevel: {
        type: "number",
        description: "Starting heading level to include. Default is 1.",
      },
      lowerHeadingLevel: {
        type: "number",
        description: "Ending heading level to include. Default is 3.",
      },
      includePageNumbers: {
        type: "boolean",
        description: "Whether to include page numbers. Default is true.",
      },
      rightAlignPageNumbers: {
        type: "boolean",
        description: "Whether to right-align page numbers. Default is true.",
      },
      useHyperlinksOnWeb: {
        type: "boolean",
        description: "Whether TOC entries should become hyperlinks when published to the web.",
      },
    },
  },
  handler: async (args) => {
    const {
      location = "replace",
      upperHeadingLevel = 1,
      lowerHeadingLevel = 3,
      includePageNumbers = true,
      rightAlignPageNumbers = true,
      useHyperlinksOnWeb = true,
    } = args as {
      location?: TocInsertLocation;
      upperHeadingLevel?: number;
      lowerHeadingLevel?: number;
      includePageNumbers?: boolean;
      rightAlignPageNumbers?: boolean;
      useHyperlinksOnWeb?: boolean;
    };

    if (upperHeadingLevel < 1 || lowerHeadingLevel > 9 || upperHeadingLevel > lowerHeadingLevel) {
      return toolFailure("Heading levels must be between 1 and 9, and upperHeadingLevel must be less than or equal to lowerHeadingLevel.");
    }

    const requiresAdvancedOptions = upperHeadingLevel !== 1
      || lowerHeadingLevel !== 3
      || includePageNumbers !== true
      || rightAlignPageNumbers !== true
      || useHyperlinksOnWeb !== true;

    try {
      return await Word.run(async (context) => {
        const selectionRange = context.document.getSelection().getRange();

        if (isWordDesktopRequirementSetSupported("1.4")) {
          const toc = context.document.tablesOfContents.add(resolveNativeInsertionRange(selectionRange, location), {
            upperHeadingLevel,
            lowerHeadingLevel,
            includePageNumbers,
            rightAlignPageNumbers,
            useBuiltInHeadingStyles: true,
            useHyperlinksOnWeb,
          });
          toc.load(["upperHeadingLevel", "lowerHeadingLevel", "arePageNumbersIncluded"]);
          await context.sync();
          return `Inserted native table of contents for heading levels ${toc.upperHeadingLevel}-${toc.lowerHeadingLevel} using ${location} placement.`;
        }

        if (!isWordRequirementSetSupported("1.5")) {
          return toolFailure("This Word host does not support native tables of contents or TOC fields via Office.js.");
        }

        if (requiresAdvancedOptions) {
          return toolFailure("Custom TOC options require a Word host with WordApiDesktop 1.4 native table-of-contents support.");
        }

        selectionRange.insertField(resolveInsertLocation(location), Word.FieldType.toc);
        await context.sync();
        return `Inserted a basic TOC field at the current selection using ${location} placement.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
