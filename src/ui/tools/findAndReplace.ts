import type { Tool } from "./types";
import { parseDocumentRangeAddress, resolveDocumentRangeTarget, toolFailure } from "./wordShared";

export const findAndReplace: Tool = {
  name: "find_and_replace",
  description: `Find and replace text in Word.

 Searches the entire document by default, or a generic target scope when address is provided.

Parameters:
- find: The text to search for
- replace: The text to replace it with
- address: Optional scope such as selection, bookmark[Name], content_control[id=12], or table[1]
- matchCase: If true, search is case-sensitive (default: false)
- matchWholeWord: If true, only match whole words, not partial matches (default: false)

Returns the number of replacements made.

Examples:
- Replace all "colour" with "color": find="colour", replace="color"
- Replace exact case "JavaScript": find="JavaScript", replace="TypeScript", matchCase=true
- Replace whole word "cat" (not "category"): find="cat", replace="dog", matchWholeWord=true`,
  parameters: {
    type: "object",
    properties: {
      find: {
        type: "string",
        description: "The text to search for.",
      },
      replace: {
        type: "string",
        description: "The text to replace matches with.",
      },
      address: {
        type: "string",
        description: "Optional scope address such as selection, bookmark[Clause], content_control[id=12], or table[1].",
      },
      matchCase: {
        type: "boolean",
        description: "If true, the search is case-sensitive. Default is false.",
      },
      matchWholeWord: {
        type: "boolean",
        description: "If true, only matches whole words. Default is false.",
      },
    },
    required: ["find", "replace"],
  },
  handler: async (args) => {
    const { find, replace, matchCase = false, matchWholeWord = false } = args as {
      find: string;
      replace: string;
      address?: string;
      matchCase?: boolean;
      matchWholeWord?: boolean;
    };
    const { address } = args as { address?: string };
    
    if (!find || find.length === 0) {
      return toolFailure("Search text cannot be empty.");
    }

    const parsed = address ? parseDocumentRangeAddress(address) : null;
    if (address && !parsed) {
      return toolFailure("Unsupported scope address. Try selection, bookmark[Name], content_control[id=12], or table[1].");
    }
    
    try {
      return await Word.run(async (context) => {
        const target = parsed
          ? await resolveDocumentRangeTarget(context, parsed)
          : { kind: "body" as const, label: "document", target: context.document.body };
        
        // Create search options
        const searchResults = target.target.search(find, {
          ignorePunct: false,
          ignoreSpace: false,
          matchCase: matchCase,
          matchWholeWord: matchWholeWord,
        });
        
        searchResults.load("items");
        await context.sync();
        
        const count = searchResults.items.length;
        
        if (count === 0) {
          return `No matches found for "${find}" in ${target.label}.`;
        }
        
        // Replace all matches
        for (const result of searchResults.items) {
          result.insertText(replace, Word.InsertLocation.replace);
        }
        
        await context.sync();
        
        return `Replaced ${count} occurrence${count === 1 ? "" : "s"} of "${find}" with "${replace}" in ${target.label}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
