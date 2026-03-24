import type { Tool } from "@github/copilot-sdk";

export const findAndReplace: Tool = {
  name: "find_and_replace",
  description: `Find and replace text in the Word document.

Searches the entire document for occurrences of the search text and replaces them.

Parameters:
- find: The text to search for
- replace: The text to replace it with
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
      matchCase?: boolean;
      matchWholeWord?: boolean;
    };
    
    if (!find || find.length === 0) {
      return { textResultForLlm: "Search text cannot be empty.", resultType: "failure", error: "Empty search", toolTelemetry: {} };
    }
    
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        
        // Create search options
        const searchResults = body.search(find, {
          ignorePunct: false,
          ignoreSpace: false,
          matchCase: matchCase,
          matchWholeWord: matchWholeWord,
        });
        
        searchResults.load("items");
        await context.sync();
        
        const count = searchResults.items.length;
        
        if (count === 0) {
          return `No matches found for "${find}".`;
        }
        
        // Replace all matches
        for (const result of searchResults.items) {
          result.insertText(replace, Word.InsertLocation.replace);
        }
        
        await context.sync();
        
        return `Replaced ${count} occurrence${count === 1 ? "" : "s"} of "${find}" with "${replace}".`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
