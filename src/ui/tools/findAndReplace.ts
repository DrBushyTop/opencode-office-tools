import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, parseDocumentRangeAddress, resolveDocumentRangeTarget, toolFailure } from "./wordShared";

const findAndReplaceArgsSchema = z.object({
  find: z.string(),
  replace: z.string(),
  address: z.string().optional(),
  matchCase: z.boolean().optional().default(false),
  matchWholeWord: z.boolean().optional().default(false),
});

export type FindAndReplaceArgs = z.infer<typeof findAndReplaceArgsSchema>;

function resolveSearchAddress(address?: string) {
  if (!address) return { ok: true as const, value: null };

  const parsed = parseDocumentRangeAddress(address);
  if (parsed) {
    return { ok: true as const, value: parsed };
  }

  return {
    ok: false as const,
    failure: toolFailure("Unsupported scope address. Try selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3]."),
  };
}

function buildSearchOptions(matchCase: boolean, matchWholeWord: boolean) {
  return {
    ignorePunct: false,
    ignoreSpace: false,
    matchCase,
    matchWholeWord,
  };
}

function formatReplacementSummary(find: string, replace: string, label: string, count: number) {
  if (count === 0) {
    return `No text matching "${find}" was found in ${label}.`;
  }

  return `Updated ${count} occurrence${count === 1 ? "" : "s"} in ${label}, replacing "${find}" with "${replace}".`;
}

function replaceSearchResults(results: Word.RangeCollection, replace: string) {
  for (const result of results.items) {
    result.insertText(replace, Word.InsertLocation.replace);
  }
}

export const findAndReplace: Tool = {
  name: "find_and_replace",
  description: `Replace matching text in a Word document or in a narrower target such as the selection, a bookmark, a content control, or a single table cell.

 Searches the full document when no address is supplied.

Parameters:
- find: The text to search for
- replace: The text to replace it with
- address: Optional scope such as selection, bookmark[Name], content_control[id=12], table[1], or table[1].cell[2,3]
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
        description: "Optional scope address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3].",
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
    const parsedArgs = findAndReplaceArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { find, replace, address, matchCase, matchWholeWord } = parsedArgs.data;
    
    if (!find.trim()) {
      return toolFailure("Search text cannot be empty.");
    }

    const resolvedAddress = resolveSearchAddress(address);
    if (!resolvedAddress.ok) return resolvedAddress.failure;
    
    try {
      return await Word.run(async (context) => {
        const target = resolvedAddress.value
          ? await resolveDocumentRangeTarget(context, resolvedAddress.value)
          : { kind: "body" as const, label: "document", target: context.document.body };

        const searchResults = target.target.search(find, buildSearchOptions(matchCase, matchWholeWord));
        
        searchResults.load("items");
        await context.sync();
        const matches = searchResults.items;

        replaceSearchResults(searchResults, replace);

        if (matches.length === 0) {
          return formatReplacementSummary(find, replace, target.label, 0);
        }

        await context.sync();

        return formatReplacementSummary(find, replace, target.label, matches.length);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
