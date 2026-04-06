import type { Tool } from "./types";
import { isWordDesktopRequirementSetSupported, toolFailure } from "./wordShared";

/** Map a Word style name to an outline level (1-6) or null. */
function headingLevel(style: string): number | null {
  if (style === "Title") return 1;
  if (style === "Subtitle") return 2;
  const m = /Heading\s*(\d)/i.exec(style);
  return m ? Number(m[1]) : null;
}

/** Truncate text to `max` chars, appending ellipsis when needed. */
function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) + "\u2026" : text;
}

/**
 * Produces a bird's-eye view of the active Word document:
 * word / paragraph / section / table / list / content-control counts
 * plus the heading outline tree.
 */
export const getDocumentOverview: Tool = {
  name: "get_document_overview",
  description: "Get a structural overview of the active Word document.",
  parameters: { type: "object", properties: {} },

  handler: async () => {
    try {
      return await Word.run(async (ctx) => {
        const doc = ctx.document;
        const body = doc.body;
        const sections = doc.sections;
        const tables = body.tables;
        const controls = body.contentControls;
        const paras = body.paragraphs;

        const hasTocApi = isWordDesktopRequirementSetSupported("1.4");
        const tocs = hasTocApi ? doc.tablesOfContents : null;

        body.load("text");
        sections.load("items");
        tables.load("items");
        controls.load("items");
        paras.load("items");
        if (tocs) tocs.load("items");

        await ctx.sync();

        // second batch: paragraph metadata
        for (const p of paras.items) p.load(["text", "style", "isListItem"]);
        await ctx.sync();

        // --- statistics ---
        const bodyText = body.text ?? "";
        const words = bodyText
          .trim()
          .split(/\s+/)
          .filter((w) => w !== "").length;

        let listItems = 0;
        const outline: string[] = [];

        for (const p of paras.items) {
          if (p.isListItem) listItems++;

          const lvl = headingLevel(p.style ?? "");
          if (lvl === null) continue;

          const indent = "  ".repeat(lvl - 1);
          const prefix = "#".repeat(Math.min(lvl, 4));
          outline.push(`${indent}${prefix} ${clip((p.text ?? "").trim(), 72)}`);
        }

        // --- assemble output ---
        const stats: string[] = [
          `words: ${words.toLocaleString()}`,
          `paragraphs: ${paras.items.length}`,
          `sections: ${sections.items.length}`,
          `tables: ${tables.items.length}`,
        ];

        if (tocs) stats.push(`tables of contents: ${tocs.items.length}`);
        if (listItems > 0) stats.push(`list items: ${listItems}`);
        if (controls.items.length > 0)
          stats.push(`content controls: ${controls.items.length}`);

        const parts: string[] = [
          "Document overview",
          "─".repeat(40),
          stats.join(" | "),
        ];

        if (outline.length > 0) {
          parts.push("", "Heading outline", "─".repeat(40), ...outline);
        } else {
          parts.push("", "(no headings detected)");
        }

        return parts.join("\n");
      });
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
