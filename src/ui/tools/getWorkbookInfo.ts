import type { Tool } from "./types";

/**
 * Lightweight workbook introspection — returns sheet listing
 * and which sheet the user currently has open. For richer
 * metadata (tables, named ranges, charts) use get_workbook_overview.
 */
export const getWorkbookInfo: Tool = {
  name: "get_workbook_info",
  description:
    "Get a lightweight workbook summary with worksheet names and the active sheet. Prefer get_workbook_overview for structural inspection.",
  parameters: { type: "object", properties: {} },

  handler: async () => {
    try {
      return await Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        const active = ctx.workbook.worksheets.getActiveWorksheet();

        sheets.load("items/name,items/position");
        active.load("name");
        await ctx.sync();

        const ordered = [...sheets.items].sort(
          (a, b) => a.position - b.position,
        );

        const lines: string[] = [
          `Sheets (${ordered.length}) — active: ${active.name}`,
          "",
        ];

        for (const s of ordered) {
          const marker = s.name === active.name ? " *" : "";
          lines.push(`  ${s.position + 1}. ${s.name}${marker}`);
        }

        return lines.join("\n");
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return {
        resultType: "failure",
        error: msg,
        textResultForLlm: `Unable to read workbook info: ${msg}`,
        toolTelemetry: {},
      };
    }
  },
};
