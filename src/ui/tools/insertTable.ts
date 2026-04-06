import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const insertTableStyleSchema = z.enum(["grid", "striped", "plain"]);
const insertTableArgsSchema = z.object({
  data: z.array(z.array(z.string())).min(1),
  hasHeader: z.boolean().optional().default(true),
  style: insertTableStyleSchema.optional().default("grid"),
});

export type InsertTableArgs = z.infer<typeof insertTableArgsSchema>;

function countTableColumns(data: string[][]) {
  return Math.max(...data.map((row) => row.length));
}

function resolveBuiltInTableStyle(style: InsertTableArgs["style"]) {
  if (style === "grid") return "TableStyleLight1";
  if (style === "striped") return "TableStyleLight3";
  return null;
}

async function loadRowCells(context: Word.RequestContext, rows: Word.TableRow[]) {
  for (const row of rows) {
    row.load("cells");
  }
  await context.sync();
}

async function tintHeaderRow(context: Word.RequestContext, row: Word.TableRow) {
  await loadRowCells(context, [row]);

  for (const cell of row.cells.items) {
    cell.shadingColor = "#4472C4";
    cell.body.load("paragraphs");
  }
  await context.sync();

  for (const cell of row.cells.items) {
    for (const paragraph of cell.body.paragraphs.items) {
      paragraph.load("font");
    }
  }
  await context.sync();

  for (const cell of row.cells.items) {
    for (const paragraph of cell.body.paragraphs.items) {
      paragraph.font.bold = true;
      paragraph.font.color = "#FFFFFF";
    }
  }
}

async function shadeAlternateRows(context: Word.RequestContext, rows: Word.TableRow[], offset: number) {
  const rowsToShade = rows.filter((_, index) => index % 2 === offset);
  if (!rowsToShade.length) return;

  await loadRowCells(context, rowsToShade);
  for (const row of rowsToShade) {
    for (const cell of row.cells.items) {
      cell.shadingColor = "#E8E8E8";
    }
  }
}

export const insertTable: Tool = {
  name: "insert_table",
  description: `Insert a Word table immediately after the current selection.

Parameters:
- data: 2D array of strings for the table body.
- hasHeader: When true, apply the header styling treatment to the first row.
- style: Choose "grid", "striped", or "plain".

Examples:
- Simple table with headers:
  data = [["Name", "Age", "City"], ["Alice", "30", "NYC"], ["Bob", "25", "LA"]]
  
- Data table without headers:
  data = [["Q1", "$100"], ["Q2", "$150"], ["Q3", "$200"]]
  hasHeader = false`,
  parameters: {
    type: "object",
    properties: {
      data: {
        type: "array",
        items: {
          type: "array",
          items: { type: "string" },
        },
        description: "2D array of cell values. Each inner array is a row.",
      },
      hasHeader: {
        type: "boolean",
        description: "If true, style the first row as headers. Default is true.",
      },
      style: {
        type: "string",
        enum: ["grid", "striped", "plain"],
        description: "Table style. Default is 'grid'.",
      },
    },
    required: ["data"],
  },
  handler: async (args) => {
    const parsedArgs = insertTableArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { data, hasHeader, style } = parsedArgs.data;

    const rowCount = data.length;
    const colCount = countTableColumns(data);

    if (colCount === 0) {
      return toolFailure("Table must have at least one column.");
    }

    try {
      return await Word.run(async (context) => {
        const table = context.document.getSelection().insertTable(rowCount, colCount, Word.InsertLocation.after, data);
        table.load("rows");
        await context.sync();

        const builtInStyle = resolveBuiltInTableStyle(style);
        if (builtInStyle) {
          table.style = builtInStyle;
        }

        if (hasHeader && table.rows.items.length > 0) {
          await tintHeaderRow(context, table.rows.items[0]);
        }

        if (style === "striped") {
          await shadeAlternateRows(context, table.rows.items.slice(hasHeader ? 1 : 0), hasHeader ? 0 : 1);
        }

        await context.sync();

        return `Inserted ${rowCount}x${colCount} table with ${style} style${hasHeader ? " and header row" : ""}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
