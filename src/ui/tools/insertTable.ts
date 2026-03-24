import type { Tool } from "@github/copilot-sdk";

export const insertTable: Tool = {
  name: "insert_table",
  description: `Insert a table at the current cursor position in Word.

Parameters:
- data: 2D array of strings representing the table data. First row can be headers.
- hasHeader: If true, the first row is styled as a header row (bold, shaded). Default is true.
- style: Table style - "grid" (borders), "striped" (alternating rows), or "plain" (minimal). Default is "grid".

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
    const { data, hasHeader = true, style = "grid" } = args as {
      data: string[][];
      hasHeader?: boolean;
      style?: string;
    };

    if (!data || data.length === 0) {
      return { textResultForLlm: "Table data cannot be empty.", resultType: "failure", error: "Empty data", toolTelemetry: {} };
    }

    const rowCount = data.length;
    const colCount = Math.max(...data.map(row => row.length));

    if (colCount === 0) {
      return { textResultForLlm: "Table must have at least one column.", resultType: "failure", error: "No columns", toolTelemetry: {} };
    }

    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        
        // Insert table at selection
        const table = selection.insertTable(rowCount, colCount, Word.InsertLocation.after, data);
        table.load("rows");
        await context.sync();

        // Apply styling
        if (style === "grid" || style === "striped") {
          table.setBorderColor("CCCCCC");
        }

        // Style header row
        if (hasHeader && table.rows.items.length > 0) {
          const headerRow = table.rows.items[0];
          headerRow.load("cells");
          await context.sync();
          
          for (const cell of headerRow.cells.items) {
            cell.shadingColor = "#4472C4";
            const cellBody = cell.body;
            cellBody.load("paragraphs");
            await context.sync();
            
            for (const para of cellBody.paragraphs.items) {
              para.load("font");
              await context.sync();
              para.font.bold = true;
              para.font.color = "#FFFFFF";
            }
          }
        }

        // Apply striped styling
        if (style === "striped") {
          for (let i = hasHeader ? 1 : 0; i < table.rows.items.length; i++) {
            if (i % 2 === (hasHeader ? 1 : 0)) {
              const row = table.rows.items[i];
              row.load("cells");
              await context.sync();
              
              for (const cell of row.cells.items) {
                cell.shadingColor = "#E8E8E8";
              }
            }
          }
        }

        await context.sync();

        return `Inserted ${rowCount}x${colCount} table with ${style} style${hasHeader ? " and header row" : ""}.`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
