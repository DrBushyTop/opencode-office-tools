import type { Tool } from "@github/copilot-sdk";

export const setSelectedRange: Tool = {
  name: "set_selected_range",
  description: "Write data to the currently selected range in Excel. The data should match the dimensions of the selection. For a single cell selection, provide a single value or a 2D array that will expand the write area. For multi-cell selections, the data dimensions should match the selection.",
  parameters: {
    type: "object",
    properties: {
      data: {
        type: "array",
        description: "2D array of values to write to the selected range. Each inner array represents a row. Example: [['Value1', 'Value2'], ['Value3', 'Value4']]",
        items: {
          type: "array",
          items: {
            type: ["string", "number", "boolean"],
          },
        },
      },
      useFormulas: {
        type: "boolean",
        description: "If true, treat string values starting with '=' as formulas. Default is true.",
      },
    },
    required: ["data"],
  },
  handler: async (args) => {
    const { data, useFormulas = true } = args as {
      data: any[][];
      useFormulas?: boolean;
    };

    try {
      return await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(["address", "rowCount", "columnCount"]);

        const worksheet = selectedRange.worksheet;
        worksheet.load("name");

        await context.sync();

        const dataRowCount = data.length;
        const dataColCount = data[0]?.length || 0;

        if (dataRowCount === 0 || dataColCount === 0) {
          return {
            textResultForLlm: "No data provided to write",
            resultType: "failure",
            error: "Empty data array",
            toolTelemetry: {}
          };
        }

        // If selection is a single cell, resize to fit the data
        // Otherwise, validate dimensions match
        let targetRange: Excel.Range;
        if (selectedRange.rowCount === 1 && selectedRange.columnCount === 1) {
          // Single cell - expand to fit data
          targetRange = selectedRange.getResizedRange(dataRowCount - 1, dataColCount - 1);
        } else {
          // Multi-cell selection - check if dimensions match
          if (dataRowCount !== selectedRange.rowCount || dataColCount !== selectedRange.columnCount) {
            return {
              textResultForLlm: `Data dimensions (${dataRowCount}x${dataColCount}) do not match selection dimensions (${selectedRange.rowCount}x${selectedRange.columnCount}). Either select a single cell to auto-expand, or provide data matching the selection size.`,
              resultType: "failure",
              error: "Dimension mismatch",
              toolTelemetry: {}
            };
          }
          targetRange = selectedRange;
        }

        // Check if we should use formulas for any cells
        if (useFormulas) {
          const hasFormulas = data.some(row => 
            row.some(cell => typeof cell === 'string' && cell.startsWith('='))
          );
          
          if (hasFormulas) {
            targetRange.formulas = data;
          } else {
            targetRange.values = data;
          }
        } else {
          targetRange.values = data;
        }

        await context.sync();

        return `Successfully wrote ${dataRowCount} rows and ${dataColCount} columns to the selected range in ${worksheet.name}`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
