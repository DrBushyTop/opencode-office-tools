import type { Tool } from "@github/copilot-sdk";

export const createNamedRange: Tool = {
  name: "create_named_range",
  description: `Create or update a named range in the Excel workbook.

Named ranges make it easier to reference cells in formulas and for the AI to understand your data.

Parameters:
- name: The name for the range (must start with letter, no spaces, e.g., "SalesData", "Q1_Revenue")
- range: The cell range to name (e.g., "A1:D100", "Sheet1!B2:E50")
- comment: Optional description of what this range contains

Examples:
- Name a data table: name="SalesData", range="A1:E100"
- Name a specific cell: name="TaxRate", range="B1"
- Name with sheet reference: name="Q1Revenue", range="Sales!C2:C50"`,
  parameters: {
    type: "object",
    properties: {
      name: {
        type: "string",
        description: "The name for the range (no spaces, must start with letter).",
      },
      range: {
        type: "string",
        description: "The cell range to name (e.g., 'A1:D100').",
      },
      comment: {
        type: "string",
        description: "Optional description of what this range contains.",
      },
    },
    required: ["name", "range"],
  },
  handler: async (args) => {
    const { name, range: rangeAddress, comment } = args as {
      name: string;
      range: string;
      comment?: string;
    };

    // Validate name format
    if (!name || !/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) {
      return { 
        textResultForLlm: "Invalid name. Must start with a letter and contain only letters, numbers, and underscores (no spaces).", 
        resultType: "failure", 
        error: "Invalid name format", 
        toolTelemetry: {} 
      };
    }

    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const names = workbook.names;
        names.load("items");
        await context.sync();

        // Check if name already exists
        let existingName: Excel.NamedItem | null = null;
        for (const n of names.items) {
          n.load("name");
        }
        await context.sync();

        for (const n of names.items) {
          if (n.name.toLowerCase() === name.toLowerCase()) {
            existingName = n;
            break;
          }
        }

        // Determine the full range reference
        let fullReference = rangeAddress;
        if (!rangeAddress.includes("!")) {
          // Add sheet reference if not present
          const activeSheet = workbook.worksheets.getActiveWorksheet();
          activeSheet.load("name");
          await context.sync();
          fullReference = `'${activeSheet.name}'!${rangeAddress}`;
        }

        if (existingName) {
          // Delete existing and recreate (can't directly update reference)
          existingName.delete();
          await context.sync();
        }

        // Add the named range
        const newName = names.add(name, fullReference, comment);
        newName.load(["name", "value"]);
        await context.sync();

        const action = existingName ? "Updated" : "Created";
        return `${action} named range "${newName.name}" pointing to ${newName.value}${comment ? ` (${comment})` : ""}.`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
