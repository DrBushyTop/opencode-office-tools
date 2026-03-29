import { z } from "zod";
import type { Tool } from "./types";
import { getWorksheet, isExcelRequirementSetSupported, nonNegativeIntegerSchema, parseToolArgs, toolFailure } from "./excelShared";

const manageWorksheetArgsSchema = z.object({
  action: z.enum(["create", "rename", "delete", "move", "setVisibility", "activate", "freeze", "unfreeze", "protect", "unprotect"]),
  sheetName: z.string().optional(),
  newName: z.string().optional(),
  targetPosition: nonNegativeIntegerSchema("targetPosition must be a non-negative integer.").optional(),
  visibility: z.enum(["Visible", "Hidden", "VeryHidden"]).optional(),
  freezeRows: nonNegativeIntegerSchema("freezeRows must be a non-negative integer.").optional(),
  freezeColumns: nonNegativeIntegerSchema("freezeColumns must be a non-negative integer.").optional(),
  freezeRange: z.string().optional(),
  password: z.string().optional(),
});

export const manageWorksheet: Tool = {
  name: "manage_worksheet",
  description: "Create, rename, delete, move, change visibility, freeze, unfreeze, activate, protect, or unprotect Excel worksheets.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "rename", "delete", "move", "setVisibility", "activate", "freeze", "unfreeze", "protect", "unprotect"],
        description: "Worksheet operation to perform.",
      },
      sheetName: {
        type: "string",
        description: "Target worksheet name. Omit for active sheet when supported by the action.",
      },
      newName: {
        type: "string",
        description: "New worksheet name for create or rename.",
      },
      targetPosition: {
        type: "number",
        description: "Zero-based worksheet position for move or create.",
      },
      visibility: {
        type: "string",
        enum: ["Visible", "Hidden", "VeryHidden"],
        description: "Visibility to apply for setVisibility.",
      },
      freezeRows: {
        type: "number",
        description: "Number of top rows to freeze for freeze action.",
      },
      freezeColumns: {
        type: "number",
        description: "Number of left columns to freeze for freeze action.",
      },
      freezeRange: {
        type: "string",
        description: "Range to freeze at for freeze action, such as 'B2'.",
      },
      password: {
        type: "string",
        description: "Optional worksheet protection password for protect or unprotect.",
      },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(manageWorksheetArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const { action, sheetName, newName, targetPosition, visibility, freezeRows, freezeColumns, freezeRange, password } = parsedArgs.data;

    try {
      return await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name,items/position");
        await context.sync();

        if (action === "create") {
          const sheet = worksheets.add(newName);
          if (targetPosition !== undefined) {
            sheet.position = targetPosition;
          }
          if (visibility) {
            sheet.visibility = visibility;
          }
          sheet.load(["name", "position", "visibility"]);
          await context.sync();
          return `Created worksheet ${sheet.name} at position ${sheet.position} with visibility ${sheet.visibility}.`;
        }

        const sheet = await getWorksheet(context, sheetName);

        switch (action) {
          case "rename":
            if (!newName) return toolFailure("newName is required for rename.");
            sheet.name = newName;
            await context.sync();
            return `Renamed worksheet to ${newName}.`;
          case "delete":
            sheet.delete();
            await context.sync();
            return `Deleted worksheet ${sheet.name}.`;
          case "move":
            if (targetPosition === undefined) return toolFailure("targetPosition is required for move.");
            sheet.position = targetPosition;
            await context.sync();
            return `Moved worksheet ${sheet.name} to position ${targetPosition}.`;
          case "setVisibility":
            if (!visibility) return toolFailure("visibility is required for setVisibility.");
            sheet.visibility = visibility;
            await context.sync();
            return `Set worksheet ${sheet.name} visibility to ${visibility}.`;
          case "activate":
            sheet.activate();
            await context.sync();
            return `Activated worksheet ${sheet.name}.`;
          case "freeze": {
            if (!isExcelRequirementSetSupported("1.7")) {
              return toolFailure("Freezing panes requires ExcelApi 1.7.");
            }
            if (freezeRange) {
              sheet.freezePanes.freezeAt(freezeRange);
            } else if (freezeRows !== undefined && freezeColumns !== undefined) {
              const freezeCell = sheet.getCell(freezeRows, freezeColumns);
              sheet.freezePanes.freezeAt(freezeCell);
            } else if (freezeRows !== undefined) {
              sheet.freezePanes.freezeRows(freezeRows);
            } else if (freezeColumns !== undefined) {
              sheet.freezePanes.freezeColumns(freezeColumns);
            } else {
              return toolFailure("Provide freezeRange, freezeRows, or freezeColumns for freeze.");
            }
            await context.sync();
            return `Updated frozen panes on ${sheet.name}.`;
          }
          case "unfreeze":
            if (!isExcelRequirementSetSupported("1.7")) {
              return toolFailure("Unfreezing panes requires ExcelApi 1.7.");
            }
            sheet.freezePanes.unfreeze();
            await context.sync();
            return `Unfroze panes on ${sheet.name}.`;
          case "protect":
            if (password) {
              if (!isExcelRequirementSetSupported("1.7")) {
                return toolFailure("Worksheet protection passwords require ExcelApi 1.7.");
              }
              sheet.protection.protect(undefined, password);
            } else {
              sheet.protection.protect();
            }
            await context.sync();
            return `Protected worksheet ${sheet.name}.`;
          case "unprotect":
            if (password) {
              if (!isExcelRequirementSetSupported("1.7")) {
                return toolFailure("Worksheet protection passwords require ExcelApi 1.7.");
              }
              sheet.protection.unprotect(password);
            } else {
              sheet.protection.unprotect();
            }
            await context.sync();
            return `Unprotected worksheet ${sheet.name}.`;
          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
