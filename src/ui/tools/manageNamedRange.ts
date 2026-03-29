import { z } from "zod";
import type { Tool } from "./types";
import { isExcelRequirementSetSupported, parseToolArgs, qualifyNamedRangeReference, toolFailure } from "./excelShared";

const workbookScopedNamePattern = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;
const a1ReferenceLikePattern = /^\$?[A-Za-z]{1,3}\$?\d+$/;
const r1c1ReferenceLikePattern = /^R(\[?-?\d+\]?|\d+)C(\[?-?\d+\]?|\d+)$/i;

export function isValidWorkbookNamedRangeName(value: string) {
  return workbookScopedNamePattern.test(value)
    && !a1ReferenceLikePattern.test(value)
    && !r1c1ReferenceLikePattern.test(value);
}

function invalidNameMessage(fieldName: "name" | "newName") {
  return `Invalid ${fieldName}. Workbook-scoped named ranges must start with a letter, underscore, or backslash; can contain letters, numbers, periods, underscores, or backslashes; and cannot look like cell references.`;
}

const manageNamedRangeArgsSchema = z.object({
  action: z.enum(["create", "update", "rename", "setVisibility", "delete"]),
  name: z.string(),
  newName: z.string().optional(),
  reference: z.string().optional(),
  comment: z.string().optional(),
  visible: z.boolean().optional(),
});

export const manageNamedRange: Tool = {
  name: "manage_named_range",
  description: "Create, update, rename, set visibility, or delete workbook-scoped Excel named ranges.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "update", "rename", "setVisibility", "delete"],
        description: "Named range operation to perform.",
      },
      name: {
        type: "string",
        description: "Existing or new workbook-scoped named range name depending on the action.",
      },
      newName: {
        type: "string",
        description: "New named range name for rename.",
      },
      reference: {
        type: "string",
        description: "Cell or formula reference for a workbook-scoped named range, such as A1:D10, Sheet1!B2, or =SUM(A:A).",
      },
      comment: {
        type: "string",
        description: "Optional description to set when creating or updating.",
      },
      visible: {
        type: "boolean",
        description: "Whether the named range is visible for setVisibility or update.",
      },
    },
    required: ["action", "name"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(manageNamedRangeArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const { action, name, newName, reference, comment, visible } = parsedArgs.data;

    if (!isValidWorkbookNamedRangeName(name)) {
      return toolFailure(invalidNameMessage("name"));
    }
    if (newName && !isValidWorkbookNamedRangeName(newName)) {
      return toolFailure(invalidNameMessage("newName"));
    }

    if ((action === "create" || action === "rename" || comment !== undefined || action === "delete") && !isExcelRequirementSetSupported("1.4")) {
      return toolFailure("This named range action requires ExcelApi 1.4.");
    }

    if ((action === "rename" || (action === "update" && reference !== undefined)) && !isExcelRequirementSetSupported("1.7")) {
      return toolFailure("This named range action requires ExcelApi 1.7.");
    }

    try {
      return await Excel.run(async (context) => {
        const names = context.workbook.names;

        if (action === "create") {
          if (!reference) return toolFailure("reference is required for create.");
          const resolvedReference = await qualifyNamedRangeReference(context, reference);
          const namedItem = names.add(name, resolvedReference, comment);
          if (visible !== undefined) namedItem.visible = visible;
          namedItem.load(["name", "value", "visible"]);
          await context.sync();
          return `Created named range ${namedItem.name} pointing to ${namedItem.value}${visible !== undefined ? ` (visible=${namedItem.visible})` : ""}.`;
        }

        const namedItem = names.getItem(name);
        const propertiesToLoad = ["name"];

        if (action === "rename") {
          propertiesToLoad.push("formula");
          if (comment === undefined) propertiesToLoad.push("comment");
          if (visible === undefined) propertiesToLoad.push("visible");
        }

        try {
          namedItem.load(propertiesToLoad);
          await context.sync();
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          if (/ItemNotFound|does not exist|cannot find/i.test(message)) {
            return toolFailure(`Named range ${name} was not found.`);
          }
          throw error;
        }

        switch (action) {
          case "update":
            if (reference !== undefined) namedItem.formula = await qualifyNamedRangeReference(context, reference);
            if (comment !== undefined) namedItem.comment = comment;
            if (visible !== undefined) namedItem.visible = visible;
            await context.sync();
            return `Updated named range ${namedItem.name}.`;
          case "rename":
            if (!newName) return toolFailure("newName is required for rename.");
            names.add(newName, namedItem.formula, comment ?? namedItem.comment);
            {
              const replacement = names.getItem(newName);
              replacement.visible = visible ?? namedItem.visible;
            }
            namedItem.delete();
            await context.sync();
            return `Renamed named range ${name} to ${newName}.`;
          case "setVisibility":
            if (visible === undefined) return toolFailure("visible is required for setVisibility.");
            namedItem.visible = visible;
            await context.sync();
            return `Set named range ${namedItem.name} visibility to ${visible}.`;
          case "delete":
            namedItem.delete();
            await context.sync();
            return `Deleted named range ${name}.`;
          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
