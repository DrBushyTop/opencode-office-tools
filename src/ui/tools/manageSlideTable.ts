import type { Tool } from "./types";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { getShapeBounds, getSlideByIndex, toPowerPointTableValues } from "./powerpointNativeContent";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { toolFailure } from "./powerpointShared";

type ManageSlideTableAction = "create" | "update" | "delete";

interface ManageSlideTableArgs {
  action: ManageSlideTableAction;
  slideIndex?: number;
  shapeId?: string | number;
  values?: Array<Array<boolean | number | string>>;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  name?: string;
}

export const manageSlideTable: Tool = {
  name: "manage_slide_table",
  description: "Create, update, or delete editable native PowerPoint tables on a slide.",
  parameters: {
    type: "object",
    properties: {
      action: { type: "string", enum: ["create", "update", "delete"] },
      slideIndex: { type: "number", description: "0-based slide index. Defaults to the active slide when available." },
      shapeId: { anyOf: [{ type: "string" }, { type: "number" }], description: "Existing table shape id for update or delete." },
      values: {
        type: "array",
        items: { type: "array", items: { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }] } },
        description: "2D table values for create or update.",
      },
      left: { type: "number" },
      top: { type: "number" },
      width: { type: "number" },
      height: { type: "number" },
      name: { type: "string" },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const tableArgs = resolvePowerPointTargetingArgs(args as ManageSlideTableArgs);
    if (!Number.isInteger(tableArgs.slideIndex) || (tableArgs.slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if ((tableArgs.action === "create" || tableArgs.action === "update") && (!tableArgs.values || tableArgs.values.length === 0 || tableArgs.values[0].length === 0)) {
      return toolFailure("values must be a non-empty 2D array for create and update.");
    }
    if ((tableArgs.action === "update" || tableArgs.action === "delete") && tableArgs.shapeId === undefined) {
      return toolFailure("shapeId is required for update and delete.");
    }

    const slideIndex = tableArgs.slideIndex as number;

    try {
      return await PowerPoint.run(async (context) => {
        const slide = await getSlideByIndex(context, slideIndex);
        const values = toPowerPointTableValues(tableArgs.values);

        if (tableArgs.action === "create") {
          const created = slide.shapes.addTable(values.length, values[0].length, {
            values,
            left: tableArgs.left,
            top: tableArgs.top,
            width: tableArgs.width,
            height: tableArgs.height,
          });
          if (tableArgs.name) created.name = tableArgs.name;
          created.load(["id", "name"]);
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Created table ${created.id} on slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: created.id,
            toolTelemetry: { slideIndex, shapeId: created.id },
          };
        }

        const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, tableArgs.shapeId!);
        if (tableArgs.action === "delete") {
          resolved.shape.delete();
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Deleted table ${resolved.shapeId} from slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: resolved.shapeId,
            toolTelemetry: { slideIndex, shapeId: resolved.shapeId },
          };
        }

        resolved.shape.load("type");
        await context.sync();
        const bounds = await getShapeBounds(resolved.shape, context);
        resolved.shape.delete();
        const created = slide.shapes.addTable(values.length, values[0].length, {
          values,
          left: tableArgs.left ?? bounds.left,
          top: tableArgs.top ?? bounds.top,
          width: tableArgs.width ?? bounds.width,
          height: tableArgs.height ?? bounds.height,
        });
        if (tableArgs.name || bounds.name) created.name = tableArgs.name || bounds.name;
        created.load(["id", "name"]);
        await context.sync();
        return {
          resultType: "success",
          textResultForLlm: `Updated table ${resolved.shapeId} on slide ${slideIndex + 1}.`,
          slideIndex,
          shapeId: created.id,
          replacedShapeId: resolved.shapeId,
          toolTelemetry: { slideIndex, replacedShapeId: resolved.shapeId, shapeId: created.id },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
