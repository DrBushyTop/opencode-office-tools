import rawRegistry from "./office-tool-registry.json";
import {
  officeToolRegistrySourceSchema,
  type OfficeToolDefinition,
  type OfficeToolHost,
  type OfficeToolRegistryEntry,
} from "./office-metadata";

export type { OfficeToolHost, OfficeToolDefinition } from "./office-metadata";

type ToolArgs = Record<string, unknown>;

const parsedRegistry = officeToolRegistrySourceSchema.parse(rawRegistry);

export const officeToolRegistry = Object.entries(parsedRegistry).reduce<Record<string, OfficeToolRegistryEntry>>((acc, [name, entry]) => {
  acc[name] = { ...entry, name };
  return acc;
}, {});

export const officeToolDefinitions: Record<string, OfficeToolDefinition> = Object.fromEntries(
  Object.entries(officeToolRegistry).map(([name, entry]) => [name, {
    name: entry.name,
    description: entry.description,
    parameters: entry.parameters,
    hosts: entry.hosts,
  }]),
);

function jsonValue(value: unknown) {
  return JSON.stringify(value ?? "");
}

export function formatOfficeToolActivity(toolName: string, args: ToolArgs) {
  const entry = officeToolRegistry[toolName];
  if (!entry) return null;

  const activity = entry.ui.activity;
  switch (activity.type) {
    case "static":
      return activity.text || toolName.replace(/_/g, " ");
    case "address":
      return `${activity.prefix || ""}${String(args.address || activity.default || "document part")}`;
    case "json_value":
      return `${activity.prefix || ""}${jsonValue(args[activity.field || ""])}`;
    case "field_or_default":
      return `${activity.prefix || ""}${String(args[activity.field || ""] || activity.default || "")}`;
    case "sheet_name_or_default":
      return args[activity.field || ""]
        ? `${activity.prefix || ""}${jsonValue(args[activity.field || ""])}`
        : (activity.default || "worksheet");
    case "slide_index_plus_one": {
      const raw = Number(args[activity.field || "slideIndex"]);
      return `${activity.prefix || ""}${Number.isFinite(raw) ? raw + 1 : "?"}${activity.suffix || ""}`;
    }
    case "slide_index_plus_one_or_all": {
      if (args[activity.field || "slideIndex"] === undefined) {
        return activity.default || "Inspecting all slides";
      }
      const raw = Number(args[activity.field || "slideIndex"]);
      return `${activity.prefix || ""}${Number.isFinite(raw) ? raw + 1 : "?"}${activity.suffix || ""}`;
    }
    case "presentation_content_range": {
      if (args.slideIndex !== undefined) return `Reading slide ${Number(args.slideIndex) + 1}`;
      if (args.startIndex !== undefined && args.endIndex !== undefined) {
        return `Reading slides ${Number(args.startIndex) + 1}-${Number(args.endIndex) + 1}`;
      }
      return "Reading all slides";
    }
    default:
      return toolName.replace(/_/g, " ");
  }
}

export function getOfficeToolUi(toolName: string) {
  const entry = officeToolRegistry[toolName];
  if (!entry) return null;
  return {
    icon: entry.ui.icon,
    format: (args: ToolArgs) => formatOfficeToolActivity(toolName, args) || toolName.replace(/_/g, " "),
  };
}

export function isReadOnlyOfficeTool(toolName: string) {
  return Boolean(officeToolRegistry[toolName] && !officeToolRegistry[toolName].mutating);
}

export function getOfficeToolNames(host: OfficeToolHost) {
  return Object.values(officeToolRegistry)
    .filter((item) => item.hosts.includes(host))
    .map((item) => item.name);
}
