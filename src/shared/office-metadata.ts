import { z } from "zod";

export const officeToolHostSchema = z.enum(["word", "powerpoint", "excel", "onenote"]);
export type OfficeToolHost = z.infer<typeof officeToolHostSchema>;

export const officeToolParametersSchema = z.record(z.string(), z.unknown());

const staticToolUiActivitySchema = z.object({
  type: z.literal("static"),
  text: z.string().optional(),
});

const addressToolUiActivitySchema = z.object({
  type: z.literal("address"),
  prefix: z.string().optional(),
  default: z.string().optional(),
});

const jsonValueToolUiActivitySchema = z.object({
  type: z.literal("json_value"),
  field: z.string().optional(),
  prefix: z.string().optional(),
});

const fieldOrDefaultToolUiActivitySchema = z.object({
  type: z.literal("field_or_default"),
  field: z.string().optional(),
  prefix: z.string().optional(),
  default: z.string().optional(),
});

const sheetNameOrDefaultToolUiActivitySchema = z.object({
  type: z.literal("sheet_name_or_default"),
  field: z.string().optional(),
  prefix: z.string().optional(),
  default: z.string().optional(),
});

const slideIndexPlusOneToolUiActivitySchema = z.object({
  type: z.literal("slide_index_plus_one"),
  field: z.string().optional(),
  prefix: z.string().optional(),
  suffix: z.string().optional(),
});

const slideIndexPlusOneOrAllToolUiActivitySchema = z.object({
  type: z.literal("slide_index_plus_one_or_all"),
  field: z.string().optional(),
  prefix: z.string().optional(),
  suffix: z.string().optional(),
  default: z.string().optional(),
});

const presentationContentRangeToolUiActivitySchema = z.object({
  type: z.literal("presentation_content_range"),
});

export const toolUiActivitySchema = z.discriminatedUnion("type", [
  staticToolUiActivitySchema,
  addressToolUiActivitySchema,
  jsonValueToolUiActivitySchema,
  fieldOrDefaultToolUiActivitySchema,
  sheetNameOrDefaultToolUiActivitySchema,
  slideIndexPlusOneToolUiActivitySchema,
  slideIndexPlusOneOrAllToolUiActivitySchema,
  presentationContentRangeToolUiActivitySchema,
]);
export type ToolUiActivity = z.infer<typeof toolUiActivitySchema>;

export const officeToolDefinitionSchema = z.object({
  name: z.string(),
  description: z.string(),
  parameters: officeToolParametersSchema,
  hosts: z.array(officeToolHostSchema),
});
export type OfficeToolDefinition = z.infer<typeof officeToolDefinitionSchema>;

export const officeToolRegistrySourceEntrySchema = z.object({
  hosts: z.array(officeToolHostSchema),
  description: z.string(),
  mutating: z.boolean(),
  parameters: officeToolParametersSchema,
  ui: z.object({
    icon: z.string(),
    activity: toolUiActivitySchema,
  }),
});

export const officeToolRegistrySourceSchema = z.record(z.string(), officeToolRegistrySourceEntrySchema);
export type OfficeToolRegistrySourceEntry = z.infer<typeof officeToolRegistrySourceEntrySchema>;

export const officeToolRegistryEntrySchema = officeToolDefinitionSchema.extend({
  mutating: z.boolean(),
  ui: z.object({
    icon: z.string(),
    activity: toolUiActivitySchema,
  }),
});
export type OfficeToolRegistryEntry = z.infer<typeof officeToolRegistryEntrySchema>;

export const officePermissionMetadataSchema = z.record(z.string(), z.unknown());

export const officePermissionRequestSchema = z.object({
  id: z.string(),
  sessionID: z.string(),
  permission: z.string(),
  patterns: z.array(z.string()),
  metadata: officePermissionMetadataSchema,
  always: z.array(z.string()),
  tool: z.object({
    messageID: z.string(),
    callID: z.string(),
  }).optional(),
});
export type OfficePermissionRequest = z.infer<typeof officePermissionRequestSchema>;
