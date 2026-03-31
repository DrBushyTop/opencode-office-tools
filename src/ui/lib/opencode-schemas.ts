import { z } from "zod";

export const jsonObjectSchema = z.record(z.string(), z.unknown());

export const timestampValueSchema = z.union([z.string(), z.number(), z.date()]);

export const timeRangeSchema = z.object({
  start: timestampValueSchema.optional(),
}).passthrough();

const opencodeTokenSchema = z.object({
  total: z.number().optional(),
  input: z.number(),
  output: z.number(),
  reasoning: z.number(),
  cache: z.object({
    read: z.number(),
    write: z.number(),
  }),
});

export const opencodePartStateSchema = z.object({
  status: z.string().optional(),
  input: jsonObjectSchema.optional(),
  output: z.unknown().optional(),
  error: z.unknown().optional(),
  metadata: jsonObjectSchema.optional(),
  time: timeRangeSchema.optional(),
}).passthrough();

export const opencodeMessagePartSchema = z.object({
  id: z.string().optional(),
  messageID: z.string().optional(),
  type: z.string(),
  synthetic: z.boolean().optional(),
  text: z.string().optional(),
  mime: z.string().optional(),
  url: z.string().optional(),
  filename: z.string().optional(),
  tool: z.string().optional(),
  state: opencodePartStateSchema.optional(),
  time: timeRangeSchema.optional(),
}).passthrough();

export type OpencodeMessagePart = z.infer<typeof opencodeMessagePartSchema>;

export const opencodeMessageInfoSchema = z.object({
  id: z.string(),
  role: z.string(),
  time: z.object({
    created: timestampValueSchema,
    completed: timestampValueSchema.optional(),
  }).passthrough(),
  providerID: z.string().optional(),
  modelID: z.string().optional(),
  cost: z.number().optional(),
  tokens: opencodeTokenSchema.optional(),
}).passthrough();

export const opencodeMessageSchema = z.object({
  info: opencodeMessageInfoSchema.optional(),
  parts: z.array(opencodeMessagePartSchema).optional(),
}).passthrough();

export type OpencodeMessage = z.infer<typeof opencodeMessageSchema>;

export const modelInfoSchema = z.object({
  key: z.string(),
  label: z.string(),
  providerID: z.string(),
  modelID: z.string(),
  limitContext: z.number().nonnegative().optional(),
});

export const sessionInfoSchema = z.object({
  id: z.string(),
  title: z.string().nullable().optional(),
  parentID: z.string().nullable().optional(),
});

export const opencodeConfigSchema = z.object({
  agent: z.record(z.string(), z.object({
    model: z.string().optional(),
  }).catchall(z.unknown())).optional(),
}).catchall(z.unknown());

export const opencodeSessionInfoSchema = z.object({
  id: z.string(),
  title: z.string(),
  directory: z.string(),
  time: z.object({
    created: z.number(),
    updated: z.number(),
  }),
});
