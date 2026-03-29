import { z } from "zod";

const dateSchema = z.preprocess((value) => {
  if (value instanceof Date) return value;
  if (typeof value === "string" || typeof value === "number") {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return date;
  }
  return value;
}, z.date());

export const officeHostSchema = z.enum(["powerpoint", "word", "excel", "onenote"]);

export type OfficeHost = z.infer<typeof officeHostSchema>;

export const savedMessageSchema = z.object({
  id: z.string(),
  text: z.string(),
  sender: z.enum(["user", "assistant", "tool", "thinking"]),
  timestamp: dateSchema,
  toolName: z.string().optional(),
  toolArgs: z.record(z.string(), z.unknown()).optional(),
  toolResult: z.unknown().optional(),
  toolError: z.string().optional(),
  toolMeta: z.record(z.string(), z.unknown()).optional(),
  toolStatus: z.enum(["running", "completed", "error"]).optional(),
  images: z.array(z.object({
    dataUrl: z.string(),
    name: z.string(),
  })).optional(),
});

export const savedSessionSchema = z.object({
  id: z.string(),
  title: z.string(),
  model: z.string(),
  messages: z.array(savedMessageSchema),
  createdAt: z.string(),
  updatedAt: z.string(),
});

export type SavedSession = z.infer<typeof savedSessionSchema>;

export function getHostFromOfficeHost(host: typeof Office.HostType[keyof typeof Office.HostType]): OfficeHost {
  switch (host) {
    case Office.HostType.PowerPoint:
      return "powerpoint";
    case Office.HostType.Word:
      return "word";
    case Office.HostType.Excel:
      return "excel";
    case Office.HostType.OneNote:
      return "onenote";
    default:
      return "word";
  }
}
