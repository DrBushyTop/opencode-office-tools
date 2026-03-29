import { z } from "zod";

export const toolArgumentsSchema = z.record(z.string(), z.unknown());

export const binaryToolResultSchema = z.object({
  data: z.string(),
  mimeType: z.string(),
  type: z.string(),
  description: z.string().optional(),
});

export const toolTelemetrySchema = z.record(z.string(), z.unknown());

export const toolResultFailureSchema = z.object({
  textResultForLlm: z.string(),
  resultType: z.literal("failure"),
  error: z.string(),
  toolTelemetry: toolTelemetrySchema,
  binaryResultsForLlm: z.array(binaryToolResultSchema).optional(),
});

export const toolContextSchema = z.object({
  sessionId: z.string(),
  toolCallId: z.string(),
  toolName: z.string(),
  arguments: toolArgumentsSchema,
});

export const toolParametersSchema = z.object({
  type: z.string(),
}).catchall(z.unknown());

export type ToolArguments = z.infer<typeof toolArgumentsSchema>;
export type ToolResultFailure = z.infer<typeof toolResultFailureSchema>;
export type ToolContext = z.infer<typeof toolContextSchema>;
export type ToolParameters = z.infer<typeof toolParametersSchema>;

export function describeError(error: unknown): string {
  if (!(error instanceof Error)) return String(error);
  return error.message;
}

export function describeErrorWithCode(error: unknown): string {
  if (!(error instanceof Error)) return String(error);
  const code = (error as { code?: string }).code;
  return code ? `${error.message} [${code}]` : error.message;
}

export function createToolFailure(
  error: unknown,
  options: { hint?: string; describe?: (error: unknown) => string } = {},
): ToolResultFailure {
  const message = (options.describe || describeError)(error);
  const fullMessage = options.hint ? `${message} ${options.hint}` : message;
  return {
    textResultForLlm: fullMessage,
    resultType: "failure",
    error: fullMessage,
    toolTelemetry: {},
  };
}

export function isToolResultFailure(result: unknown): result is ToolResultFailure {
  return toolResultFailureSchema.safeParse(result).success;
}

export function summarizePlainText(text: string, limit: number): string {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return "(empty)";
  return normalized.length > limit ? `${normalized.slice(0, limit - 3)}...` : normalized;
}
