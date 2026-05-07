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

interface OfficeDebugInfo {
  message?: string;
  errorLocation?: string;
  statement?: string;
  surroundingStatements?: string[] | string;
  fullStatements?: string[] | string;
  innerError?: unknown;
}

interface OfficeErrorLike {
  code?: string;
  debugInfo?: OfficeDebugInfo;
  innerError?: unknown;
}

function truncateStatement(statement: string, maxLength = 240): string {
  const normalized = statement.replace(/\s+/g, " ").trim();
  if (normalized.length <= maxLength) return normalized;
  return `${normalized.slice(0, maxLength - 3)}...`;
}

function formatSurroundingStatements(value: OfficeDebugInfo["surroundingStatements"], maxLines = 6): string | null {
  if (!value) return null;
  const lines = Array.isArray(value)
    ? value
    : String(value).split(/\r?\n/);
  const trimmed = lines.map((line) => line.trim()).filter(Boolean);
  if (trimmed.length === 0) return null;
  const head = trimmed.slice(0, maxLines).map((line) => `  ${truncateStatement(line, 200)}`).join("\n");
  const suffix = trimmed.length > maxLines ? `\n  ...(${trimmed.length - maxLines} more)` : "";
  return head + suffix;
}

export function describeErrorWithCode(error: unknown, depth = 0): string {
  if (!(error instanceof Error)) return String(error);
  const officeLike = error as OfficeErrorLike;
  const code = officeLike.code;
  const debugInfo = officeLike.debugInfo;

  let base = error.message;

  // Use debugInfo.message if it provides additional detail beyond the base message
  if (debugInfo?.message && debugInfo.message !== base && debugInfo.message !== code) {
    base = `${base}: ${debugInfo.message}`;
  }

  // Append error location from debugInfo (e.g. "Shapes.addGeometricShape")
  if (debugInfo?.errorLocation) {
    base = `${base} (at ${debugInfo.errorLocation})`;
  }

  // Only append [code] if it adds information not already in the message
  if (code && code !== error.message) {
    base = `${base} [${code}]`;
  }

  // Append the failing statement from extendedErrorLogging (the single most useful
  // piece of info the LLM needs to retry). Requires
  // OfficeExtension.config.extendedErrorLogging = true before the batch ran.
  if (debugInfo?.statement) {
    base = `${base}\n  statement: ${truncateStatement(debugInfo.statement)}`;
  }

  const surrounding = formatSurroundingStatements(debugInfo?.surroundingStatements);
  if (surrounding) {
    base = `${base}\n  surrounding statements:\n${surrounding}`;
  }

  // Unwrap a single level of inner error, which Office.js uses to chain
  // the underlying host exception (e.g. real reason behind a GenericException).
  const innerError = debugInfo?.innerError ?? officeLike.innerError;
  if (innerError && depth < 2) {
    const innerDescription = describeErrorWithCode(
      innerError instanceof Error ? innerError : Object.assign(new Error(String((innerError as { message?: string })?.message || innerError)), innerError),
      depth + 1,
    );
    if (innerDescription && innerDescription !== base) {
      base = `${base}\n  caused by: ${innerDescription.replace(/\n/g, "\n  ")}`;
    }
  }

  return base;
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
