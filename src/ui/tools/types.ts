export interface ToolResultFailure {
  textResultForLlm: string;
  resultType: "failure";
  error: string;
  toolTelemetry: Record<string, unknown>;
  binaryResultsForLlm?: Array<{
    data: string;
    mimeType: string;
    type: string;
    description?: string;
  }>;
}

export type ToolHandlerResult = string | ToolResultFailure | Record<string, unknown>;

export interface ToolContext {
  sessionId: string;
  toolCallId: string;
  toolName: string;
  arguments: Record<string, unknown>;
}

export interface Tool {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
  handler: (args?: unknown, context?: ToolContext) => Promise<ToolHandlerResult>;
}
