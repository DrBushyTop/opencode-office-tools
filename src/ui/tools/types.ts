import type {
  ToolArguments,
  ToolContext,
  ToolParameters,
  ToolResultFailure,
} from "./toolShared";

export type { ToolArguments, ToolContext, ToolParameters, ToolResultFailure } from "./toolShared";

export type ToolHandlerResult = string | ToolResultFailure | Record<string, unknown>;

export interface Tool<TArgs = unknown, TResult extends ToolHandlerResult = ToolHandlerResult> {
  name: string;
  description: string;
  parameters: ToolParameters;
  handler: (args?: TArgs, context?: ToolContext) => Promise<TResult>;
}
