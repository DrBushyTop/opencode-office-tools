import { z } from "zod";
import type { Tool } from "./types";
import { toolFailure } from "./powerpointShared";
import { AsyncFunction, createScopedConsole, NO_RESULT, normalizeScriptValue } from "./scriptExecutionShared";

const executeOfficeJsArgsSchema = z.object({
  code: z.string().trim().min(1),
});

export async function executePowerPointOfficeJs(code: string, context: PowerPoint.RequestContext) {
  let explicitResult: unknown = NO_RESULT;
  const logs: Array<{ level: string; values: unknown[] }> = [];
  const scopedConsole = createScopedConsole(logs);

  const runner = new AsyncFunction(
    "context",
    "presentation",
    "PowerPoint",
    "Office",
    "console",
    "sync",
    "setResult",
    code,
  );

  const returned = await runner(
    context,
    context.presentation,
    PowerPoint,
    Office,
    scopedConsole,
    () => context.sync(),
    (value: unknown) => {
      explicitResult = value;
      return value;
    },
  );

  const rawResult = explicitResult === NO_RESULT ? returned : explicitResult;
  const hasResult = rawResult !== undefined;

  return {
    result: hasResult ? normalizeScriptValue(rawResult) : null,
    hasResult,
    logs,
    usedExplicitResult: explicitResult !== NO_RESULT,
  };
}

export const executeOfficeJs: Tool = {
  name: "execute_office_js",
  description: "Primary Office.js escape hatch for live PowerPoint automation. Use this for custom visualizations, geometric shape work, slide insertion or movement, fills, z-order, and host operations the higher-level tools cannot express cleanly.",
  parameters: {
    type: "object",
    properties: {
      code: {
        type: "string",
        description: "Async JavaScript function body that runs inside PowerPoint.run with `context`, `presentation`, `PowerPoint`, `Office`, `console`, `sync()`, and `setResult(value)` in scope.",
      },
    },
    required: ["code"],
  },
  handler: async (args) => {
    const parsedArgs = executeOfficeJsArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const execution = await executePowerPointOfficeJs(parsedArgs.data.code, context);
        return {
          summary: execution.hasResult
            ? "Executed custom Office.js against the live PowerPoint presentation."
            : "Executed custom Office.js against the live PowerPoint presentation. No value was returned; use return or setResult(value) to include output.",
          result: execution.result,
          logs: execution.logs,
          hasResult: execution.hasResult,
          usedExplicitResult: execution.usedExplicitResult,
        };
      });
    } catch (error) {
      return toolFailure(error);
    }
  },
};
