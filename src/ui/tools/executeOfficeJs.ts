import { z } from "zod";
import type { Tool } from "./types";
import { toolFailure } from "./powerpointShared";
import { AsyncFunction, createScopedConsole, NO_RESULT, normalizeScriptValue } from "./scriptExecutionShared";

const executeOfficeJsArgsSchema = z.object({
  code: z.string().trim().min(1),
});

export async function executePowerPointOfficeJs(
  code: string,
  context: PowerPoint.RequestContext,
  externalLogs?: Array<{ level: string; values: unknown[] }>,
) {
  let explicitResult: unknown = NO_RESULT;
  const logs = externalLogs || [];
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
  description: "Office.js escape hatch for live PowerPoint automation. Use only when the higher-level tools (manage_slide_shapes, edit_slide_xml, edit_slide_text, etc.) cannot express the operation cleanly. Do not use for batch shape creation or text formatting that edit_slide_xml or manage_slide_shapes can handle.",
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

    const logs: Array<{ level: string; values: unknown[] }> = [];

    try {
      return await PowerPoint.run(async (context) => {
        const execution = await executePowerPointOfficeJs(parsedArgs.data.code, context, logs);
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
      const logsHint = logs.length > 0
        ? `\n\nConsole output before error:\n${logs.map((l) => `[${l.level}] ${l.values.map((v) => (typeof v === "string" ? v : JSON.stringify(v))).join(" ")}`).join("\n")}`
        : undefined;
      return toolFailure(error, logsHint);
    }
  },
};
