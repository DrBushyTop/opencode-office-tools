import { z } from "zod";
import type { Tool } from "./types";
import { toolFailure } from "./powerpointShared";
import { describeErrorWithCode } from "./toolShared";
import { AsyncFunction, createScopedConsole, NO_RESULT, normalizeScriptValue } from "./scriptExecutionShared";

const executeOfficeJsArgsSchema = z.object({
  code: z.string().trim().min(1),
});

/**
 * Enable Office.js extended error logging so that OfficeExtension.Error.debugInfo
 * carries `statement`, `surroundingStatements`, and `fullStatements` fields. Without
 * this flag, batch errors thrown at `context.sync()` typically expose only a code
 * like `InvalidArgument` with no way to tell which call failed.
 *
 * Guarded so it is a no-op in test environments where OfficeExtension is not defined.
 */
function enableExtendedErrorLogging() {
  try {
    const globalOfficeExtension = (globalThis as { OfficeExtension?: { config?: { extendedErrorLogging?: boolean } } }).OfficeExtension;
    if (globalOfficeExtension?.config) {
      globalOfficeExtension.config.extendedErrorLogging = true;
    }
  } catch {
    // Intentionally ignored — extended logging is a best-effort diagnostic aid.
  }
}

/**
 * Determine how many lines the `AsyncFunction` constructor prepends before the
 * user-supplied body in the generated source. V8 (Node), V8 (Chromium/Edge), and
 * other engines format the synthetic header differently — sometimes as a single
 * line `async function anonymous(a,b,...) {`, sometimes splitting each argument
 * onto its own line. We probe at runtime with a throwing body whose line we know
 * (body line 1) and read back the reported line in the stack trace.
 *
 * Cached after the first successful probe.
 */
let cachedHeaderOffset: number | null = null;
function getAsyncFunctionHeaderOffset(): number {
  if (cachedHeaderOffset !== null) return cachedHeaderOffset;
  try {
    // Use a synchronous Function probe with the same argument list so the generated
    // header spans the same number of lines as the async variant on this engine.
    // eslint-disable-next-line @typescript-eslint/no-implied-eval, no-new-func
    const probe = new Function(
      "context",
      "presentation",
      "PowerPoint",
      "Office",
      "console",
      "sync",
      "setResult",
      "throw new Error('__offset_probe__');",
    );
    try {
      (probe as () => void)();
    } catch (error) {
      const stack = (error as { stack?: string })?.stack;
      if (stack) {
        const frame = findAnonymousFrame(stack);
        if (frame && frame.line >= 1) {
          // Body line 1 is `throw ...`, so reportedLine - 1 is the header length.
          cachedHeaderOffset = frame.line - 1;
          return cachedHeaderOffset;
        }
      }
    }
  } catch {
    // Fall through.
  }
  cachedHeaderOffset = 1;
  return cachedHeaderOffset;
}

/**
 * Pull the first anonymous (user-code) frame from a stack trace.
 *
 * Different engines format the generated `Function` / `AsyncFunction` frame
 * differently. On Node the top frame looks like
 *   `at eval (eval at <anonymous> ([eval]:1:65), <anonymous>:3:7)`
 * and on Chromium (Office.js host) it is typically
 *   `at anonymous:3:7`
 * or `at <anonymous>:3:7`. We walk the stack frames top-down and return the
 * first `<anonymous>:L:C` / `anonymous:L:C` location we find.
 */
function findAnonymousFrame(stack: string): { line: number; column: number } | null {
  const frameLines = stack.split(/\r?\n/);
  for (const frameLine of frameLines) {
    // Prefer an explicit `<anonymous>:L:C` occurrence when present; otherwise
    // fall back to a bare `anonymous:L:C`. Match the LAST occurrence on the
    // line so nested wrappers like `eval at <anonymous> ([eval]:1:65), <anonymous>:3:7`
    // resolve to the inner user frame.
    const anchoredMatches = [...frameLine.matchAll(/<anonymous>:(\d+):(\d+)/g)];
    if (anchoredMatches.length > 0) {
      const last = anchoredMatches[anchoredMatches.length - 1];
      return { line: Number.parseInt(last[1], 10), column: Number.parseInt(last[2], 10) };
    }
    const bareMatch = frameLine.match(/\banonymous[^<:\n]*:(\d+):(\d+)/);
    if (bareMatch) {
      return { line: Number.parseInt(bareMatch[1], 10), column: Number.parseInt(bareMatch[2], 10) };
    }
  }
  return null;
}

/**
 * Extract the first user-code frame from a stack trace and map it back to a
 * snippet of the user-supplied `code`, accounting for the synthetic header
 * injected by the `AsyncFunction` constructor.
 */
export function formatUserCodeSnippet(code: string, stack: string | undefined, contextLines = 2): string | null {
  if (!stack) return null;
  const frame = findAnonymousFrame(stack);
  if (!frame) return null;
  const { line: reportedLine, column: reportedCol } = frame;
  if (!Number.isFinite(reportedLine) || reportedLine < 1) return null;

  const headerOffset = getAsyncFunctionHeaderOffset();
  const userLineIndex = reportedLine - headerOffset - 1;
  const sourceLines = code.split(/\r?\n/);
  if (userLineIndex < 0 || userLineIndex >= sourceLines.length) return null;

  const start = Math.max(0, userLineIndex - contextLines);
  const end = Math.min(sourceLines.length, userLineIndex + contextLines + 1);
  const width = String(end).length;
  const snippet: string[] = [];
  for (let i = start; i < end; i += 1) {
    const lineNumber = String(i + 1).padStart(width, " ");
    const marker = i === userLineIndex ? ">" : " ";
    snippet.push(`${marker} ${lineNumber}: ${sourceLines[i]}`);
  }
  const caretIndent = " ".repeat(width + 4) + " ".repeat(Math.max(0, reportedCol - 1));
  const insertionPoint = userLineIndex - start + 1;
  snippet.splice(insertionPoint, 0, `${caretIndent}^`);
  return snippet.join("\n");
}

/**
 * Detect Office.js batch errors (thrown from `context.sync()`). These carry the
 * useful debugInfo fields; plain JS errors do not.
 */
function isOfficeBatchError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  const candidate = error as { code?: unknown; debugInfo?: unknown };
  return typeof candidate.code === "string" || typeof candidate.debugInfo === "object";
}

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
    const userCode = parsedArgs.data.code;

    try {
      return await PowerPoint.run(async (context) => {
        enableExtendedErrorLogging();
        const execution = await executePowerPointOfficeJs(userCode, context, logs);
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
      const parts: string[] = [describeErrorWithCode(error)];

      // For non-Office errors (TypeError, ReferenceError, syntax) the stack points at
      // a line inside the user-supplied code. Surfacing a snippet with a caret is far
      // more actionable than just a message.
      if (!isOfficeBatchError(error)) {
        const snippet = formatUserCodeSnippet(userCode, (error as Error | undefined)?.stack);
        if (snippet) {
          parts.push(`\nIn user code:\n${snippet}`);
        }
      }

      if (logs.length > 0) {
        const formattedLogs = logs
          .map((l) => `[${l.level}] ${l.values.map((v) => (typeof v === "string" ? v : JSON.stringify(v))).join(" ")}`)
          .join("\n");
        parts.push(`\nConsole output before error:\n${formattedLogs}`);
      }

      return toolFailure(parts.join(""));
    }
  },
};
