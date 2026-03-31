export const AsyncFunction = Object.getPrototypeOf(async function () { /* noop */ }).constructor as new (
  ...args: string[]
) => (...fnArgs: unknown[]) => Promise<unknown>;

export interface ScriptLogEntry {
  level: string;
  values: unknown[];
}

export const NO_RESULT = Symbol("no-result");

export function normalizeScriptValue(value: unknown, depth = 0): unknown {
  if (depth > 6) return "[MaxDepth]";
  if (value === undefined) return "[undefined]";
  if (value === null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  if (typeof value === "bigint") return value.toString();
  if (value instanceof Error) {
    return {
      name: value.name,
      message: value.message,
      stack: value.stack,
    };
  }
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeScriptValue(entry, depth + 1));
  }
  if (typeof value === "function") {
    return `[Function ${value.name || "anonymous"}]`;
  }
  if (typeof value === "object") {
    const prototype = Object.getPrototypeOf(value);
    if (prototype === Object.prototype || prototype === null) {
      return Object.fromEntries(
        Object.entries(value as Record<string, unknown>).map(([key, entry]) => [key, normalizeScriptValue(entry, depth + 1)]),
      );
    }

    const constructorName = (value as { constructor?: { name?: string } }).constructor?.name;
    return `[${constructorName || "Object"}]`;
  }

  return String(value);
}

export function createScopedConsole(logs: ScriptLogEntry[]) {
  const push = (level: string, args: unknown[]) => {
    logs.push({
      level,
      values: args.map((arg) => normalizeScriptValue(arg)),
    });
  };

  return {
    log: (...args: unknown[]) => { push("log", args); },
    info: (...args: unknown[]) => { push("info", args); },
    warn: (...args: unknown[]) => { push("warn", args); },
    error: (...args: unknown[]) => { push("error", args); },
  };
}
