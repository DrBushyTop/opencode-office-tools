import { afterEach, describe, expect, it, vi } from "vitest";
import { executeOfficeJs, executePowerPointOfficeJs, formatUserCodeSnippet } from "./executeOfficeJs";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("executeOfficeJs helpers", () => {
  it("executes async office code and returns explicit results with logs", async () => {
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", {});
    const sync = vi.fn().mockResolvedValue(undefined);
    const context = {
      sync,
      presentation: {
        slides: { items: [{ id: "slide-1" }] },
      },
    } as unknown as PowerPoint.RequestContext;

    const result = await executePowerPointOfficeJs(
      "console.log('hello', { slideCount: presentation.slides.items.length }); await sync(); setResult({ ok: true, count: presentation.slides.items.length });",
      context,
    );

    expect(sync).toHaveBeenCalledTimes(1);
    expect(result.hasResult).toBe(true);
    expect(result.usedExplicitResult).toBe(true);
    expect(result.result).toEqual({ ok: true, count: 1 });
    expect(result.logs).toEqual([
      {
        level: "log",
        values: ["hello", { slideCount: 1 }],
      },
    ]);
  });

  it("falls back to the returned value when setResult is not used", async () => {
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", {});
    const context = {
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext;

    const result = await executePowerPointOfficeJs(
      "return { slideCount: presentation.slides.items.length, marker: 'done' };",
      context,
    );

    expect(result.usedExplicitResult).toBe(false);
    expect(result.hasResult).toBe(true);
    expect(result.result).toEqual({ slideCount: 0, marker: "done" });
  });

  it("returns null when the script does not return a value", async () => {
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", {});
    const context = {
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext;

    const result = await executePowerPointOfficeJs(
      "await sync();",
      context,
    );

    expect(result.usedExplicitResult).toBe(false);
    expect(result.hasResult).toBe(false);
    expect(result.result).toBeNull();
  });

  it("normalizes non-plain office objects in logs and results", async () => {
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", {});
    class FakeOfficeObject {
      id = "shape-1";
    }
    const context = {
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext;

    const result = await executePowerPointOfficeJs(
      "const value = new (class FakeOfficeObject { constructor() { this.id = 'shape-1'; } })(); console.info(value); return value;",
      context,
    );

    expect(result.logs).toEqual([{ level: "info", values: ["[FakeOfficeObject]"] }]);
    expect(result.result).toBe("[FakeOfficeObject]");
  });
  it("populates external logs array when provided", async () => {
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", {});
    const context = {
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext;

    const externalLogs: Array<{ level: string; values: unknown[] }> = [];
    const result = await executePowerPointOfficeJs(
      "console.log('step 1'); console.warn('step 2'); return 'done';",
      context,
      externalLogs,
    );

    expect(result.logs).toBe(externalLogs);
    expect(externalLogs).toHaveLength(2);
    expect(externalLogs[0]).toEqual({ level: "log", values: ["step 1"] });
    expect(externalLogs[1]).toEqual({ level: "warn", values: ["step 2"] });
    expect(result.result).toBe("done");
  });
});

describe("executeOfficeJs tool", () => {
  it("returns structured execution output", async () => {
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [{ id: "slide-1" }] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", {
      run,
    });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({
      code: "return { slides: presentation.slides.items.length };",
    });

    expect(run).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      summary: "Executed custom Office.js against the live PowerPoint presentation.",
      result: { slides: 1 },
      logs: [],
      hasResult: true,
      usedExplicitResult: false,
    });
  });

  it("returns a no-result summary when the script has no returned value", async () => {
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({
      code: "await sync();",
    });

    expect(result).toEqual({
      summary: "Executed custom Office.js against the live PowerPoint presentation. No value was returned; use return or setResult(value) to include output.",
      result: null,
      logs: [],
      hasResult: false,
      usedExplicitResult: false,
    });
  });

  it("includes console logs captured before an error in the failure message", async () => {
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockRejectedValue(Object.assign(new Error("InvalidArgument"), { code: "InvalidArgument" })),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({
      code: "console.log('creating shapes'); console.warn('about to sync'); await sync();",
    }) as { resultType: string; error: string; textResultForLlm: string };

    expect(result.resultType).toBe("failure");
    expect(result.error).toContain("InvalidArgument");
    expect(result.error).toContain("Console output before error");
    expect(result.error).toContain("[log] creating shapes");
    expect(result.error).toContain("[warn] about to sync");
  });

  it("includes debugInfo.errorLocation in the failure message", async () => {
    const officeError = Object.assign(new Error("InvalidArgument"), {
      code: "InvalidArgument",
      debugInfo: { message: "The argument is invalid or missing", errorLocation: "Shapes.addGeometricShape" },
    });
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockRejectedValue(officeError),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({
      code: "await sync();",
    }) as { resultType: string; error: string };

    expect(result.resultType).toBe("failure");
    expect(result.error).toContain("Shapes.addGeometricShape");
    expect(result.error).toContain("The argument is invalid or missing");
  });

  it("enables OfficeExtension.config.extendedErrorLogging before running user code", async () => {
    const config = { extendedErrorLogging: false } as { extendedErrorLogging: boolean };
    vi.stubGlobal("OfficeExtension", { config });
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    await executeOfficeJs.handler({ code: "return 1;" });

    expect(config.extendedErrorLogging).toBe(true);
  });

  it("surfaces the failing statement from extended error logging", async () => {
    const officeError = Object.assign(new Error("InvalidArgument"), {
      code: "InvalidArgument",
      debugInfo: {
        errorLocation: "Shape.geometricShapeType",
        statement: 'shape.geometricShapeType = "RoundRectangle"',
        surroundingStatements: [
          "shape = shapes.addGeometricShape(...)",
          'shape.geometricShapeType = "RoundRectangle"',
        ],
      },
    });
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockRejectedValue(officeError),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({ code: "await sync();" }) as { error: string };
    expect(result.error).toContain('statement: shape.geometricShapeType = "RoundRectangle"');
    expect(result.error).toContain("surrounding statements:");
  });

  it("includes a code snippet with caret for plain JS errors thrown in user code", async () => {
    const run = vi.fn(async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) => callback({
      sync: vi.fn().mockResolvedValue(undefined),
      presentation: { slides: { items: [] } },
    } as unknown as PowerPoint.RequestContext));

    vi.stubGlobal("PowerPoint", { run });
    vi.stubGlobal("Office", {});

    const result = await executeOfficeJs.handler({
      code: [
        "const ok = 1;",
        "const missing = notDefinedSymbol;",
        "return ok;",
      ].join("\n"),
    }) as { resultType: string; error: string };

    expect(result.resultType).toBe("failure");
    expect(result.error).toContain("notDefinedSymbol is not defined");
    expect(result.error).toContain("In user code:");
    expect(result.error).toContain("> 2: const missing = notDefinedSymbol;");
  });
});

describe("formatUserCodeSnippet", () => {
  it("produces a snippet with a caret for an anonymous stack frame", () => {
    // We cannot hard-code the reported line because the AsyncFunction header
    // spans a different number of lines per engine. Instead we generate a real
    // error via the same AsyncFunction path used by the tool and feed its stack
    // straight back into formatUserCodeSnippet — this exercises the header
    // detection exactly as the tool does at runtime.
    const code = [
      "const a = 1;",
      "throw new Error('boom');",
      "const c = 3;",
    ].join("\n");
    const AsyncFn = Object.getPrototypeOf(async function () { /* noop */ }).constructor as new (
      ...args: string[]
    ) => () => Promise<unknown>;
    const fn = new AsyncFn(
      "context",
      "presentation",
      "PowerPoint",
      "Office",
      "console",
      "sync",
      "setResult",
      code,
    );

    return fn().then(
      () => { throw new Error("expected probe to throw"); },
      (error: Error) => {
        const snippet = formatUserCodeSnippet(code, error.stack);
        expect(snippet).not.toBeNull();
        expect(snippet).toContain("> 2: throw new Error('boom');");
        expect(snippet).toContain("^");
      },
    );
  });

  it("returns null when the stack has no anonymous frame", () => {
    expect(formatUserCodeSnippet("a();", "Error\n    at Foo.bar (/tmp/x.ts:1:1)")).toBeNull();
  });

  it("returns null when the mapped line is out of range", () => {
    expect(formatUserCodeSnippet("a();", "Error\n    at anonymous:999:1")).toBeNull();
  });
});
