import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { describe, expect, it, vi } from "vitest";

function runCommonJsModule(filePath: string, mocks: Record<string, unknown>, globals: Record<string, unknown>) {
  const source = fs.readFileSync(filePath, "utf8").replace(/^#!.*\n/, "");
  const module = { exports: {} };
  const customRequire: any = (specifier: string) => {
    if (specifier in mocks) return mocks[specifier];
    throw new Error(`Unexpected require: ${specifier}`);
  };
  customRequire.resolve = (specifier: string) => specifier;
  customRequire.cache = Object.create(null);
  const context = vm.createContext({
    ...globals,
    module,
    exports: module.exports,
    require: customRequire,
    __filename: filePath,
    __dirname: path.dirname(filePath),
  });

  new vm.Script(source, { filename: filePath }).runInContext(context);
  return { module, context };
}

describe("tray main entrypoint", () => {
  it("creates the tray, starts the server, and advertises the running service state", async () => {
    const filePath = path.resolve("src/tray/main.js");
    let readyHandler: (() => Promise<void>) | undefined;
    const appEvents: Record<string, (...args: unknown[]) => void> = {};
    const processEvents: Record<string, (...args: unknown[]) => void> = {};
    const trayInstances: Array<{ setContextMenu: ReturnType<typeof vi.fn>; setToolTip: ReturnType<typeof vi.fn>; on: ReturnType<typeof vi.fn>; popUpContextMenu: ReturnType<typeof vi.fn> }> = [];
    const buildFromTemplate = vi.fn((template) => ({ template }));
    const createServer = vi.fn(async () => ({ close: vi.fn() }));
    const icon = {
      resize: vi.fn(() => icon),
      setTemplateImage: vi.fn(),
    };

    const app = {
      requestSingleInstanceLock: vi.fn(() => true),
      quit: vi.fn(),
      dock: { hide: vi.fn() },
      isPackaged: false,
      whenReady: vi.fn(() => ({ then: (callback: () => Promise<void>) => {
        readyHandler = callback;
      } })),
      on: vi.fn((event: string, handler: (...args: unknown[]) => void) => {
        appEvents[event] = handler;
      }),
      getPath: vi.fn(() => "/Users/test/Library/Application Support/OpenCode"),
    };

    class FakeTray {
      setContextMenu = vi.fn();
      setToolTip = vi.fn();
      on = vi.fn();
      popUpContextMenu = vi.fn();

      constructor() {
        trayInstances.push(this);
      }
    }

    const fakeProcess: any = {
      env: {},
      platform: "darwin",
      resourcesPath: "/Applications/OpenCode.app/Contents/Resources",
      exit: vi.fn(),
      on: vi.fn((event: string, handler: (...args: unknown[]) => void) => {
        processEvents[event] = handler;
      }),
    };

    runCommonJsModule(filePath, {
      electron: {
        app,
        Tray: FakeTray,
        Menu: { buildFromTemplate },
        nativeImage: {
          createFromPath: vi.fn(() => icon),
          createEmpty: vi.fn(() => ({ empty: true })),
        },
        shell: { showItemInFolder: vi.fn() },
      },
      path,
      zod: await import("zod"),
      "../server/devLogger": {
        logInfo: vi.fn(),
        logError: vi.fn(),
        getLogFilePath: vi.fn(() => "/tmp/fallback.log"),
      },
      "../server-prod.js": { createServer },
    }, {
      console,
      process: fakeProcess,
    });

    expect(app.requestSingleInstanceLock).toHaveBeenCalledTimes(1);
    expect(app.dock.hide).toHaveBeenCalledTimes(1);
    expect(readyHandler).toEqual(expect.any(Function));

    await readyHandler?.();

    expect(createServer).toHaveBeenCalledTimes(1);
    expect(fakeProcess.env).toMatchObject({
      OPENCODE_OFFICE_BASE_PATH: path.resolve("."),
      OPENCODE_OFFICE_DIRECTORY: path.resolve("."),
      OPENCODE_OFFICE_CONFIG_DIR: path.join(path.resolve("."), ".opencode"),
    });
    expect(fakeProcess.env.OPENCODE_OFFICE_LOG_FILE).toBe(path.join(path.resolve("."), ".opencode", "debug.log"));
    expect(processEvents.uncaughtException).toEqual(expect.any(Function));
    expect(processEvents.unhandledRejection).toEqual(expect.any(Function));
    expect(trayInstances).toHaveLength(1);

    const tray = trayInstances[0];
    const lastMenu = tray.setContextMenu.mock.calls[tray.setContextMenu.mock.calls.length - 1]?.[0] as { template: Array<{ label?: string }> };
    expect(lastMenu.template.map((entry) => entry.label).filter(Boolean)).toEqual(expect.arrayContaining([
      "OpenCode Office Add-in",
      "● Service Running",
      "Disable Service",
      "Open Debug Log",
      "Quit",
    ]));
    expect(tray.setToolTip.mock.calls[tray.setToolTip.mock.calls.length - 1]?.[0]).toBe("OpenCode Office Add-in - Running");
    expect(appEvents["before-quit"]).toEqual(expect.any(Function));
  });
});
