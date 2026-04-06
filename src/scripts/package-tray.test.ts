import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { describe, expect, it, vi } from "vitest";

function runCommonJsModule(filePath: string, mocks: Record<string, unknown>) {
  const source = fs.readFileSync(filePath, "utf8").replace(/^#!.*\n/, "");
  const module = { exports: {} };
  const customRequire = ((specifier: string) => {
    if (specifier in mocks) return mocks[specifier];
    throw new Error(`Unexpected require: ${specifier}`);
  }) as NodeRequire;
  const context = vm.createContext({
    module,
    exports: module.exports,
    require: customRequire,
    __filename: filePath,
    __dirname: path.dirname(filePath),
    console,
    process: { exit: vi.fn() },
  });

  new vm.Script(source, { filename: filePath }).runInContext(context);
}

describe("package-tray script", () => {
  it("packages the macOS tray build with the expected copy and zip sequence", () => {
    const filePath = path.resolve("scripts/package-tray.js");
    const projectRoot = path.resolve(".");
    const buildDir = path.join(projectRoot, "build", "tray-package");
    const macArmDir = path.join(projectRoot, "build", "electron", "mac-arm64");
    const sourceApp = path.join(macArmDir, "OpenCode Office Add-in.app");
    const zipPath = path.join(projectRoot, "build", "opencode-office-addin-macos-v1.2.3.zip");
    const existingPaths = new Set([buildDir, macArmDir, sourceApp, zipPath]);
    const execSync = vi.fn();
    const fsMock = {
      existsSync: vi.fn((target: string) => existingPaths.has(target)),
      rmSync: vi.fn(),
      mkdirSync: vi.fn(),
      copyFileSync: vi.fn(),
      chmodSync: vi.fn(),
      unlinkSync: vi.fn(),
    };

    runCommonJsModule(filePath, {
      child_process: { execSync },
      fs: fsMock,
      path,
      os: { platform: vi.fn(() => "darwin") },
      [path.join(projectRoot, "package.json")]: { version: "1.2.3" },
    });

    expect(fsMock.rmSync).toHaveBeenCalledWith(buildDir, { recursive: true });
    expect(fsMock.mkdirSync).toHaveBeenCalledWith(buildDir, { recursive: true });
    expect(execSync.mock.calls.map(([command]) => command)).toEqual([
      "bun run clean:extraneous && bun run build",
      "bunx electron-builder --mac --dir",
      `cp -R \"${sourceApp}\" \"${buildDir}/\"`,
      `ditto -c -k --sequesterRsrc \"${buildDir}\" \"${zipPath}\"`,
    ]);
    expect(fsMock.copyFileSync).toHaveBeenCalledWith(path.join(projectRoot, "register.sh"), path.join(buildDir, "register.sh"));
    expect(fsMock.copyFileSync).toHaveBeenCalledWith(path.join(projectRoot, "installer", "GETTING_STARTED_RELEASE.md"), path.join(buildDir, "GETTING_STARTED.md"));
    expect(fsMock.chmodSync).toHaveBeenCalledWith(path.join(buildDir, "register.sh"), 0o755);
    expect(fsMock.unlinkSync).toHaveBeenCalledWith(zipPath);
  });

  it("packages the Windows tray build with register and archive steps", () => {
    const filePath = path.resolve("scripts/package-tray.js");
    const projectRoot = path.resolve(".");
    const buildDir = path.join(projectRoot, "build", "tray-package");
    const winDir = path.join(projectRoot, "build", "electron", "win-unpacked");
    const zipPath = path.join(projectRoot, "build", "opencode-office-addin-windows-v2.0.0.zip");
    const existingPaths = new Set([winDir]);
    const execSync = vi.fn();
    const fsMock = {
      existsSync: vi.fn((target: string) => existingPaths.has(target)),
      rmSync: vi.fn(),
      mkdirSync: vi.fn(),
      copyFileSync: vi.fn(),
      chmodSync: vi.fn(),
      unlinkSync: vi.fn(),
    };

    runCommonJsModule(filePath, {
      child_process: { execSync },
      fs: fsMock,
      path,
      os: { platform: vi.fn(() => "win32") },
      [path.join(projectRoot, "package.json")]: { version: "2.0.0" },
    });

    expect(execSync.mock.calls.map(([command]) => command)).toEqual([
      "bun run clean:extraneous && bun run build",
      "bunx electron-builder --win --dir",
      `xcopy \"${winDir}\\*\" \"${buildDir}\\\" /E /I /Y`,
      `powershell -Command \"Compress-Archive -Path '${buildDir}\\*' -DestinationPath '${zipPath}' -Force\"`,
    ]);
    expect(fsMock.copyFileSync).toHaveBeenCalledWith(path.join(projectRoot, "register.ps1"), path.join(buildDir, "register.ps1"));
    expect(fsMock.copyFileSync).toHaveBeenCalledWith(path.join(projectRoot, "installer", "GETTING_STARTED_RELEASE.md"), path.join(buildDir, "GETTING_STARTED.md"));
  });
});
