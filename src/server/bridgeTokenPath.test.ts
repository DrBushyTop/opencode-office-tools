import fs from "node:fs";
import path from "node:path";
import { createRequire } from "module";
import { afterEach, describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);
const {
  bridgeTokenDirectory,
  bridgeTokenPath,
  writeBridgeToken,
  readBridgeToken,
  removeBridgeToken,
} = require("./bridgeTokenPath.js");

const cleanupPaths = new Set<string>();

afterEach(() => {
  for (const filePath of cleanupPaths) {
    if (fs.existsSync(filePath)) {
      fs.rmSync(filePath, { force: true });
    }
  }
  cleanupPaths.clear();
});

describe("bridge token path", () => {
  it("writes and reads bridge tokens via a locked-down file", () => {
    const port = 62000 + Math.floor(Math.random() * 1000);
    const token = `bridge-token-${Date.now()}`;
    const filePath = writeBridgeToken(port, token);
    cleanupPaths.add(filePath);

    expect(filePath.startsWith(bridgeTokenDirectory())).toBe(true);
    expect(readBridgeToken(port)).toBe(token);

    removeBridgeToken(port);
    cleanupPaths.delete(filePath);
    expect(fs.existsSync(filePath)).toBe(false);
  });

  const symlinkCase = process.platform === "win32" ? it.skip : it;

  symlinkCase("rejects symbolic links for bridge token files", () => {
    const port = 63000 + Math.floor(Math.random() * 1000);
    const tokenPath = bridgeTokenPath(port);
    const targetPath = path.join(bridgeTokenDirectory(), `${port}.target`);
    fs.mkdirSync(bridgeTokenDirectory(), { recursive: true });
    fs.writeFileSync(targetPath, "unexpected-token", "utf8");
    fs.symlinkSync(targetPath, tokenPath);
    cleanupPaths.add(targetPath);
    cleanupPaths.add(tokenPath);

    expect(() => readBridgeToken(port)).toThrow(/symbolic link/);
  });
});
