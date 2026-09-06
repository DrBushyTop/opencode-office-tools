import express from "express";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createServer, type Server } from "node:http";
import { createRequire } from "module";
import { afterEach, describe, expect, it, vi } from "vitest";

const require = createRequire(import.meta.url);
const { createApiRouter, createBridgeRouter } = require("./api.js");
const { OfficeToolBridge } = require("./officeToolBridge.js");
const { OpencodeRuntime } = require("./opencodeRuntime.js");

const closers: Array<() => Promise<void>> = [];

afterEach(async () => {
  while (closers.length) {
    await closers.pop()?.();
  }
});

function httpsOrigin(baseUrl: string) {
  return baseUrl.replace("http://", "https://");
}

async function startServer(router: any, options?: { trustProxy?: boolean; secure?: boolean }) {
  const app = express();
  app.set("trust proxy", options?.trustProxy === true);
  if (options?.secure) {
    app.use((req, _res, next) => {
      Object.defineProperty(req.socket, "encrypted", { configurable: true, value: true });
      next();
    });
  }
  app.use("/api", router);

  const server = await new Promise<Server>((resolve) => {
    const nextServer = createServer(app);
    nextServer.listen(0, "127.0.0.1", () => resolve(nextServer));
  });
  closers.push(() => new Promise<void>((resolve, reject) => {
    server.close((error) => {
      if (error) reject(error);
      else resolve();
    });
  }));

  const address = server.address();
  if (!address || typeof address === "string") {
    throw new Error("Expected an IPv4 test server address");
  }
  return { baseUrl: `http://127.0.0.1:${address.port}` };
}

async function startApiServer(
  options?: { trustProxy?: boolean; secure?: boolean },
  runtimeOverrides?: Record<string, unknown>,
) {
  const runtime = {
    directory: () => process.cwd(),
    status: async () => ({ models: [], directory: process.cwd() }),
    request: async () => ({ ok: true }),
    ...runtimeOverrides,
  };
  const bridge = new OfficeToolBridge();
  return {
    runtime,
    bridge,
    ...(await startServer(createApiRouter(runtime, bridge), options)),
  };
}

async function startBridgeServer() {
  const bridge = new OfficeToolBridge();
  return {
    bridge,
    ...(await startServer(createBridgeRouter(bridge))),
  };
}

describe("server api hardening", () => {
  it("rejects spoofed forwarded https when proxy trust is disabled", async () => {
    const { baseUrl } = await startApiServer();
    const response = await fetch(`${baseUrl}/api/office-tools/session`, {
      headers: { "x-forwarded-proto": "https" },
    });

    expect(response.status).toBe(403);
    await expect(response.json()).resolves.toMatchObject({ error: expect.stringContaining("HTTPS") });
  });

  it("accepts forwarded https only when proxy trust is enabled", async () => {
    const { baseUrl } = await startApiServer({ trustProxy: true });
    const response = await fetch(`${baseUrl}/api/office-tools/session`, {
      headers: { "x-forwarded-proto": "https" },
    });

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toMatchObject({ sessionToken: expect.any(String) });
  });

  it("accepts direct secure requests for bridge sessions", async () => {
    const { baseUrl } = await startApiServer({ secure: true });
    const response = await fetch(`${baseUrl}/api/office-tools/session`);

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toMatchObject({ sessionToken: expect.any(String) });
  });

  it("returns 401 for invalid bridge tokens on execute", async () => {
    const { baseUrl } = await startApiServer();
    const response = await fetch(`${baseUrl}/api/office-tools/execute`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-office-bridge-token": "invalid-token",
      },
      body: JSON.stringify({ host: "word", toolName: "get_document_content", args: {} }),
    });

    expect(response.status).toBe(401);
    await expect(response.json()).resolves.toMatchObject({ error: expect.stringContaining("Invalid Office bridge token") });
  });

  it("sanitizes uploaded filenames before writing them", async () => {
    const { baseUrl } = await startApiServer({ secure: true });
    const response = await fetch(`${baseUrl}/api/upload-image`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        origin: httpsOrigin(baseUrl),
        referer: `${httpsOrigin(baseUrl)}/index.html`,
        "sec-fetch-site": "same-origin",
      },
      body: JSON.stringify({
        dataUrl: "data:image/png;base64,AA==",
        name: "../../evil name?.png",
      }),
    });

    expect(response.status).toBe(200);
    const payload = await response.json();
    const basename = path.basename(String(payload.path || ""));
    expect(basename).toMatch(/^evil-name-\d+-[0-9a-f-]+\.png$/);
    expect(basename).not.toContain("..");
    fs.unlinkSync(payload.path);
  });

  it("normalizes local PowerPoint image paths before bridge execution", async () => {
    const { baseUrl, bridge } = await startApiServer();
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "opencode-office-image-test-"));
    const imagePath = path.join(tempDir, "tiny.png");
    fs.writeFileSync(imagePath, Buffer.from([0]));
    const execute = vi.spyOn(bridge, "execute").mockResolvedValue({ result: { textResultForLlm: "ok" } });

    const response = await fetch(`${baseUrl}/api/office-tools/execute`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-office-bridge-token": bridge.bridgeToken,
      },
      body: JSON.stringify({
        host: "powerpoint",
        toolName: "manage_slide_media",
        args: { action: "insertImage", slideIndex: 0, imagePath },
      }),
    });

    expect(response.status).toBe(200);
    expect(execute).toHaveBeenCalledWith("powerpoint", "manage_slide_media", expect.objectContaining({
      action: "insertImage",
      slideIndex: 0,
      imageBase64: "AA==",
    }), bridge.bridgeToken);
    expect(execute.mock.calls[0]?.[2]).not.toHaveProperty("imagePath");
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it("normalizes local image paths in PowerPoint layout bindings", async () => {
    const { baseUrl, bridge } = await startApiServer();
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "opencode-office-binding-image-test-"));
    const imagePath = path.join(tempDir, "hero.jpg");
    fs.writeFileSync(imagePath, Buffer.from([1, 2]));
    const execute = vi.spyOn(bridge, "execute").mockResolvedValue({ result: { textResultForLlm: "ok" } });

    const response = await fetch(`${baseUrl}/api/office-tools/execute`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-office-bridge-token": bridge.bridgeToken,
      },
      body: JSON.stringify({
        host: "powerpoint",
        toolName: "create_slide_from_layout",
        args: { layoutId: "layout-1", bindings: [{ placeholderName: "Hero", imagePath }] },
      }),
    });

    expect(response.status).toBe(200);
    expect(execute).toHaveBeenCalledWith("powerpoint", "create_slide_from_layout", expect.objectContaining({
      bindings: [expect.objectContaining({ placeholderName: "Hero", imageBase64: "AQI=" })],
    }), bridge.bridgeToken);
    expect((execute.mock.calls[0]?.[2] as any).bindings[0]).not.toHaveProperty("imagePath");
    fs.rmSync(tempDir, { recursive: true, force: true });
  });
});

describe("bridge router hardening", () => {
  it("does not expose non-execute api routes on the http bridge", async () => {
    const { baseUrl } = await startBridgeServer();
    const statusResponse = await fetch(`${baseUrl}/api/opencode/status`);
    const sessionResponse = await fetch(`${baseUrl}/api/office-tools/session`);

    expect(statusResponse.status).toBe(404);
    expect(sessionResponse.status).toBe(404);
  });
});

describe("opencode config proxy", () => {
  it("forwards QA settings as JSON through the real runtime", async () => {
    const upstream = express.Router();
    upstream.use(express.json());
    upstream.patch("/config", (req, res) => {
      if (!req.is("application/json")) {
        res.status(415).send("Expected application/json");
        return;
      }
      res.json({
        body: req.body,
        contentType: req.get("content-type"),
        directory: req.get("x-opencode-directory"),
      });
    });
    const endpoint = await startServer(upstream);
    const runtime = new OpencodeRuntime();
    runtime.runtime = { baseUrl: `${endpoint.baseUrl}/api`, mode: "attached" };
    const { baseUrl } = await startServer(createApiRouter(runtime, new OfficeToolBridge()));
    const config = { agent: { "visual-qa": { model: "test/model", variant: "high" } } };
    const directory = path.resolve("test workspace");

    const response = await fetch(`${baseUrl}/api/opencode/config`, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        "x-opencode-directory": directory,
      },
      body: JSON.stringify(config),
    });

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual({
      body: config,
      contentType: "application/json",
      directory: encodeURIComponent(directory),
    });
  });
});

describe("directory-scoped opencode routing", () => {
  it("forwards request directory overrides to the runtime", async () => {
    const calls: Array<{ url: string; options: any }> = [];
    const { baseUrl } = await startApiServer(undefined, {
      request: async (url: string, options?: any) => {
        calls.push({ url, options });
        return { ok: true };
      },
    });

    const response = await fetch(`${baseUrl}/api/opencode/session`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-opencode-directory": "/tmp/folder",
      },
      body: JSON.stringify({ title: "test" }),
    });

    expect(response.status).toBe(200);
    expect(calls).toHaveLength(1);
    expect(calls[0]).toMatchObject({
      url: "/session",
      options: expect.objectContaining({
        directory: "/tmp/folder",
      }),
    });
  });

  it("uses the requested directory when filtering local session history", async () => {
    const { baseUrl } = await startApiServer(undefined, {
      directory: () => "/repo/root",
      request: async (url: string) => {
        if (String(url).startsWith("/session?")) {
          return [
            { id: "one", title: "Word: A", directory: "/repo/root" },
            { id: "two", title: "Word: B", directory: "/tmp/folder" },
          ];
        }
        return { ok: true };
      },
    });

    const response = await fetch(`${baseUrl}/api/opencode/sessions?host=word&directory=${encodeURIComponent("/tmp/folder")}`);

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual([
      expect.objectContaining({ id: "two", directory: "/tmp/folder" }),
    ]);
  });

  it.each(["canonical", "symlink"])("matches %s session paths for a symlinked directory", async (stored) => {
    const temp = fs.mkdtempSync(path.join(os.tmpdir(), "opencode-office-history-test-"));
    closers.push(async () => fs.rmSync(temp, { recursive: true, force: true }));
    const directory = path.join(temp, "project");
    const alias = path.join(temp, "alias");
    const other = path.join(temp, "other");
    fs.mkdirSync(directory);
    fs.mkdirSync(other);
    fs.symlinkSync(directory, alias, "junction");
    const canonical = fs.realpathSync(directory);
    const session = { id: "matching", title: "Word: A", directory: stored === "canonical" ? canonical : alias };
    const calls: string[] = [];
    const { baseUrl } = await startApiServer(undefined, {
      request: async (url: string) => {
        calls.push(url);
        return [
          session,
          { id: "other-host", title: "Excel: A", directory: canonical },
          { id: "other-directory", title: "Word: B", directory: other },
          { id: "missing-directory", title: "Word: C", directory: path.join(temp, "missing") },
        ];
      },
    });

    const response = await fetch(`${baseUrl}/api/opencode/sessions?host=word&directory=${encodeURIComponent(alias)}`);

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual([session]);
    expect(calls).toEqual([`/session?roots=true&limit=100&directory=${encodeURIComponent(canonical)}`]);
  });

  it("proxies file mention search requests", async () => {
    const calls: Array<{ url: string; options: any }> = [];
    const { baseUrl } = await startApiServer(undefined, {
      request: async (url: string, options?: any) => {
        calls.push({ url, options });
        return ["src/App.tsx"];
      },
    });

    const response = await fetch(
      `${baseUrl}/api/opencode/find/files?query=${encodeURIComponent("App")}&dirs=true&limit=15`,
      { headers: { "x-opencode-directory": "/repo" } },
    );

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual(["src/App.tsx"]);
    expect(calls[0]).toMatchObject({
      url: "/find/file?query=App&dirs=true&limit=15",
      options: expect.objectContaining({ directory: "/repo" }),
    });
  });
});
