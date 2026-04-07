import express from "express";
import fs from "node:fs";
import path from "node:path";
import { createServer, type Server } from "node:http";
import { createRequire } from "module";
import { afterEach, describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);
const { createApiRouter, createBridgeRouter } = require("./api.js");
const { OfficeToolBridge } = require("./officeToolBridge.js");

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
});
