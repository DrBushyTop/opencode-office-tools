import { createRequire } from "module";
import { describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);
const { OfficeToolBridge } = require("./officeToolBridge.js");

describe("office tool bridge", () => {
  it("requires bridge and session authentication", async () => {
    const bridge = new OfficeToolBridge();
    const sessionToken = bridge.issueClientSession();
    const { executorId } = bridge.register("word", sessionToken);

    await expect(bridge.execute("word", "get_document_content", {}, "bad-token")).rejects.toThrow(/Invalid Office bridge token/);

    const pending = bridge.execute("word", "get_document_content", {}, bridge.bridgeToken);
    const request = bridge.poll(executorId, sessionToken);
    expect(request.toolName).toBe("get_document_content");
    bridge.respond(request.id, executorId, sessionToken, { result: "ok" });

    await expect(pending).resolves.toEqual({ result: "ok" });
  });

  it("binds pending work to the registered executor", async () => {
    const bridge = new OfficeToolBridge();
    const sessionToken = bridge.issueClientSession();
    const { executorId } = bridge.register("word", sessionToken);
    const otherSession = bridge.issueClientSession();
    const { executorId: otherExecutorId } = bridge.register("excel", otherSession);

    const pending = bridge.execute("word", "get_document_content", {}, bridge.bridgeToken);
    const request = bridge.poll(executorId, sessionToken);

    expect(() => bridge.respond(request.id, otherExecutorId, otherSession, { result: "bad" })).toThrow(/does not belong/);
    bridge.respond(request.id, executorId, sessionToken, { result: "good" });
    await expect(pending).resolves.toEqual({ result: "good" });
  });

  it("keeps an assigned executor valid while work is pending", () => {
    const bridge = new OfficeToolBridge();
    const sessionToken = bridge.issueClientSession();
    const { executorId } = bridge.register("word", sessionToken);

    const pending = bridge.execute("word", "get_document_content", {}, bridge.bridgeToken);
    const request = bridge.poll(executorId, sessionToken);
    const executor = bridge.executors.get(executorId);
    executor.lastSeen = Date.now() - 20000;

    expect(() => bridge.respond(request.id, executorId, sessionToken, { result: "ok" })).not.toThrow();
    return expect(pending).resolves.toEqual({ result: "ok" });
  });

  it("does not allow another session to replace an active executor", () => {
    const bridge = new OfficeToolBridge();
    const sessionToken = bridge.issueClientSession();
    bridge.register("word", sessionToken);

    const otherSession = bridge.issueClientSession();
    expect(() => bridge.register("word", otherSession)).toThrow(/already registered/);
  });

  it("waits for work and resolves poll immediately when a request arrives", async () => {
    const bridge = new OfficeToolBridge();
    const sessionToken = bridge.issueClientSession();
    const { executorId } = bridge.register("word", sessionToken);

    const requestPromise = bridge.poll(executorId, sessionToken);
    const pending = bridge.execute("word", "get_document_content", {}, bridge.bridgeToken);
    const request = await requestPromise;

    expect(request.toolName).toBe("get_document_content");
    bridge.respond(request.id, executorId, sessionToken, { result: "ok" });
    await expect(pending).resolves.toEqual({ result: "ok" });
  });
});
