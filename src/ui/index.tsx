/// <reference types="@types/office-js" />

import * as React from "react";
import { createRoot } from "react-dom/client";
import { z } from "zod";
import { App } from "./App";
import { remoteLog } from "./lib/remoteLog";
import { applyDefaultTaskpaneWidth } from "./lib/taskpaneWidth";

declare global {
  interface Window {
    __opencodeTaskpaneBoot?: (entry: {
      level: "info" | "warn" | "error";
      tag: string;
      message: string;
      detail?: unknown;
      _skipRemote?: boolean;
    }) => void;
  }
}

const bootLevelSchema = z.enum(["info", "warn", "error"]);

const bootPayloadSchema = z.object({
  level: bootLevelSchema,
  tag: z.string().min(1),
  message: z.string().min(1),
  detail: z.unknown().optional(),
});

const officeReadySchema = z.object({
  host: z.string().optional(),
  platform: z.string().optional(),
  addin: z.string().optional(),
}).passthrough();

type BootLevel = z.infer<typeof bootLevelSchema>;
type BootPayload = z.infer<typeof bootPayloadSchema>;

function publishBoot(level: BootLevel, tag: string, message: string, detail?: unknown) {
  const payload = bootPayloadSchema.parse({ level, tag, message, detail });
  window.__opencodeTaskpaneBoot?.({ ...payload, _skipRemote: true });
  remoteLog(payload.tag, payload.message, payload.detail, payload.level);
}

function instrumentConsole() {
  const sinks = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  };

  console.log = (...args: unknown[]) => {
    sinks.log(...args);
    remoteLog("ui.console", "console.log", args, "info");
  };
  console.warn = (...args: unknown[]) => {
    sinks.warn(...args);
    remoteLog("ui.console", "console.warn", args, "warn");
  };
  console.error = (...args: unknown[]) => {
    sinks.error(...args);
    remoteLog("ui.console", "console.error", args, "error");
  };
}

function subscribeToWindowFailures() {
  window.addEventListener("error", (event) => {
    publishBoot("error", "ui.window", event.message || "Unhandled window error", {
      filename: event.filename,
      lineno: event.lineno,
      colno: event.colno,
      error: event.error,
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    publishBoot("error", "ui.promise", "Unhandled promise rejection", event.reason);
  });
}

type RenderGuardState = { failure: Error | null };

class RenderGuard extends React.Component<React.PropsWithChildren, RenderGuardState> {
  state: RenderGuardState = { failure: null };

  static getDerivedStateFromError(failure: Error): RenderGuardState {
    return { failure };
  }

  componentDidCatch(failure: Error, info: React.ErrorInfo) {
    remoteLog("ui.error-boundary", "React render failed", {
      error: failure,
      componentStack: info.componentStack,
    });
  }

  render() {
    if (!this.state.failure) {
      return this.props.children;
    }

    return (
      <section style={{ fontFamily: 'Inter, "Segoe UI", sans-serif', padding: 24, lineHeight: 1.6, color: "#241d17" }}>
        <h1 style={{ marginTop: 0, fontSize: 24 }}>OpenCode could not finish rendering</h1>
        <p>The task pane reached React, but a frontend error stopped the workspace from loading.</p>
        <p>Open the tray menu and inspect the debug log, or visit <code>/api/debug/logs</code> on the local server.</p>
        <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word", background: "#f4eee4", padding: 12, borderRadius: 10 }}>
          {this.state.failure.stack || this.state.failure.message}
        </pre>
      </section>
    );
  }
}

function waitForOfficeHost() {
  return new Promise<unknown>((resolve) => {
    Office.onReady((info) => resolve(info));
  });
}

function renderTaskpane() {
  const rootNode = document.getElementById("root");
  if (!rootNode) {
    publishBoot("error", "ui.bootstrap", "Root container not found");
    return;
  }

  publishBoot("info", "ui.bootstrap", "Rendering React task pane");
  createRoot(rootNode).render(
    <RenderGuard>
      <App />
    </RenderGuard>,
  );
}

async function startTaskpane() {
  publishBoot("info", "ui.bootstrap", "Bootstrap script loaded", {
    href: window.location.href,
    userAgent: navigator.userAgent,
  });

  if (typeof Office === "undefined") {
    publishBoot("error", "ui.bootstrap", "Office global missing before readiness wait");
    return;
  }

  publishBoot("info", "ui.bootstrap", "Waiting for Office host");
  const officeInfo = await waitForOfficeHost();
  const parsedInfo = officeReadySchema.safeParse(officeInfo);
  publishBoot("info", "ui.bootstrap", "Office host is ready", parsedInfo.success ? parsedInfo.data : officeInfo);

  applyDefaultTaskpaneWidth();
  renderTaskpane();
}

instrumentConsole();
subscribeToWindowFailures();

void startTaskpane().catch((error) => {
  publishBoot("error", "ui.bootstrap", "Task pane bootstrap failed", error);
  throw error;
});
