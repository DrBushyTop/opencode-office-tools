/// <reference types="@types/office-js" />

import * as React from "react";
import { createRoot } from "react-dom/client";
import { z } from "zod";
import { App } from "./App";
import { remoteLog } from "./lib/remoteLog";
import { applyDefaultTaskpaneWidth } from "./lib/taskpaneWidth";

declare global {
  interface Window {
    __opencodeTaskpaneBoot?: (entry: { level: "info" | "warn" | "error"; tag: string; message: string; detail?: unknown; _skipRemote?: boolean }) => void;
  }
}

const LogLevelSchema = z.enum(["info", "warn", "error"]);
type LogLevel = z.infer<typeof LogLevelSchema>;

const BootEventSchema = z.object({
  level: LogLevelSchema,
  tag: z.string(),
  message: z.string(),
  detail: z.unknown().optional(),
});

const OfficeReadyInfoSchema = z.object({
  host: z.string().optional(),
  platform: z.string().optional(),
  addin: z.string().optional(),
}).passthrough();

function emitBootEvent(level: LogLevel, tag: string, message: string, detail?: unknown) {
  const event = BootEventSchema.parse({ level, tag, message, detail });
  // Pass _skipRemote so the inline boot monitor updates its UI but does NOT
  // call sendBootLog — remoteLog below is the single remote transport,
  // preventing duplicate POST /api/log requests and double terminal output.
  window.__opencodeTaskpaneBoot?.({ ...event, _skipRemote: true });
  remoteLog(tag, message, detail, level);
}

function wireConsoleTelemetry() {
  const original = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  };

  console.log = (...args: unknown[]) => {
    original.log(...args);
    remoteLog("ui.console", "console.log", args, "info");
  };
  console.warn = (...args: unknown[]) => {
    original.warn(...args);
    remoteLog("ui.console", "console.warn", args, "warn");
  };
  console.error = (...args: unknown[]) => {
    original.error(...args);
    remoteLog("ui.console", "console.error", args, "error");
  };
}

function installGlobalErrorTelemetry() {
  window.addEventListener("error", (event) => {
    emitBootEvent("error", "ui.window", event.message || "Unhandled window error", {
      filename: event.filename,
      lineno: event.lineno,
      colno: event.colno,
      error: event.error,
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    emitBootEvent("error", "ui.promise", "Unhandled promise rejection", event.reason);
  });
}

type ErrorBoundaryState = { error: Error | null };

class ErrorBoundary extends React.Component<React.PropsWithChildren, ErrorBoundaryState> {
  state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    remoteLog("ui.error-boundary", "React render failed", { error, componentStack: info.componentStack });
  }

  render() {
    if (!this.state.error) {
      return this.props.children;
    }

    return (
      <div style={{ fontFamily: 'Inter, "Segoe UI", sans-serif', padding: 24, lineHeight: 1.6, color: "#241d17" }}>
        <h1 style={{ marginTop: 0, fontSize: 24 }}>OpenCode could not finish rendering</h1>
        <p>The task pane reached React, but a frontend error stopped the workspace from loading.</p>
        <p>Open the tray menu and inspect the debug log, or visit <code>/api/debug/logs</code> on the local server.</p>
        <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word", background: "#f4eee4", padding: 12, borderRadius: 10 }}>
          {this.state.error.stack || this.state.error.message}
        </pre>
      </div>
    );
  }
}

function waitForOfficeReady() {
  return new Promise<unknown>((resolve) => {
    Office.onReady((info) => resolve(info));
  });
}

function mountApp() {
  const container = document.getElementById("root");
  if (!container) {
    emitBootEvent("error", "ui.bootstrap", "Root container not found");
    return;
  }

  emitBootEvent("info", "ui.bootstrap", "Rendering React task pane");
  const root = createRoot(container);
  root.render(
    <ErrorBoundary>
      <App />
    </ErrorBoundary>,
  );
}

async function bootstrapTaskpane() {
  emitBootEvent("info", "ui.bootstrap", "Bootstrap script loaded", {
    href: window.location.href,
    userAgent: navigator.userAgent,
  });

  if (typeof Office === "undefined") {
    emitBootEvent("error", "ui.bootstrap", "Office global missing before readiness wait");
    return;
  }

  emitBootEvent("info", "ui.bootstrap", "Waiting for Office host");
  const info = await waitForOfficeReady();
  const readyInfo = OfficeReadyInfoSchema.safeParse(info);
  emitBootEvent("info", "ui.bootstrap", "Office host is ready", readyInfo.success ? readyInfo.data : info);
  applyDefaultTaskpaneWidth();
  mountApp();
}

wireConsoleTelemetry();
installGlobalErrorTelemetry();

void bootstrapTaskpane().catch((error) => {
  emitBootEvent("error", "ui.bootstrap", "Task pane bootstrap failed", error);
  throw error;
});
