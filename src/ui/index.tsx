/// <reference types="@types/office-js" />

import * as React from "react";
import { createRoot } from "react-dom/client";
import { z } from "zod";
import { App } from "./App";
import { remoteLog } from "./lib/remoteLog";

declare global {
  interface Window {
    __opencodeBootstrapLog?: (level: "info" | "warn" | "error", tag: string, message: string, detail?: unknown) => void;
  }
}

const LogLevelSchema = z.enum(["info", "warn", "error"]);
type LogLevel = z.infer<typeof LogLevelSchema>;

const BootstrapLogSchema = z.object({
  level: LogLevelSchema,
  tag: z.string(),
  message: z.string(),
  detail: z.unknown().optional(),
});

type BootstrapLog = z.infer<typeof BootstrapLogSchema>;

const OfficeReadyInfoSchema = z.object({
  host: z.string().optional(),
  platform: z.string().optional(),
  addin: z.string().optional(),
}).passthrough();

const earlyClientLogs: BootstrapLog[] = [];

function pushEarlyLog(level: LogLevel, tag: string, message: string, detail?: unknown) {
  const payload = BootstrapLogSchema.parse({ level, tag, message, detail });
  earlyClientLogs.push(payload);
  window.__opencodeBootstrapLog?.(level, tag, message, detail);
  remoteLog(tag, message, detail, level);
}

function installConsoleForwarding() {
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

type ErrorBoundaryState = {
  error: Error | null;
};

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
      <div style={{ fontFamily: "Georgia, serif", padding: 24, lineHeight: 1.5 }}>
        <h1 style={{ marginTop: 0, fontSize: 24 }}>Add-in failed to load</h1>
        <p>The add-in hit a frontend error before it could render.</p>
        <p>
          Open the tray menu and inspect the debug log, or visit <code>/api/debug/logs</code> on the local server.
        </p>
        <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word", background: "#f5f1ea", padding: 12, borderRadius: 6 }}>
          {this.state.error.stack || this.state.error.message}
        </pre>
      </div>
    );
  }
}

window.addEventListener("error", (event) => {
  pushEarlyLog("error", "ui.window", event.message || "Unhandled window error", {
    filename: event.filename,
    lineno: event.lineno,
    colno: event.colno,
    error: event.error,
  });
});

window.addEventListener("unhandledrejection", (event) => {
  pushEarlyLog("error", "ui.promise", "Unhandled promise rejection", event.reason);
});

installConsoleForwarding();
pushEarlyLog("info", "ui.bootstrap", "Bootstrap script loaded", {
  href: window.location.href,
  userAgent: navigator.userAgent,
});

if (typeof Office === "undefined") {
  pushEarlyLog("error", "ui.bootstrap", "Office global missing before onReady registration");
} else {
  pushEarlyLog("info", "ui.bootstrap", "Registering Office.onReady handler");
}

try {
  Office.onReady((info) => {
    const readyInfo = OfficeReadyInfoSchema.safeParse(info);
    pushEarlyLog("info", "ui.bootstrap", "Office.onReady resolved", readyInfo.success ? readyInfo.data : info);
    const container = document.getElementById("root");
    if (container) {
      pushEarlyLog("info", "ui.bootstrap", "Root container found, rendering app");
      const root = createRoot(container);
      root.render(
        <ErrorBoundary>
          <App />
        </ErrorBoundary>
      );
    } else {
      pushEarlyLog("error", "ui.bootstrap", "Root container not found");
    }
    console.log("Add-in loaded successfully");
  });
} catch (error) {
  pushEarlyLog("error", "ui.bootstrap", "Failed to register Office.onReady", error);
  throw error;
}
