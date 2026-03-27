/// <reference types="@types/office-js" />

import * as React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { remoteLog } from "./lib/remoteLog";

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
  remoteLog("ui.window", event.message || "Unhandled window error", {
    filename: event.filename,
    lineno: event.lineno,
    colno: event.colno,
    error: event.error,
  });
});

window.addEventListener("unhandledrejection", (event) => {
  remoteLog("ui.promise", "Unhandled promise rejection", event.reason);
});

Office.onReady(() => {
  remoteLog("ui.bootstrap", "Office.onReady resolved", undefined, "info");
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    );
  } else {
    remoteLog("ui.bootstrap", "Root container not found");
  }
  console.log("Add-in loaded successfully");
});
