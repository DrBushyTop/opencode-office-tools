import { useState, useEffect, useRef, useMemo } from "react";
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  makeStyles,
} from "@fluentui/react-components";
import { ChatInput, ImageAttachment } from "./components/ChatInput";
import { Message, MessageList, DebugEvent } from "./components/MessageList";
import { HeaderBar, ModelType } from "./components/HeaderBar";
import { SessionHistory } from "./components/SessionHistory";
import { PermissionDialog, type PermissionDecision } from "./components/PermissionDialog";
import { useIsDarkMode } from "./useIsDarkMode";
import { useLocalStorage } from "./useLocalStorage";
import { createOpencodeClient, ModelInfo, SessionInfo } from "./lib/opencode-client";
import { createOfficeToolBridge } from "./lib/office-tool-bridge";
import { makeSessionTitle, restoreSession, updateSessionTitle, type OpencodeSessionInfo } from "./lib/opencode-session-history";
import { trafficStats } from "./lib/opencode-events";
import { getOfficeToolExecutor, getToolNamesForHost } from "./tools";
import { canAutoApprove, type OfficePermissionRequest } from "../shared/office-permissions";
import { formatOfficeToolActivity } from "../shared/office-tool-registry";
import {
  SavedSession,
  OfficeHost,
  getHostFromOfficeHost,
} from "./sessionStorage";
import React from "react";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    padding: "10px",
    boxSizing: "border-box",
    background: "var(--oc-page)",
    color: "var(--oc-text)",
    fontFamily: '"Inter", "Segoe UI", sans-serif',
  },
  shell: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    borderRadius: "14px",
    background: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
  },
  error: {
    margin: "0 12px 8px",
    padding: "10px 12px",
    borderRadius: "10px",
    background: "var(--oc-danger-bg)",
    border: "1px solid var(--oc-danger-border)",
    color: "var(--oc-danger-text)",
  },
});

function getHostLabel(host: OfficeHost) {
  return host === "powerpoint" ? "PowerPoint" : host === "excel" ? "Excel" : host === "onenote" ? "OneNote" : "Word";
}

function getSurfaceVars(isDarkMode: boolean): React.CSSProperties {
  return {
    "--oc-page": isDarkMode ? "#131010" : "#f3f3f3",
    "--oc-bg": isDarkMode ? "#1b1818" : "#fcfcfc",
    "--oc-bg-strong": isDarkMode ? "#252121" : "#f8f8f8",
    "--oc-bg-soft": isDarkMode ? "rgba(255,255,255,0.05)" : "rgba(0,0,0,0.03)",
    "--oc-bg-soft-hover": isDarkMode ? "rgba(255,255,255,0.08)" : "rgba(0,0,0,0.05)",
    "--oc-border": isDarkMode ? "rgba(255,255,255,0.10)" : "#e5e5e5",
    "--oc-border-strong": isDarkMode ? "rgba(255,255,255,0.16)" : "rgba(0,0,0,0.14)",
    "--oc-text": isDarkMode ? "#f1ecec" : "#171717",
    "--oc-text-muted": isDarkMode ? "#b7b1b1" : "#6f6f6f",
    "--oc-text-faint": isDarkMode ? "#7f7979" : "#8f8f8f",
    "--oc-accent": isDarkMode ? "#89b5ff" : "#034cff",
    "--oc-accent-strong": isDarkMode ? "#2558d0" : "#0443de",
    "--oc-accent-bg": isDarkMode ? "rgba(137,181,255,0.16)" : "#ecf3ff",
    "--oc-shadow": isDarkMode
      ? "0 0 0 1px rgba(255,255,255,0.06), 0 16px 48px rgba(0,0,0,0.24)"
      : "0 0 0 1px rgba(0,0,0,0.05), 0 16px 48px rgba(0,0,0,0.06)",
    "--oc-danger-bg": isDarkMode ? "rgba(252, 83, 58, 0.14)" : "#fff2f0",
    "--oc-danger-border": isDarkMode ? "rgba(252, 83, 58, 0.28)" : "rgba(252, 83, 58, 0.24)",
    "--oc-danger-text": isDarkMode ? "#fe806a" : "#ed4831",
  } as React.CSSProperties;
}

const FALLBACK_MODELS: ModelInfo[] = [
  {
    key: "anthropic/claude-sonnet-4-5",
    label: "Anthropic / Claude Sonnet 4.5",
    providerID: "anthropic",
    modelID: "claude-sonnet-4-5",
  },
];

function pickDefaultModel(models: { key: string }[]): ModelType {
  const preferred = ["anthropic/claude-sonnet-4-6", "anthropic/claude-sonnet-4-5"];
  for (const id of preferred) {
    if (models.some((model) => model.key === id)) return id;
  }
  return models[0]?.key || FALLBACK_MODELS[0].key;
}

// Per-host tool guidance and verification instructions are now defined in the
// OpenCode agent prompts under .opencode/agents/{powerpoint,word,excel,onenote}.md.
// The system message passed per-call is kept minimal; the agent prompt is primary.

function getEnabledTools(host: OfficeHost) {
  const tools = Object.fromEntries(
    getToolNamesForHost(host).map((name) => [name, true]),
  ) as Record<string, boolean>;

  tools.task = true;

  return tools;
}

async function getSessionFamily(client: ReturnType<typeof createOpencodeClient>, rootId: string) {
  const ids = new Set<string>([rootId]);
  const titles = new Map<string, string>();
  const queue = [rootId];

  const root = await client.getSession(rootId).catch(() => null);
  if (root?.title) titles.set(root.id, root.title);

  while (queue.length > 0) {
    const id = queue.shift()!;
    const children = await client.getSessionChildren(id).catch(() => [] as SessionInfo[]);
    for (const child of children) {
      if (!child?.id || ids.has(child.id)) continue;
      ids.add(child.id);
      if (child.title) titles.set(child.id, child.title);
      queue.push(child.id);
    }
  }

  return { ids, titles };
}

function describeToolActivity(toolName: string, toolArgs: Record<string, unknown>) {
  if (toolName === "task") {
    const subagentType = typeof toolArgs.subagent_type === "string" ? toolArgs.subagent_type : "subagent";
    const description = typeof toolArgs.description === "string" ? toolArgs.description : "Working";
    return `Launching ${subagentType}: ${description}`;
  }

  return formatOfficeToolActivity(toolName, toolArgs) || `Calling ${toolName}...`;
}

function redactSensitiveFields(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(redactSensitiveFields);
  }
  if (!value || typeof value !== "object") {
    return value;
  }

  const sensitiveKeys = new Set(["password", "token", "secret", "apikey", "api_key"]);
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, innerValue]) => [
      key,
      sensitiveKeys.has(key.toLowerCase()) ? "[REDACTED]" : redactSensitiveFields(innerValue),
    ]),
  );
}

function previewEvent(eventType: string, data: Record<string, unknown>) {
  if (eventType === "assistant.message_delta") return String(data.deltaContent || "").slice(0, 80);
  if (eventType === "assistant.message") return String(data.content || "").slice(0, 80);
  if (eventType === "tool.execution_start") {
    const toolName = String(data.toolName || "");
    if (toolName === "task") {
      const args = (data.arguments || {}) as Record<string, unknown>;
      const subagentType = typeof args.subagent_type === "string" ? args.subagent_type : "subagent";
      const description = typeof args.description === "string" ? args.description : "Working";
      return `${subagentType}: ${description}`.slice(0, 100);
    }
    return toolName;
  }
  if (eventType === "session.error") return String(data.message || "").slice(0, 80);
  return JSON.stringify(data).slice(0, 100);
}

function getSystemMessage(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const hostName = host === Office.HostType.PowerPoint
    ? "PowerPoint"
    : host === Office.HostType.Word
      ? "Word"
      : host === Office.HostType.Excel
        ? "Excel"
        : host === Office.HostType.OneNote
          ? "OneNote"
        : "Office";

  return `The user's Microsoft ${hostName} document is currently open. Always operate on the open document through the available tools.`;
}

function toPromptParts(text: string, images: Array<{ path: string; name: string; mime: string }>) {
  return [
    { type: "text" as const, text: text || "Here are some images for you to analyze." },
    ...images.map((image) => ({
      type: "file" as const,
      filename: image.name,
      mime: image.mime,
      url: `file://${image.path}`,
    })),
  ];
}

export const App: React.FC = () => {
  const styles = useStyles();
  const client = useMemo(() => createOpencodeClient(), []);
  const [availableModels, setAvailableModels] = useState(FALLBACK_MODELS);
  const [messages, setMessages] = useState<Message[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [images, setImages] = useState<ImageAttachment[]>([]);
  const [isTyping, setIsTyping] = useState(false);
  const [currentActivity, setCurrentActivity] = useState<string>("");
  const [streamingText, setStreamingText] = useState<string>("");
  const [debugEvents, setDebugEvents] = useState<DebugEvent[]>([]);
  const [error, setError] = useState("");
  const [selectedModel, setSelectedModel] = useLocalStorage<ModelType>("word-addin-selected-model", "");
  const [showHistory, setShowHistory] = useState(false);
  const [currentSessionId, setCurrentSessionId] = useState<string>("");
  const [officeHost, setOfficeHost] = useState<OfficeHost>("word");
  const [debugEnabled, setDebugEnabled] = useLocalStorage<boolean>("opencode-debug", false);
  const [sharedHistory, setSharedHistory] = useLocalStorage<boolean>("opencode-shared-history", false);
  const [runtimeMode, setRuntimeMode] = useState<string>("");
  const [permission, setPermission] = useState<OfficePermissionRequest | null>(null);
  const [permissionSessionTitle, setPermissionSessionTitle] = useState<string | null>(null);
  const isDarkMode = useIsDarkMode();
  const hostLabel = getHostLabel(officeHost);
  const sessionCreatedAt = useRef<string>("");
  const started = useRef(false);
  const run = useRef<AbortController | null>(null);

  const fetchModels = async () => {
    try {
      const status = await client.getStatus();
      setRuntimeMode(status.mode);
      const models = status.models?.length ? status.models : FALLBACK_MODELS;
      setAvailableModels(models);
      if (!selectedModel || !models.some((model) => model.key === selectedModel)) {
        setSelectedModel(pickDefaultModel(models));
      }
    } catch {
      setAvailableModels(FALLBACK_MODELS);
      if (!selectedModel || !FALLBACK_MODELS.some((model) => model.key === selectedModel)) {
        setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
      }
    }
  };

  useEffect(() => {
    void fetchModels();
  }, []);

  useEffect(() => {
    const host = getHostFromOfficeHost(Office.context.host);
    setOfficeHost(host);
    const bridge = createOfficeToolBridge(host, getOfficeToolExecutor(Office.context.host));
    return () => {
      void bridge.stop();
    };
  }, []);

  useEffect(() => {
    const poll = async () => {
      try {
        if (!currentSessionId) {
          setPermission(null);
          setPermissionSessionTitle(null);
          return;
        }

        const response = await fetch("/api/opencode/permissions");
        if (!response.ok) return;
        const items = await response.json();
        const family = await getSessionFamily(client, currentSessionId);
        const next = items.find((item: OfficePermissionRequest) => family.ids.has(item.sessionID));

        if (!next) {
          setPermission(null);
          setPermissionSessionTitle(null);
          return;
        }

        if (canAutoApprove(next)) {
          await fetch(`/api/opencode/permission/${next.id}/reply`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ reply: "once" }),
          });
          return;
        }
        setPermission((current) => current?.id === next.id ? current : next);
        setPermissionSessionTitle(family.titles.get(next.sessionID) || null);
      } catch {}
    };

    const timer = window.setInterval(poll, 1000);
    void poll();
    return () => window.clearInterval(timer);
  }, [client, currentSessionId]);

  const handlePermissionDecision = async (decision: PermissionDecision) => {
    if (!permission) return;
    const reply = decision === "deny" ? "reject" : decision === "always" ? "always" : "once";
    await fetch(`/api/opencode/permission/${permission.id}/reply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reply }),
    }).catch(() => undefined);
    setPermission(null);
    setPermissionSessionTitle(null);
  };

  const startNewSession = async (_modelKey: ModelType, restored?: SavedSession) => {
    setError("");
    setShowHistory(false);
    setStreamingText("");
    setCurrentActivity("");
    setDebugEvents([]);
    setIsTyping(false);
    setPermission(null);
    setPermissionSessionTitle(null);

    const host = Office.context.host;
    const office = getHostFromOfficeHost(host);
    setOfficeHost(office);

    setCurrentSessionId("");
    sessionCreatedAt.current = restored?.createdAt || new Date().toISOString();
    setMessages(restored?.messages || []);

    if (restored) {
      setCurrentSessionId(restored.id);
      const status = await client.getStatus().catch(() => null);
      if (status?.mode) setRuntimeMode(status.mode);
      return;
    }

    try {
      const session = await client.createSession({ title: `${office === "powerpoint" ? "PowerPoint" : office === "excel" ? "Excel" : office === "onenote" ? "OneNote" : "Word"}: New chat` });
      setCurrentSessionId(session.id);
    } catch (err) {
      setError(`Failed to create session: ${err instanceof Error ? err.message : String(err)}`);
    }

    const status = await client.getStatus().catch(() => null);
    if (status?.mode) setRuntimeMode(status.mode);
    if (!selectedModel && availableModels.length > 0) {
      setSelectedModel(pickDefaultModel(availableModels));
    }
  };

  const handleRestoreSession = async (saved: OpencodeSessionInfo) => {
    const restored = await restoreSession(saved.id, selectedModel);
    void startNewSession(selectedModel, restored);
  };

  useEffect(() => {
    if (!selectedModel || started.current) return;
    started.current = true;
    void startNewSession(selectedModel);
  }, [selectedModel]);

  const handleModelChange = (model: ModelType) => {
    setSelectedModel(model);
    if (!currentSessionId) {
      void startNewSession(model);
    }
  };

  const handleSend = async () => {
    if (isTyping || (!inputValue.trim() && images.length === 0) || !currentSessionId) return;

    const userInput = inputValue;
    const userImages = [...images];
    const isFirstUserTurn = !messages.some((message) => message.sender === "user");

    setMessages((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        text: userInput || `Sent ${userImages.length} image${userImages.length === 1 ? "" : "s"}`,
        sender: "user",
        timestamp: new Date(),
        images: userImages.length > 0 ? userImages.map((image) => ({ dataUrl: image.dataUrl, name: image.name })) : undefined,
      },
    ]);

    setInputValue("");
    setImages([]);
    setIsTyping(true);
    setCurrentActivity("Processing...");
    setStreamingText("");
    setDebugEvents([]);
    setError("");
    trafficStats.reset();

    try {
      const ctl = new AbortController();
      run.current = ctl;
      const uploads: Array<{ path: string; name: string; mime: string }> = [];

      for (const image of userImages) {
        const response = await fetch("/api/upload-image", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ dataUrl: image.dataUrl, name: image.name }),
        });

        if (!response.ok) {
          throw new Error(`Failed to upload image: ${response.statusText}`);
        }

        const result = await response.json();
        uploads.push({ path: result.path, name: image.name, mime: result.mime || "image/png" });
      }

      const model = availableModels.find((item) => item.key === selectedModel) || FALLBACK_MODELS[0];
      const parts = toPromptParts(userInput, uploads);
      const tools = getEnabledTools(officeHost);

      if (isFirstUserTurn && userInput.trim()) {
        updateSessionTitle(currentSessionId, makeSessionTitle(officeHost, userInput)).catch(() => undefined);
      }

      let count = 0;

      for await (const event of client.query(currentSessionId, {
        model: { providerID: model.providerID, modelID: model.modelID },
        agent: officeHost,
        system: getSystemMessage(Office.context.host),
        parts,
        tools,
      }, { signal: ctl.signal })) {
        count += 1;
        const data = event.data || {};
        const preview = previewEvent(event.type, data);

        setDebugEvents((prev) => [...prev, { type: event.type, preview, timestamp: Date.now() }]);

        if (event.type === "assistant.message_delta") {
          setStreamingText((prev) => prev + String(data.deltaContent || ""));
          setCurrentActivity("");
          continue;
        }

        if (event.type === "assistant.message") {
          const text = String(data.content || "");
          setStreamingText("");
          setCurrentActivity("");
          if (text) {
            setMessages((prev) => [...prev, {
              id: String(event.id || crypto.randomUUID()),
              text,
              sender: "assistant",
              timestamp: new Date(event.timestamp || Date.now()),
            }]);
          }
          continue;
        }

        if (event.type === "tool.execution_start") {
          const toolName = String(data.toolName || "tool");
          const toolArgs = (data.arguments || {}) as Record<string, unknown>;
          const safeToolArgs = redactSensitiveFields(toolArgs) as Record<string, unknown>;
          setCurrentActivity(describeToolActivity(toolName, toolArgs));
          setMessages((prev) => [...prev, {
            id: String(event.id || crypto.randomUUID()),
            text: JSON.stringify(safeToolArgs, null, 2),
            sender: "tool",
            toolName,
            toolArgs: safeToolArgs,
            timestamp: new Date(),
          }]);
          continue;
        }

        if (event.type === "tool.execution_complete") {
          if (data.error) {
            const text = String(data.error);
            setCurrentActivity("");
            setMessages((prev) => [...prev, {
              id: `tool-error-${Date.now()}`,
              text: `Tool failed: ${text}`,
              sender: "assistant",
              timestamp: new Date(),
            }]);
            continue;
          }
          setCurrentActivity("Processing result...");
          continue;
        }

        if (event.type === "assistant.reasoning_delta") {
          setCurrentActivity("Thinking...");
          continue;
        }

        if (event.type === "assistant.turn_start") {
          setCurrentActivity("Starting response...");
          continue;
        }

        if (event.type === "assistant.turn_end") {
          setCurrentActivity("");
          setStreamingText("");
          continue;
        }

        if (event.type === "session.error") {
          const text = String(data.message || "Unknown session error");
          setMessages((prev) => [...prev, {
            id: `error-${Date.now()}`,
            text: `⚠️ Session error: ${text}`,
            sender: "assistant",
            timestamp: new Date(),
          }]);
        }
      }

      if (count === 0 && !ctl.signal.aborted) {
        setMessages((prev) => [...prev, {
          id: `debug-${Date.now()}`,
          text: "⚠️ No events received from the OpenCode runtime.",
          sender: "assistant",
          timestamp: new Date(),
        }]);
      }
    } catch (err) {
      const text = err instanceof Error ? err.message : String(err);
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `❌ Error: ${text}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
    } finally {
      run.current = null;
      setIsTyping(false);
      setCurrentActivity("");
    }
  };

  const handleStop = async () => {
    if (!currentSessionId || !isTyping) return;
    setCurrentActivity("Stopping...");

    try {
      await client.abortSession(currentSessionId);
      run.current?.abort();
      setStreamingText("");
    } catch (err) {
      const text = err instanceof Error ? err.message : String(err);
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `❌ Error stopping session: ${text}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
      setCurrentActivity("");
    }
  };

  if (showHistory) {
    return (
      <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
        <div style={getSurfaceVars(isDarkMode)}>
          <SessionHistory
            host={officeHost}
            shared={sharedHistory}
            onSharedChange={setSharedHistory}
            onSelectSession={handleRestoreSession}
            onClose={() => setShowHistory(false)}
          />
        </div>
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
      <div className={styles.container} style={getSurfaceVars(isDarkMode)}>
        <div className={styles.shell}>
          <HeaderBar
            onNewChat={() => {
              setCurrentSessionId("");
              void startNewSession(selectedModel);
            }}
            onShowHistory={() => setShowHistory(true)}
            selectedModel={selectedModel}
            onModelChange={handleModelChange}
            models={availableModels}
            debugEnabled={debugEnabled}
            onDebugChange={setDebugEnabled}
            subtitle={runtimeMode ? `${hostLabel} • ${runtimeMode} • ${getToolNamesForHost(officeHost).length} tools` : `${hostLabel} • ${getToolNamesForHost(officeHost).length} tools`}
          />

          <MessageList
            messages={messages}
            isTyping={isTyping}
            isConnecting={!currentSessionId && !error}
            currentActivity={currentActivity}
            streamingText={streamingText}
            debugEvents={debugEnabled ? debugEvents : undefined}
            hostLabel={hostLabel}
          />

          {error && <div className={styles.error}>{error}</div>}

          <ChatInput
            value={inputValue}
            onChange={setInputValue}
            onSend={handleSend}
            onStop={handleStop}
            isRunning={isTyping}
            images={images}
            onImagesChange={setImages}
          />
          {permission && (
            <PermissionDialog
              request={permission}
              cwd={null}
              sessionTitle={permissionSessionTitle}
              onDecision={handlePermissionDecision}
            />
          )}
        </div>
      </div>
    </FluentProvider>
  );
};
