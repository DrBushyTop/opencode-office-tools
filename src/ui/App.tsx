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
import { createOpencodeClient, ModelInfo, OpencodeConfig, SessionInfo } from "./lib/opencode-client";
import { createOfficeToolBridge, readPowerPointContextSnapshot } from "./lib/office-tool-bridge";
import { carry, makeSessionTitle, mapAssistantParts, restoreSession, updateSessionTitle, type OpencodeSessionInfo } from "./lib/opencode-session-history";
import { trafficStats } from "./lib/opencode-events";
import { getOfficeToolExecutor, getToolNamesForHost } from "./tools";
import { setPowerPointContextSnapshot, type PowerPointContextSnapshot } from "./tools/powerpointContext";
import { canAutoApprove, type OfficePermissionRequest } from "../shared/office-permissions";
import { formatOfficeToolActivity } from "../shared/office-tool-registry";
import {
  SavedSession,
  OfficeHost,
  getHostFromOfficeHost,
} from "./sessionStorage";
import React from "react";
import { z } from "zod";

const NavigationTargetSchema = z.enum([
  "edit-selection",
  "insert-timeline",
  "insert-estimate-table",
]);
const PowerPointContextSnapshotSchema = z.object({
  activeSlideIndex: z.number().int().nonnegative().optional(),
  selectedSlideIds: z.array(z.string()),
  selectedShapeIds: z.array(z.string()),
}) satisfies z.ZodType<PowerPointContextSnapshot>;
const OfficePermissionRequestSchema = z.object({
  id: z.string().min(1),
  sessionID: z.string().min(1),
  permission: z.string().min(1),
  patterns: z.array(z.string()),
  metadata: z.record(z.string(), z.unknown()),
  always: z.array(z.string()),
  tool: z.object({
    messageID: z.string().min(1),
    callID: z.string().min(1),
  }).optional(),
}) satisfies z.ZodType<OfficePermissionRequest>;
const UploadImageResponseSchema = z.object({
  path: z.string().min(1),
  mime: z.string().optional(),
});
const PersistedBooleanSchema = z.boolean();
const PersistedModelSchema = z.string();

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

function mergeLiveMessages(
  previous: Message[],
  next: Message,
) {
  const index = previous.findIndex((item) => item.id === next.id);
  if (index === -1) return [...previous, next];
  return previous.map((item, current) => current === index ? { ...item, ...next } : item);
}

function qaModel(config: OpencodeConfig) {
  const value = config.agent?.["visual-qa"]?.model;
  return typeof value === "string" ? value : "";
}

function mergeQaModel(config: OpencodeConfig, model: ModelType) {
  const next = {
    ...config,
    agent: {
      ...(config.agent || {}),
      "visual-qa": {
        ...(config.agent?.["visual-qa"] || {}),
      },
    },
  } satisfies OpencodeConfig;

  if (model) {
    next.agent!["visual-qa"] = {
      ...next.agent!["visual-qa"],
      model,
    };
    return next;
  }

  if (next.agent?.["visual-qa"]) {
    const visual = { ...next.agent["visual-qa"] };
    delete visual.model;
    next.agent["visual-qa"] = visual;
  }

  return next;
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

  const base = `The user's Microsoft ${hostName} document is currently open. Always operate on the open document through the available tools.`;
  if (host !== Office.HostType.PowerPoint) return base;

  const snapshot = (() => {
    try {
      const raw = localStorage.getItem("opencode-powerpoint-context");
      if (!raw) return null;
      const parsed = PowerPointContextSnapshotSchema.safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data : null;
    } catch {
      return null;
    }
  })();
  if (!snapshot) return base;

  const contextBits = [
    snapshot.activeSlideIndex !== undefined ? `active slide index: ${snapshot.activeSlideIndex}` : "active slide index: unknown",
    `selected slide ids: ${snapshot.selectedSlideIds.length ? snapshot.selectedSlideIds.join(", ") : "none"}`,
    `selected shape ids: ${snapshot.selectedShapeIds.length ? snapshot.selectedShapeIds.join(", ") : "none"}`,
  ];
  return `${base} Current PowerPoint context: ${contextBits.join("; ")}.`;
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
  const [debugEvents, setDebugEvents] = useState<DebugEvent[]>([]);
  const [error, setError] = useState("");
  const [selectedModel, setSelectedModel] = useLocalStorage<ModelType>("word-addin-selected-model", "");
  const [showHistory, setShowHistory] = useState(false);
  const [currentSessionId, setCurrentSessionId] = useState<string>("");
  const [officeHost, setOfficeHost] = useState<OfficeHost>("word");
  const [debugEnabled, setDebugEnabled] = useLocalStorage<boolean>("opencode-debug", false);
  const [showThinking, setShowThinking] = useLocalStorage<boolean>("opencode-show-thinking", true);
  const [showToolResponses, setShowToolResponses] = useLocalStorage<boolean>("opencode-show-tool-responses", false);
  const [sharedHistory, setSharedHistory] = useLocalStorage<boolean>("opencode-shared-history", false);
  const [qaSubagentModel, setQaSubagentModel] = useState<ModelType>("");
  const [runtimeMode, setRuntimeMode] = useState<string>("");
  const [permission, setPermission] = useState<OfficePermissionRequest | null>(null);
  const [permissionSessionTitle, setPermissionSessionTitle] = useState<string | null>(null);
  const [liveMessages, setLiveMessages] = useState<Message[]>([]);
  const [pptContext, setPptContext] = useState<PowerPointContextSnapshot | null>(null);
  const isDarkMode = useIsDarkMode();
  const safeSelectedModel = PersistedModelSchema.catch("").parse(selectedModel);
  const safeDebugEnabled = PersistedBooleanSchema.catch(false).parse(debugEnabled);
  const safeShowThinking = PersistedBooleanSchema.catch(true).parse(showThinking);
  const safeShowToolResponses = PersistedBooleanSchema.catch(false).parse(showToolResponses);
  const safeSharedHistory = PersistedBooleanSchema.catch(false).parse(sharedHistory);
  const safeQaSubagentModel = PersistedModelSchema.catch("").parse(qaSubagentModel);
  const hostLabel = getHostLabel(officeHost);
  const sessionCreatedAt = useRef<string>("");
  const run = useRef<AbortController | null>(null);
  const liveRef = useRef<Message[]>([]);

  useEffect(() => {
    liveRef.current = liveMessages;
  }, [liveMessages]);

  const fetchModels = async () => {
    try {
      const status = await client.getStatus();
      setRuntimeMode(status.mode);
      const models = status.models?.length ? status.models : FALLBACK_MODELS;
      setAvailableModels(models);
      if (!safeSelectedModel || !models.some((model) => model.key === safeSelectedModel)) {
        setSelectedModel(pickDefaultModel(models));
      }
    } catch {
      setAvailableModels(FALLBACK_MODELS);
      if (!safeSelectedModel || !FALLBACK_MODELS.some((model) => model.key === safeSelectedModel)) {
        setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
      }
    }

    try {
      const config = await client.getConfig();
      setQaSubagentModel(qaModel(config));
    } catch {
      setQaSubagentModel("");
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
    if (officeHost !== "powerpoint") {
      setPptContext(null);
      setPowerPointContextSnapshot(null);
      localStorage.removeItem("opencode-powerpoint-context");
      return;
    }

    const update = async () => {
      const snapshot = await readPowerPointContextSnapshot();
      setPptContext(snapshot);
      setPowerPointContextSnapshot(snapshot);
      if (snapshot) {
        localStorage.setItem("opencode-powerpoint-context", JSON.stringify(snapshot));
      }
    };

    const timer = window.setInterval(() => {
      void update();
    }, 2000);
    void update();
    return () => window.clearInterval(timer);
  }, [officeHost]);

  useEffect(() => {
    const target = localStorage.getItem("navigationTarget");
    if (!target) return;
    localStorage.removeItem("navigationTarget");
    const parsedTarget = NavigationTargetSchema.safeParse(target);
    if (!parsedTarget.success) return;
    if (parsedTarget.data === "edit-selection") {
      setInputValue("Edit the current PowerPoint selection using the selected shapes and slide context.");
    } else if (parsedTarget.data === "insert-timeline") {
      setInputValue("Insert a project timeline on the current slide using the deck template when possible.");
    } else if (parsedTarget.data === "insert-estimate-table") {
      setInputValue("Insert an estimate summary table on the current slide using editable native PowerPoint content.");
    }
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
        const items = z.array(OfficePermissionRequestSchema).catch([]).parse(await response.json());
        const family = await getSessionFamily(client, currentSessionId);
        const next = items.find((item) => family.ids.has(item.sessionID));

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
    setCurrentActivity("");
    setDebugEvents([]);
    setIsTyping(false);
    setPermission(null);
    setPermissionSessionTitle(null);
    setLiveMessages([]);

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

    const status = await client.getStatus().catch(() => null);
    if (status?.mode) setRuntimeMode(status.mode);
    if (!safeSelectedModel && availableModels.length > 0) {
      setSelectedModel(pickDefaultModel(availableModels));
    }
  };

  const handleRestoreSession = async (saved: OpencodeSessionInfo) => {
    const restored = await restoreSession(saved.id, safeSelectedModel);
    void startNewSession(safeSelectedModel, restored);
  };

  const handleModelChange = (model: ModelType) => {
    setSelectedModel(model);
  };

  const ensureSession = async () => {
    if (currentSessionId) return currentSessionId;

    const office = getHostFromOfficeHost(Office.context.host);
    const session = await client.createSession({
      title: `${office === "powerpoint" ? "PowerPoint" : office === "excel" ? "Excel" : office === "onenote" ? "OneNote" : "Word"}: New chat`,
    });
    setCurrentSessionId(session.id);
    return session.id;
  };

  const handleQaSubagentModelChange = async (model: ModelType) => {
    const previous = qaSubagentModel;
    setQaSubagentModel(model);

    try {
      const config = await client.getConfig();
      await client.updateConfig(mergeQaModel(config, model));
    } catch {
      setQaSubagentModel(previous);
    }
  };

  const handleSend = async () => {
    if (isTyping || (!inputValue.trim() && images.length === 0)) return;

    const userInput = inputValue;
    const userImages = [...images];
    const isFirstUserTurn = !messages.some((message) => message.sender === "user");

    let sessionId = currentSessionId;

    try {
      sessionId = await ensureSession();
    } catch (err) {
      setError(`Failed to create session: ${err instanceof Error ? err.message : String(err)}`);
      return;
    }

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
    setLiveMessages([]);
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
        const upload = UploadImageResponseSchema.parse(result);
        uploads.push({ path: upload.path, name: image.name, mime: upload.mime || "image/png" });
      }

      const model = availableModels.find((item) => item.key === safeSelectedModel) || FALLBACK_MODELS[0];
      const parts = toPromptParts(userInput, uploads);
      const tools = getEnabledTools(officeHost);

      if (isFirstUserTurn && userInput.trim()) {
        updateSessionTitle(sessionId, makeSessionTitle(officeHost, userInput)).catch(() => undefined);
      }

      let count = 0;

      for await (const event of client.query(sessionId, {
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
          setLiveMessages((prev) => mergeLiveMessages(prev, {
            id: String(event.id || `assistant-${Date.now()}`),
            text: `${prev.find((item) => item.id === String(event.id || ""))?.text || ""}${String(data.deltaContent || "")}`,
            sender: "assistant",
            timestamp: new Date(),
          }));
          setCurrentActivity("");
          continue;
        }

        if (event.type === "assistant.message") {
          const parts = Array.isArray(data.parts) ? data.parts : [];
          const mapped = mapAssistantParts(parts, Date.now());
          const kept = carry(liveRef.current, mapped);
          setLiveMessages([]);
          setCurrentActivity("");
          if (mapped.length > 0) {
            setMessages((prev) => [...prev, ...kept, ...mapped]);
            continue;
          }
          const text = String(data.content || "");
          if (text || kept.length > 0) {
            setMessages((prev) => [...prev, ...kept, ...(!text ? [] : [{
              id: String(event.id || crypto.randomUUID()),
              text,
              sender: "assistant" as const,
              timestamp: new Date(event.timestamp || Date.now()),
            }])]);
          }
          continue;
        }

        if (event.type === "tool.execution_start") {
          const toolName = String(data.toolName || "tool");
          const toolArgs = (data.arguments || {}) as Record<string, unknown>;
          const safeToolArgs = redactSensitiveFields(toolArgs) as Record<string, unknown>;
          const toolMeta = redactSensitiveFields((data.metadata || {}) as Record<string, unknown>) as Record<string, unknown>;
          setCurrentActivity(describeToolActivity(toolName, toolArgs));
          setLiveMessages((prev) => mergeLiveMessages(prev, {
            id: String(event.id || crypto.randomUUID()),
            text: JSON.stringify(safeToolArgs, null, 2),
            sender: "tool",
            startedAt: new Date(),
            toolName,
            toolArgs: safeToolArgs,
            toolMeta,
            toolStatus: "running",
            timestamp: new Date(),
          }));
          continue;
        }

        if (event.type === "tool.execution_complete") {
          const toolId = String(event.id || crypto.randomUUID());

          if (data.error) {
            const text = String(data.error);
            setCurrentActivity("");
            setLiveMessages((prev) => {
              const tool = prev.find((item) => item.id === toolId);
              return mergeLiveMessages(prev, {
                id: toolId,
                text: tool?.text || "{}",
                sender: "tool",
                startedAt: tool?.startedAt,
                finishedAt: new Date(),
                toolName: String(data.toolName || tool?.toolName || "tool"),
                toolArgs: tool?.toolArgs || {},
                toolMeta: redactSensitiveFields((data.metadata || tool?.toolMeta || {}) as Record<string, unknown>) as Record<string, unknown>,
                toolResult: undefined,
                toolError: text,
                toolStatus: "error",
                timestamp: new Date(),
              });
            });
            continue;
          }
          setLiveMessages((prev) => {
            const tool = prev.find((item) => item.id === toolId);
            return mergeLiveMessages(prev, {
              id: toolId,
              text: tool?.text || "{}",
              sender: "tool",
              startedAt: tool?.startedAt,
              finishedAt: new Date(),
              toolName: String(data.toolName || tool?.toolName || "tool"),
              toolArgs: tool?.toolArgs || {},
              toolMeta: redactSensitiveFields((data.metadata || tool?.toolMeta || {}) as Record<string, unknown>) as Record<string, unknown>,
              toolResult: data.result,
              toolError: undefined,
              toolStatus: "completed",
              timestamp: new Date(),
            });
          });
          setCurrentActivity("Processing result...");
          continue;
        }

        if (event.type === "assistant.reasoning_delta") {
          setLiveMessages((prev) => mergeLiveMessages(prev, {
            id: String(event.id || `thinking-${Date.now()}`),
            text: `${prev.find((item) => item.id === String(event.id || ""))?.text || ""}${String(data.deltaContent || "")}`,
            sender: "thinking",
            timestamp: new Date(),
          }));
          setCurrentActivity("Thinking...");
          continue;
        }

        if (event.type === "assistant.turn_start") {
          setCurrentActivity((current) => current || "Starting response...");
          continue;
        }

        if (event.type === "assistant.turn_end") {
          setCurrentActivity("");
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
      setLiveMessages([]);
    }
  };

  const handleStop = async () => {
    if (!currentSessionId || !isTyping) return;
    setCurrentActivity("Stopping...");

    try {
      await client.abortSession(currentSessionId);
      run.current?.abort();
      setLiveMessages([]);
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
        <div className={styles.container} style={getSurfaceVars(isDarkMode)}>
          <div className={styles.shell}>
            <SessionHistory
              host={officeHost}
              shared={safeSharedHistory}
              onSharedChange={setSharedHistory}
              onSelectSession={handleRestoreSession}
              onClose={() => setShowHistory(false)}
            />
          </div>
        </div>
      </FluentProvider>
    );
  }

  const powerpointContextLabel = officeHost === "powerpoint"
    ? [
        pptContext?.activeSlideIndex !== undefined ? `Slide ${pptContext.activeSlideIndex + 1}` : "No active slide",
        pptContext?.selectedShapeIds.length ? `${pptContext.selectedShapeIds.length} shape${pptContext.selectedShapeIds.length === 1 ? "" : "s"} selected` : "No shapes selected",
      ].join(" • ")
    : undefined;

  const headerSubtitle = [
    runtimeMode ? `${hostLabel} • ${runtimeMode} • ${getToolNamesForHost(officeHost).length} tools` : `${hostLabel} • ${getToolNamesForHost(officeHost).length} tools`,
    powerpointContextLabel,
  ].filter(Boolean).join(" • ");

  return (
    <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
      <div className={styles.container} style={getSurfaceVars(isDarkMode)}>
        <div className={styles.shell}>
          <HeaderBar
            onNewChat={() => {
              setCurrentSessionId("");
              void startNewSession(safeSelectedModel);
            }}
            onShowHistory={() => setShowHistory(true)}
            selectedModel={safeSelectedModel}
            onModelChange={handleModelChange}
            models={availableModels}
            debugEnabled={safeDebugEnabled}
            onDebugChange={setDebugEnabled}
            showThinking={safeShowThinking}
            onShowThinkingChange={setShowThinking}
            showToolResponses={safeShowToolResponses}
            onShowToolResponsesChange={setShowToolResponses}
            qaSubagentModel={safeQaSubagentModel}
            onQaSubagentModelChange={handleQaSubagentModelChange}
            subtitle={headerSubtitle}
          />

          <MessageList
            messages={messages}
            liveMessages={liveMessages}
            isTyping={isTyping}
            isConnecting={false}
            currentActivity={currentActivity}
            debugEvents={safeDebugEnabled ? debugEvents : undefined}
            hostLabel={hostLabel}
            showThinking={safeShowThinking}
            showToolResponses={safeShowToolResponses}
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
