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
import { useLocalStorage } from "./useLocalStorage";
import { createOpencodeClient, ModelInfo, OpencodeConfig, SessionInfo } from "./lib/opencode-client";
import { createOfficeToolBridge, readPowerPointContextSnapshot } from "./lib/office-tool-bridge";
import { carry, makeSessionTitle, mapAssistantParts, restoreSession, updateSessionTitle, type OpencodeSessionInfo } from "./lib/opencode-session-history";
import { trafficStats, type UiEvent } from "./lib/opencode-events";
import { formatTokenUsage, sessionUsageSchema, type SessionUsage } from "./lib/opencode-usage";
import { buildHeaderSubtitle, buildPowerPointContextLabel, deriveConnectionIndicator } from "./lib/chat-shell";
import { defaultThemeId, getThemeCssVars, resolveThemeTokens, themeOptions, useThemeMode, type ThemePreference } from "./lib/ui-theme";
import { getOfficeHostLabel, normalizeOfficeHost } from "./lib/officeHost";
import { getOfficeToolExecutor, getToolNamesForHost } from "./tools";
import { setPowerPointContextSnapshot, type PowerPointContextSnapshot } from "./tools/powerpointContext";
import { canAutoApprove, type OfficePermissionRequest } from "../shared/office-permissions";
import { formatOfficeToolActivity } from "../shared/office-tool-registry";
import {
  SavedSession,
  OfficeHost,
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
    boxSizing: "border-box",
    background: "var(--background-base)",
    color: "var(--text-strong)",
    fontFamily: '"Inter", "Segoe UI", sans-serif',
  },
  shell: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    background: "var(--background-stronger)",
    overflow: "hidden",
  },
  error: {
    margin: "0 auto 12px",
    width: "min(100%, 760px)",
    padding: "12px 14px",
    borderRadius: "12px",
    background: "var(--oc-danger-bg)",
    border: "1px solid var(--oc-danger-border)",
    color: "var(--oc-danger-text)",
  },
});

function isDarkHex(value: string) {
  const normalized = value.replace("#", "");
  const expanded = normalized.length === 3
    ? normalized.split("").map((char) => `${char}${char}`).join("")
    : normalized;

  if (expanded.length !== 6) return false;

  const red = Number.parseInt(expanded.slice(0, 2), 16);
  const green = Number.parseInt(expanded.slice(2, 4), 16);
  const blue = Number.parseInt(expanded.slice(4, 6), 16);
  const luminance = (red * 299 + green * 587 + blue * 114) / 1000;
  return luminance < 140;
}

const FALLBACK_MODELS: ModelInfo[] = [
  {
    key: "anthropic/claude-sonnet-4-5",
    label: "Anthropic / Claude Sonnet 4.5",
    providerID: "anthropic",
    modelID: "claude-sonnet-4-5",
  },
];

/** Choose the best available model, preferring the newest Sonnet release. */
function pickDefaultModel(models: { key: string }[]): ModelType {
  const keys = new Set(models.map((m) => m.key));
  const match = ["anthropic/claude-sonnet-4-6", "anthropic/claude-sonnet-4-5"]
    .find((candidate) => keys.has(candidate));
  return match ?? models[0]?.key ?? FALLBACK_MODELS[0].key;
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

function mergeMessages(previous: Message[], next: Message[]) {
  return next.reduce((result, item) => {
    const index = result.findIndex((entry) => entry.id === item.id);
    if (index === -1) return [...result, item];
    return result.map((entry, current) => current === index ? { ...entry, ...item } : entry);
  }, previous);
}

function qaModel(config: OpencodeConfig) {
  const value = config.agent?.["visual-qa"]?.model;
  return typeof value === "string" ? value : "";
}

function qaVariant(config: OpencodeConfig) {
  const agent = config.agent?.["visual-qa"];
  if (!agent || typeof agent !== "object") return undefined;
  const value = (agent as Record<string, unknown>).variant;
  return typeof value === "string" ? value : undefined;
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

function mergeQaVariant(config: OpencodeConfig, variant: string | undefined) {
  const next = {
    ...config,
    agent: {
      ...(config.agent || {}),
      "visual-qa": {
        ...(config.agent?.["visual-qa"] || {}),
      },
    },
  } satisfies OpencodeConfig;

  if (variant) {
    (next.agent!["visual-qa"] as Record<string, unknown>).variant = variant;
    return next;
  }

  if (next.agent?.["visual-qa"]) {
    const visual = { ...next.agent["visual-qa"] } as Record<string, unknown>;
    delete visual.variant;
    next.agent!["visual-qa"] = visual as typeof next.agent["visual-qa"];
  }

  return next;
}

function previewEvent(eventType: string, data: Record<string, unknown>) {
  if (eventType === "assistant.message_delta") return String(data.deltaContent || "").slice(0, 80);
  if (eventType === "assistant.message_update") return String(data.content || "").slice(0, 80);
  if (eventType === "assistant.reasoning_delta") return String(data.deltaContent || "").slice(0, 80);
  if (eventType === "assistant.reasoning_update") return String(data.content || "").slice(0, 80);
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

type SessionStream = {
  sessionId: string;
  ready: Promise<void>;
  close: () => void;
};

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
  const [selectedThemeId, setSelectedThemeId] = useLocalStorage<string>("opencode-ui-theme", defaultThemeId);
  const [qaSubagentModel, setQaSubagentModel] = useState<ModelType>("");
  const [qaSubagentVariant, setQaSubagentVariant] = useState<string | undefined>(undefined);
  const [selectedVariant, setSelectedVariant] = useLocalStorage<string>("opencode-selected-variant", "");
  const [usage, setUsage] = useState<SessionUsage | null>(null);
  const [permission, setPermission] = useState<OfficePermissionRequest | null>(null);
  const [permissionSessionTitle, setPermissionSessionTitle] = useState<string | null>(null);
  const [liveMessages, setLiveMessages] = useState<Message[]>([]);
  const [pptContext, setPptContext] = useState<PowerPointContextSnapshot | null>(null);
  const [connectionState, setConnectionState] = useState({ isLoading: true, hasLoaded: false, hasFailed: false });
  const [themePreference, setThemePreference] = useLocalStorage<ThemePreference>("opencode-theme-mode", "system");
  const safeThemePreference: ThemePreference = (themePreference === "light" || themePreference === "dark") ? themePreference : "system";
  const themeMode = useThemeMode(safeThemePreference);
  const safeSelectedModel = PersistedModelSchema.catch("").parse(selectedModel);
  const safeDebugEnabled = PersistedBooleanSchema.catch(false).parse(debugEnabled);
  const safeShowThinking = PersistedBooleanSchema.catch(true).parse(showThinking);
  const safeShowToolResponses = PersistedBooleanSchema.catch(false).parse(showToolResponses);
  const safeSharedHistory = PersistedBooleanSchema.catch(false).parse(sharedHistory);
  const safeSelectedThemeId = themeOptions.some((theme) => theme.id === selectedThemeId) ? selectedThemeId : defaultThemeId;
  const safeQaSubagentModel = PersistedModelSchema.catch("").parse(qaSubagentModel);
  const hostLabel = getOfficeHostLabel(officeHost);
  const safeSelectedVariant = useMemo(() => {
    const raw = PersistedModelSchema.catch("").parse(selectedVariant) || undefined;
    if (!raw) return undefined;
    const model = availableModels.find((item) => item.key === safeSelectedModel);
    if (!model?.variants?.includes(raw)) return undefined;
    return raw;
  }, [selectedVariant, safeSelectedModel, availableModels]);
  const usageSummary = useMemo(() => formatTokenUsage(usage, availableModels), [availableModels, usage]);
  const enabledToolCount = useMemo(
    () => Object.values(getEnabledTools(officeHost)).filter(Boolean).length,
    [officeHost],
  );
  const resolvedTheme = useMemo(() => resolveThemeTokens(safeSelectedThemeId), [safeSelectedThemeId]);
  const fluentTheme = isDarkHex(resolvedTheme[themeMode].background) ? webDarkTheme : webLightTheme;
  const surfaceVars = useMemo(
    () => getThemeCssVars(safeSelectedThemeId, themeMode) as React.CSSProperties,
    [safeSelectedThemeId, themeMode],
  );
  const connectionStatus = useMemo(() => deriveConnectionIndicator(connectionState), [connectionState]);
  const sessionCreatedAt = useRef<string>("");
  const messagesRef = useRef<Message[]>([]);
  const liveRef = useRef<Message[]>([]);
  const streamTextRef = useRef<Map<string, string>>(new Map());
  const streamRef = useRef<SessionStream | null>(null);
  const eventHandlerRef = useRef<(event: UiEvent) => void>(() => undefined);

  useEffect(() => {
    if (safeSelectedThemeId !== selectedThemeId) {
      setSelectedThemeId(safeSelectedThemeId);
    }
  }, [safeSelectedThemeId, selectedThemeId, setSelectedThemeId]);

  useEffect(() => {
    messagesRef.current = messages;
  }, [messages]);

  useEffect(() => {
    liveRef.current = liveMessages;
  }, [liveMessages]);

  const upsertStreamMessage = (next: Message) => {
    if (messagesRef.current.some((item) => item.id === next.id)) {
      setMessages((prev) => mergeMessages(prev, [next]));
      return;
    }

    setLiveMessages((prev) => mergeLiveMessages(prev, next));
  };

  const streamText = (id: string, sender: "assistant" | "thinking", text: string, append: boolean) => {
    const current = streamTextRef.current.get(id)
      || [...messagesRef.current, ...liveRef.current].find((item) => item.id === id)?.text
      || "";
    const nextText = append ? `${current}${text}` : text;
    streamTextRef.current.set(id, nextText);
    upsertStreamMessage({
      id,
      text: nextText,
      sender,
      timestamp: new Date(),
    });
  };

  eventHandlerRef.current = (event: UiEvent) => {
    const data = event.data || {};
    const preview = previewEvent(event.type, data);

    setDebugEvents((prev) => [...prev, { type: event.type, preview, timestamp: Date.now() }]);
    setConnectionState({ isLoading: false, hasLoaded: true, hasFailed: false });

    if (event.type === "assistant.usage") {
      const nextUsage = sessionUsageSchema.nullish().catch(null).parse(data.usage);
      if (nextUsage) setUsage(nextUsage);
      return;
    }

    if (event.type === "assistant.message_delta") {
      setIsTyping(true);
      streamText(String(event.id || `assistant-${Date.now()}`), "assistant", String(data.deltaContent || ""), true);
      setCurrentActivity("");
      return;
    }

    if (event.type === "assistant.message_update") {
      setIsTyping(true);
      streamText(String(event.id || `assistant-${Date.now()}`), "assistant", String(data.content || ""), false);
      setCurrentActivity("");
      return;
    }

    if (event.type === "assistant.message") {
      const nextUsage = sessionUsageSchema.nullish().catch(null).parse(data.usage);
      if (nextUsage) setUsage(nextUsage);
      const parts = Array.isArray(data.parts) ? data.parts : [];
      const mapped = mapAssistantParts(parts, Date.now());
      const kept = carry(liveRef.current, mapped);
      streamTextRef.current.clear();
      setLiveMessages([]);
      setIsTyping(false);
      setCurrentActivity("");
      if (mapped.length > 0) {
        setMessages((prev) => mergeMessages(prev, [...kept, ...mapped]));
        return;
      }
      const text = String(data.content || "");
      if (text || kept.length > 0) {
        setMessages((prev) => mergeMessages(prev, [...kept, ...(!text ? [] : [{
          id: String(event.id || crypto.randomUUID()),
          text,
          sender: "assistant" as const,
          timestamp: new Date(event.timestamp || Date.now()),
        }])]));
      }
      return;
    }

    if (event.type === "tool.execution_start") {
      setIsTyping(true);
      const toolName = String(data.toolName || "tool");
      const toolArgs = (data.arguments || {}) as Record<string, unknown>;
      const safeToolArgs = redactSensitiveFields(toolArgs) as Record<string, unknown>;
      const toolMeta = redactSensitiveFields((data.metadata || {}) as Record<string, unknown>) as Record<string, unknown>;
      setCurrentActivity(describeToolActivity(toolName, toolArgs));
      upsertStreamMessage({
        id: String(event.id || crypto.randomUUID()),
        text: JSON.stringify(safeToolArgs, null, 2),
        sender: "tool",
        startedAt: new Date(),
        toolName,
        toolArgs: safeToolArgs,
        toolMeta,
        toolStatus: "running",
        timestamp: new Date(),
      });
      return;
    }

    if (event.type === "tool.execution_complete") {
      setIsTyping(true);
      const toolId = String(event.id || crypto.randomUUID());

      if (data.error) {
        const text = String(data.error);
        setCurrentActivity("");
        const tool = [...messagesRef.current, ...liveRef.current].find((item) => item.id === toolId);
        upsertStreamMessage({
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
        return;
      }

      const tool = [...messagesRef.current, ...liveRef.current].find((item) => item.id === toolId);
      upsertStreamMessage({
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
      setCurrentActivity("Processing result...");
      return;
    }

    if (event.type === "assistant.reasoning_delta") {
      setIsTyping(true);
      streamText(String(event.id || `thinking-${Date.now()}`), "thinking", String(data.deltaContent || ""), true);
      setCurrentActivity("Thinking...");
      return;
    }

    if (event.type === "assistant.reasoning_update") {
      setIsTyping(true);
      streamText(String(event.id || `thinking-${Date.now()}`), "thinking", String(data.content || ""), false);
      setCurrentActivity("Thinking...");
      return;
    }

    if (event.type === "assistant.turn_start") {
      setIsTyping(true);
      setCurrentActivity((current) => current || "Starting response...");
      return;
    }

    if (event.type === "assistant.turn_end") {
      setIsTyping(false);
      setCurrentActivity("");
      return;
    }

    if (event.type === "session.error") {
      const text = String(data.message || "Unknown session error");
      streamTextRef.current.clear();
      setIsTyping(false);
      setCurrentActivity("");
      setConnectionState({ isLoading: false, hasLoaded: false, hasFailed: true });
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `⚠️ Session error: ${text}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
    }
  };

  const ensureSessionStream = async (sessionId: string) => {
    if (streamRef.current?.sessionId === sessionId) {
      await streamRef.current.ready;
      return;
    }

    streamRef.current?.close();

    const controller = new AbortController();
    const subscription = client.subscribe(sessionId, {
      onEvent: (event) => {
        eventHandlerRef.current(event);
      },
    }, { signal: controller.signal });
    const next = {
      sessionId,
      ready: subscription.ready,
      close: () => {
        controller.abort();
        subscription.close();
      },
    } satisfies SessionStream;

    streamRef.current = next;
    await next.ready;
  };

  const fetchModels = async () => {
    setConnectionState((current) => ({ ...current, isLoading: true, hasFailed: false }));

    try {
      const status = await client.getStatus();
      setConnectionState({ isLoading: false, hasLoaded: true, hasFailed: false });
      const models = status.models?.length ? status.models : FALLBACK_MODELS;
      setAvailableModels(models);
      if (!safeSelectedModel || !models.some((model) => model.key === safeSelectedModel)) {
        setSelectedModel(pickDefaultModel(models));
      }
    } catch {
      setConnectionState({ isLoading: false, hasLoaded: false, hasFailed: true });
      setAvailableModels(FALLBACK_MODELS);
      if (!safeSelectedModel || !FALLBACK_MODELS.some((model) => model.key === safeSelectedModel)) {
        setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
      }
    }

    try {
      const config = await client.getConfig();
      setQaSubagentModel(qaModel(config));
      setQaSubagentVariant(qaVariant(config));
    } catch {
      setQaSubagentModel("");
      setQaSubagentVariant(undefined);
    }
  };

  useEffect(() => {
    void fetchModels();
  }, []);

  useEffect(() => {
    const host = normalizeOfficeHost(Office.context.host);
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

  useEffect(() => {
    if (!currentSessionId) {
      streamRef.current?.close();
      streamRef.current = null;
      return;
    }

    void ensureSessionStream(currentSessionId).catch((err) => {
      const text = err instanceof Error ? err.message : String(err);
      setConnectionState({ isLoading: false, hasLoaded: false, hasFailed: true });
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `❌ Error: ${text}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
    });

    return () => {
      if (streamRef.current?.sessionId !== currentSessionId) return;
      streamRef.current.close();
      streamRef.current = null;
    };
  }, [currentSessionId]);

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
    setUsage(restored?.usage || null);
    setPermission(null);
    setPermissionSessionTitle(null);
    setLiveMessages([]);
    streamTextRef.current.clear();

    const host = Office.context.host;
    const office = normalizeOfficeHost(host);
    setOfficeHost(office);

    setCurrentSessionId("");
    sessionCreatedAt.current = restored?.createdAt || new Date().toISOString();
    setMessages(restored?.messages || []);

    if (restored) {
      setCurrentSessionId(restored.id);
      return;
    }

    await client.getStatus().catch(() => null);
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

  const handleVariantChange = (variant: string | undefined) => {
    setSelectedVariant(variant ?? "");
  };

  const ensureSession = async () => {
    if (currentSessionId) return currentSessionId;

    const office = normalizeOfficeHost(Office.context.host);
    const session = await client.createSession({
      title: `${getOfficeHostLabel(office)}: New chat`,
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

  const handleQaSubagentVariantChange = async (variant: string | undefined) => {
    const previous = qaSubagentVariant;
    setQaSubagentVariant(variant);

    try {
      const config = await client.getConfig();
      await client.updateConfig(mergeQaVariant(config, variant));
    } catch {
      setQaSubagentVariant(previous);
    }
  };

  const handleSend = async () => {
    if (!inputValue.trim() && images.length === 0) return;

    const userInput = inputValue;
    const userImages = [...images];
    const userMessage: Message = {
      id: crypto.randomUUID(),
      text: userInput || `Sent ${userImages.length} image${userImages.length === 1 ? "" : "s"}`,
      sender: "user",
      timestamp: new Date(),
      images: userImages.length > 0 ? userImages.map((image) => ({ dataUrl: image.dataUrl, name: image.name })) : undefined,
    };
    const isFirstUserTurn = !messages.some((message) => message.sender === "user");
    const wasTyping = isTyping;
    const carried = wasTyping ? liveRef.current : [];

    let sessionId = currentSessionId;

    try {
      sessionId = await ensureSession();
    } catch (err) {
      setError(`Failed to create session: ${err instanceof Error ? err.message : String(err)}`);
      return;
    }

    setMessages((prev) => {
      const next = mergeMessages(prev, carried);
      const merged = [...next, userMessage];
      messagesRef.current = merged;
      return merged;
    });

    setInputValue("");
    setImages([]);
    setIsTyping(true);
    setCurrentActivity("Processing...");
    setError("");

    if (carried.length > 0) {
      liveRef.current = [];
      setLiveMessages([]);
    }

    if (!wasTyping) {
      liveRef.current = [];
      setLiveMessages([]);
      setDebugEvents([]);
      streamTextRef.current.clear();
      trafficStats.reset();
    }

    try {
      await ensureSessionStream(sessionId);

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

      await client.sendMessage(sessionId, {
        model: { providerID: model.providerID, modelID: model.modelID },
        agent: officeHost,
        system: getSystemMessage(Office.context.host),
        parts,
        tools,
        variant: safeSelectedVariant,
      });
    } catch (err) {
      setConnectionState({ isLoading: false, hasLoaded: false, hasFailed: true });
      const text = err instanceof Error ? err.message : String(err);
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `❌ Error: ${text}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
      if (!wasTyping) {
        setIsTyping(false);
        setCurrentActivity("");
      }
    }
  };

  const handleStop = async () => {
    if (!currentSessionId || !isTyping) return;
    setCurrentActivity("Stopping...");

    try {
      await client.abortSession(currentSessionId);
      streamTextRef.current.clear();
      setMessages((prev) => mergeMessages(prev, liveRef.current));
      setLiveMessages([]);
      setIsTyping(false);
      setCurrentActivity("");
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
      <FluentProvider theme={fluentTheme}>
        <div className={styles.container} style={surfaceVars}>
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

  const headerSubtitle = buildHeaderSubtitle({
    host: officeHost,
    enabledToolCount,
  });
  const headerContext = officeHost === "powerpoint" ? buildPowerPointContextLabel(pptContext) : undefined;

  return (
    <FluentProvider theme={fluentTheme}>
      <div className={styles.container} style={surfaceVars}>
        <div className={styles.shell}>
          <HeaderBar
            onNewChat={() => {
              setCurrentSessionId("");
              void startNewSession(safeSelectedModel);
            }}
            onShowHistory={() => setShowHistory(true)}
            selectedModel={safeSelectedModel}
            models={availableModels}
            debugEnabled={safeDebugEnabled}
            onDebugChange={setDebugEnabled}
            showThinking={safeShowThinking}
            onShowThinkingChange={setShowThinking}
            showToolResponses={safeShowToolResponses}
            onShowToolResponsesChange={setShowToolResponses}
            qaSubagentModel={safeQaSubagentModel}
            onQaSubagentModelChange={handleQaSubagentModelChange}
            qaSubagentVariant={qaSubagentVariant}
            onQaSubagentVariantChange={handleQaSubagentVariantChange}
            themes={themeOptions}
            selectedThemeId={safeSelectedThemeId}
            onThemeChange={setSelectedThemeId}
            themePreference={safeThemePreference}
            onThemePreferenceChange={setThemePreference}
            connectionStatus={connectionStatus}
            subtitle={headerSubtitle}
            contextLabel={headerContext}
            usageSummary={usageSummary || undefined}
          />

          <MessageList
            messages={messages}
            liveMessages={liveMessages}
            isTyping={isTyping}
            isConnecting={connectionStatus.state === "connecting"}
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
            selectedModel={safeSelectedModel}
            onModelChange={handleModelChange}
            models={availableModels}
            selectedVariant={safeSelectedVariant}
            onVariantChange={handleVariantChange}
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
