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
import { useIsDarkMode } from "./useIsDarkMode";
import { useLocalStorage } from "./useLocalStorage";
import { createOpencodeClient, ModelInfo } from "./lib/opencode-client";
import { createOfficeToolBridge } from "./lib/office-tool-bridge";
import { trafficStats } from "./lib/opencode-events";
import { getOfficeToolExecutor, getToolNamesForHost } from "./tools";
import {
  SavedSession,
  OfficeHost,
  saveSession,
  generateSessionTitle,
  getHostFromOfficeHost,
} from "./sessionStorage";
import React from "react";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: "var(--colorNeutralBackground3)",
  },
});

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

function getSystemMessage(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const hostName = host === Office.HostType.PowerPoint
    ? "PowerPoint"
    : host === Office.HostType.Word
      ? "Word"
      : host === Office.HostType.Excel
        ? "Excel"
        : "Office";

  return `You are a helpful AI assistant embedded inside Microsoft ${hostName} as an Office Add-in. The user's ${hostName} document is already open.

Use the available ${hostName} tools to inspect or update the active document directly. Do not ask for file paths, exports, or saved files on disk.

${host === Office.HostType.PowerPoint ? `For PowerPoint:
- Use get_presentation_overview first to understand the deck
- Use get_presentation_content to inspect slide text
- Use get_slide_image when visual layout matters` : ""}

${host === Office.HostType.Word ? `For Word:
- Use get_document_overview first to map the document structure
- Use get_document_content to read the document
- Use get_document_section or selection tools for targeted edits
- Use mutation tools directly against the active document instead of asking the user to paste content` : ""}

${host === Office.HostType.Excel ? `For Excel:
- Use get_workbook_info to understand workbook structure
- Use get_workbook_content to inspect sheet data before making changes` : ""}

Always operate on the open document through tools.`;
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
  const [runtimeMode, setRuntimeMode] = useState<string>("");
  const isDarkMode = useIsDarkMode();
  const sessionCreatedAt = useRef<string>("");
  const started = useRef(false);

  const fetchModels = async () => {
    try {
      const status = await client.getStatus();
      setRuntimeMode(status.mode);
      const models = status.models?.length ? status.models : FALLBACK_MODELS;
      setAvailableModels(models);
      if (!selectedModel) {
        setSelectedModel(pickDefaultModel(models));
      }
    } catch {
      setAvailableModels(FALLBACK_MODELS);
      if (!selectedModel) {
        setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
      }
    }
  };

  useEffect(() => {
    void fetchModels();
  }, []);

  useEffect(() => {
    if (messages.length === 0 || !currentSessionId) return;
    if (!messages.some((message) => message.sender === "user")) return;

    const saved: SavedSession = {
      id: currentSessionId,
      title: generateSessionTitle(messages),
      model: selectedModel,
      messages,
      createdAt: sessionCreatedAt.current,
      updatedAt: new Date().toISOString(),
    };

    saveSession(officeHost, saved);
  }, [messages, currentSessionId, selectedModel, officeHost]);

  useEffect(() => {
    const host = getHostFromOfficeHost(Office.context.host);
    setOfficeHost(host);
    const bridge = createOfficeToolBridge(host, getOfficeToolExecutor(Office.context.host));
    return () => {
      void bridge.stop();
    };
  }, []);

  const startNewSession = async (_modelKey: ModelType, restored?: SavedSession) => {
    setError("");
    setShowHistory(false);
    setStreamingText("");
    setCurrentActivity("");
    setDebugEvents([]);
    setIsTyping(false);

    const host = Office.context.host;
    const office = getHostFromOfficeHost(host);
    setOfficeHost(office);

    setCurrentSessionId("");
    sessionCreatedAt.current = restored?.createdAt || new Date().toISOString();
    setMessages(restored?.messages || []);

    try {
      const session = await client.createSession();
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

  const handleRestoreSession = (saved: SavedSession) => {
    setSelectedModel(saved.model);
    void startNewSession(saved.model, saved);
  };

  useEffect(() => {
    if (!selectedModel || started.current) return;
    started.current = true;
    void startNewSession(selectedModel);
  }, [selectedModel]);

  const handleModelChange = (model: ModelType) => {
    setSelectedModel(model);
    setCurrentSessionId("");
    void startNewSession(model);
  };

  const handleSend = async () => {
    if ((!inputValue.trim() && images.length === 0) || !currentSessionId) return;

    const userInput = inputValue;
    const userImages = [...images];

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
      const tools = Object.fromEntries(
        getToolNamesForHost(officeHost).map((name) => [name, true]),
      );
      let count = 0;

      for await (const event of client.query(currentSessionId, {
        model: { providerID: model.providerID, modelID: model.modelID },
        system: getSystemMessage(Office.context.host),
        parts,
        tools,
      })) {
        count += 1;
        const data = event.data || {};
        const preview = event.type === "assistant.message_delta"
          ? String(data.deltaContent || "").slice(0, 80)
          : event.type === "assistant.message"
            ? String(data.content || "").slice(0, 80)
            : event.type === "tool.execution_start"
              ? String(data.toolName || "")
              : event.type === "session.error"
                ? String(data.message || "")
                : JSON.stringify(data).slice(0, 100);

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
          setCurrentActivity(`Calling ${toolName}...`);
          setMessages((prev) => [...prev, {
            id: String(event.id || crypto.randomUUID()),
            text: JSON.stringify(toolArgs, null, 2),
            sender: "tool",
            toolName,
            toolArgs,
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

      if (count === 0) {
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
      setIsTyping(false);
    }
  };

  if (showHistory) {
    return (
      <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
        <SessionHistory
          host={officeHost}
          onSelectSession={handleRestoreSession}
          onClose={() => setShowHistory(false)}
        />
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
      <div className={styles.container}>
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
          subtitle={runtimeMode ? `OpenCode ${runtimeMode} • ${getToolNamesForHost(officeHost).length} Office tools` : undefined}
        />

        <MessageList
          messages={messages}
          isTyping={isTyping}
          isConnecting={!currentSessionId && !error}
          currentActivity={currentActivity}
          streamingText={streamingText}
          debugEvents={debugEnabled ? debugEvents : undefined}
        />

        {error && <div style={{ color: "red", padding: "8px" }}>{error}</div>}

        <ChatInput
          value={inputValue}
          onChange={setInputValue}
          onSend={handleSend}
          images={images}
          onImagesChange={setImages}
        />
      </div>
    </FluentProvider>
  );
};
