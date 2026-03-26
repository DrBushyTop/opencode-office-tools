import * as React from "react";
import { useRef, useEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { trafficStats } from "../lib/opencode-events";

export interface Message {
  id: string;
  text: string;
  sender: "user" | "assistant" | "tool";
  timestamp: Date;
  toolName?: string;
  toolArgs?: Record<string, unknown>;
  images?: Array<{ dataUrl: string; name: string }>;
}

export interface DebugEvent {
  type: string;
  preview: string;
  timestamp: number;
}

function summarizeTaskTool(args: Record<string, unknown>) {
  const subagentType = typeof args.subagent_type === "string" ? args.subagent_type : "subagent";
  const description = typeof args.description === "string" ? args.description : "Working";
  const prefix = /fresh[ -]?eyes|verif|review|qa/i.test(description) ? "Verification" : "Subagent";
  return `${prefix}: ${subagentType}: ${description}`;
}

interface MessageListProps {
  messages: Message[];
  isTyping: boolean;
  isConnecting?: boolean;
  currentActivity?: string;
  streamingText?: string;
  debugEvents?: DebugEvent[];
  hostLabel?: string;
}

// Tool display configuration
const toolConfig: Record<string, { icon: string; format: (args: Record<string, unknown>) => string }> = {
  get_slide_image: {
    icon: "📸",
    format: (args) => `Capturing slide ${(args.slideIndex as number) + 1}`,
  },
  get_presentation_overview: {
    icon: "📊",
    format: () => "Getting presentation overview",
  },
  get_presentation_content: {
    icon: "📄",
    format: (args) => {
      if (args.slideIndex !== undefined) return `Reading slide ${(args.slideIndex as number) + 1}`;
      if (args.startIndex !== undefined && args.endIndex !== undefined) 
        return `Reading slides ${(args.startIndex as number) + 1}–${(args.endIndex as number) + 1}`;
      return "Reading all slides";
    },
  },
  set_presentation_content: {
    icon: "✏️",
    format: (args) => `Adding content to slide ${(args.slideIndex as number) + 1}`,
  },
  clear_slide: {
    icon: "🗑️",
    format: (args) => `Clearing slide ${(args.slideIndex as number) + 1}`,
  },
  add_slide_from_code: {
    icon: "➕",
    format: () => "Creating new slide",
  },
  update_slide_shape: {
    icon: "✏️",
    format: (args) => `Updating shape on slide ${(args.slideIndex as number) + 1}`,
  },
  get_document_content: {
    icon: "📄",
    format: () => "Reading document",
  },
  get_document_part: {
    icon: "🧩",
    format: (args) => `Reading ${String(args.address || "document part")}`,
  },
  get_document_overview: {
    icon: "🧭",
    format: () => "Mapping document structure",
  },
  get_document_section: {
    icon: "📑",
    format: (args) => `Reading section ${JSON.stringify(args.headingText || "")}`,
  },
  get_selection: {
    icon: "✂️",
    format: () => "Reading selection markup",
  },
  get_selection_text: {
    icon: "✂️",
    format: () => "Reading selected text",
  },
  set_document_content: {
    icon: "✏️",
    format: () => "Updating document",
  },
  set_document_part: {
    icon: "🧩",
    format: (args) => `Updating ${String(args.address || "document part")}`,
  },
  insert_content_at_selection: {
    icon: "➕",
    format: (args) => `Inserting content at ${String(args.location || "replace")}`,
  },
  find_and_replace: {
    icon: "🔁",
    format: (args) => `Replacing ${JSON.stringify(args.find || "")}`,
  },
  insert_table: {
    icon: "▦",
    format: () => "Inserting table",
  },
  apply_style_to_selection: {
    icon: "🎨",
    format: () => "Styling selection",
  },
  get_workbook_info: {
    icon: "📊",
    format: () => "Getting workbook structure",
  },
  get_workbook_overview: {
    icon: "🧭",
    format: () => "Mapping workbook structure",
  },
  get_workbook_content: {
    icon: "📄",
    format: (args) => args.sheetName ? `Reading "${args.sheetName}"` : "Reading worksheet",
  },
  set_workbook_content: {
    icon: "✏️",
    format: (args) => args.sheetName ? `Updating "${args.sheetName}"` : "Updating worksheet",
  },
  get_selected_range: {
    icon: "✂️",
    format: () => "Reading selected range",
  },
  set_selected_range: {
    icon: "✏️",
    format: () => "Updating selected range",
  },
  find_and_replace_cells: {
    icon: "🔁",
    format: (args) => `Replacing ${JSON.stringify(args.find || "")}`,
  },
  insert_chart: {
    icon: "📈",
    format: () => "Creating chart",
  },
  apply_cell_formatting: {
    icon: "🎨",
    format: () => "Formatting cells",
  },
  create_named_range: {
    icon: "🏷️",
    format: (args) => `Naming range ${JSON.stringify(args.name || "")}`,
  },
  get_slide_notes: {
    icon: "📝",
    format: () => "Reading slide notes",
  },
  set_slide_notes: {
    icon: "📝",
    format: () => "Updating slide notes",
  },
  duplicate_slide: {
    icon: "📑",
    format: () => "Duplicating slide",
  },
  web_fetch: {
    icon: "🌐",
    format: (args) => {
      try {
        const url = new URL(args.url as string);
        return `Fetching ${url.hostname}`;
      } catch {
        return "Fetching web content";
      }
    },
  },
  report_intent: {
    icon: "💭",
    format: (args) => args.intent as string || "Working",
  },
  task: {
    icon: "🧠",
    format: (args) => summarizeTaskTool(args),
  },
};

function formatToolCall(toolName: string, args: Record<string, unknown>): { icon: string; description: string } {
  const config = toolConfig[toolName];
  if (config) {
    return { icon: config.icon, description: config.format(args) };
  }
  // Fallback for unknown tools
  return { icon: "🔧", description: toolName.replace(/_/g, " ") };
}

const useStyles = makeStyles({
  chatContainer: {
    flex: 1,
    overflowY: "scroll",
    padding: "20px 14px 14px",
    display: "flex",
    flexDirection: "column",
    gap: "24px",
    scrollbarColor: "var(--oc-text-faint) transparent",
    scrollbarWidth: "thin",
  },
  content: {
    width: "100%",
    maxWidth: "760px",
    margin: "0 auto",
    display: "flex",
    flexDirection: "column",
    gap: "24px",
    minHeight: "100%",
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    minHeight: "100%",
    textAlign: "center",
    color: "var(--oc-text-faint)",
    gap: "8px",
  },
  emptyTitle: {
    fontSize: "28px",
    lineHeight: "1.2",
    color: "var(--oc-text)",
    fontWeight: "500",
  },
  emptyMeta: {
    fontSize: "13px",
    color: "var(--oc-text-muted)",
  },
  assistantIcon: {
    width: "24px",
    height: "24px",
    borderRadius: "6px",
  },
  messageUser: {
    alignSelf: "flex-end",
    backgroundColor: "var(--oc-accent-bg)",
    color: "var(--oc-text)",
    padding: "10px 14px",
    borderRadius: "12px",
    maxWidth: "70%",
    wordWrap: "break-word",
    border: "1px solid rgba(3, 76, 255, 0.10)",
  },
  messageAssistant: {
    alignSelf: "flex-start",
    maxWidth: "100%",
    wordWrap: "break-word",
    display: "grid",
    gridTemplateColumns: "24px 1fr",
    gap: "10px",
    alignItems: "start",
    justifyItems: "center",
    color: "var(--oc-text)",
    "& p:first-child": {
      marginTop: 0,
    },
    "& p:last-child": {
      marginBottom: 0,
    },
  },
  assistantBody: {
    width: "100%",
    minWidth: 0,
    lineHeight: "1.6",
    color: "var(--oc-text)",
    "& pre": {
      background: "var(--oc-bg-strong)",
      border: "1px solid var(--oc-border)",
      borderRadius: "10px",
      padding: "10px 12px",
      overflowX: "auto",
    },
    "& code": {
      background: "var(--oc-bg-soft)",
      borderRadius: "6px",
      padding: "1px 4px",
    },
  },
  messageTool: {
    alignSelf: "flex-start",
    fontSize: "12px",
    color: "var(--oc-text-muted)",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    padding: "6px 10px",
    borderRadius: "999px",
    backgroundColor: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    transition: "background-color 0.15s",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
  toolIcon: {
    fontSize: "14px",
  },
  toolArgs: {
    fontSize: "11px",
    fontFamily: "monospace",
    whiteSpace: "pre-wrap",
    marginTop: "4px",
    color: "var(--oc-text-faint)",
  },
  attachmentContainer: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    marginTop: "8px",
  },
  attachmentThumbnail: {
    width: "120px",
    height: "120px",
    borderRadius: "8px",
    objectFit: "cover",
    border: "2px solid rgba(255, 255, 255, 0.3)",
  },
  attachmentBadge: {
    fontSize: "11px",
    opacity: "0.8",
    marginTop: "4px",
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  streamingIndicator: {
    color: "var(--oc-text-muted)",
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
});

// Animated dots component for streaming indicator
const StreamingDots: React.FC = () => {
  return (
    <>
      <style>
        {`
          @keyframes pulse-dot {
            0%, 100% { opacity: 0.3; }
            50% { opacity: 1; }
          }
          .streaming-dot {
            width: 4px;
            height: 4px;
            border-radius: 50%;
            background-color: var(--colorNeutralForeground3, #666);
            animation: pulse-dot 1.4s ease-in-out infinite;
          }
          @keyframes progress-slide {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(400%); }
          }
          .activity-progress-bar {
            height: 2px;
            width: 100%;
            border-radius: 1px;
            background: var(--colorNeutralBackground3, #e0e0e0);
            overflow: hidden;
            margin-top: 6px;
          }
          .activity-progress-fill {
            height: 100%;
            width: 25%;
            border-radius: 1px;
            background: var(--colorBrandBackground, #0078d4);
            animation: progress-slide 1.5s ease-in-out infinite;
          }
        `}
      </style>
      <span style={{ display: 'inline-flex', gap: '3px', marginLeft: '2px' }}>
        <span className="streaming-dot" style={{ animationDelay: '0s' }} />
        <span className="streaming-dot" style={{ animationDelay: '0.2s' }} />
        <span className="streaming-dot" style={{ animationDelay: '0.4s' }} />
      </span>
    </>
  );
};

// Elapsed time counter
const ElapsedTime: React.FC = () => {
  const [elapsed, setElapsed] = useState(0);
  const startRef = useRef(Date.now());

  useEffect(() => {
    startRef.current = Date.now();
    setElapsed(0);
    const interval = setInterval(() => {
      setElapsed(Math.floor((Date.now() - startRef.current) / 1000));
    }, 1000);
    return () => clearInterval(interval);
  }, []);

  if (elapsed < 3) return null;
  return (
    <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3, #999)', marginLeft: '6px' }}>
      {elapsed}s
    </span>
  );
};

function formatBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(1)} MB`;
}

// Live traffic counter that polls trafficStats (reset by App before each query)
const TrafficCounter: React.FC = () => {
  const [stats, setStats] = useState({ bytesIn: 0, bytesOut: 0 });
  const prevInRef = useRef(0);
  const [flash, setFlash] = useState(false);

  useEffect(() => {
    const interval = setInterval(() => {
      setStats({ bytesIn: trafficStats.bytesIn, bytesOut: trafficStats.bytesOut });
      if (trafficStats.bytesIn !== prevInRef.current) {
        prevInRef.current = trafficStats.bytesIn;
        setFlash(true);
        setTimeout(() => setFlash(false), 200);
      }
    }, 250);
    return () => clearInterval(interval);
  }, []);

  return (
    <span style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: '6px',
      fontSize: '10px',
      fontFamily: 'monospace',
      color: 'var(--colorNeutralForeground3, #999)',
      marginLeft: '8px',
      transition: 'color 0.2s',
    }}>
      <span style={{ color: flash ? 'var(--colorBrandBackground, #0078d4)' : undefined, transition: 'color 0.2s' }}>
        ↓{formatBytes(stats.bytesIn)}
      </span>
      <span>↑{formatBytes(stats.bytesOut)}</span>
    </span>
  );
};

export const MessageList: React.FC<MessageListProps> = ({
  messages,
  isTyping,
  isConnecting,
  currentActivity,
  streamingText,
  debugEvents,
  hostLabel,
}) => {
  const styles = useStyles();
  const chatEndRef = useRef<HTMLDivElement>(null);
  const [expandedTools, setExpandedTools] = useState<Set<string>>(new Set());

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, streamingText]);

  const toggleTool = (id: string) => {
    setExpandedTools(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  return (
    <div className={styles.chatContainer}>
      <div className={styles.content}>
        {messages.length === 0 && !isConnecting && (
          <div className={styles.emptyState}>
            <div className={styles.emptyTitle}>What can I do for you?</div>
            <div className={styles.emptyMeta}>{hostLabel ? `${hostLabel} workspace` : "Open document workspace"}</div>
          </div>
        )}

        {isConnecting && (
          <div className={styles.emptyState}>
            <div className={styles.emptyTitle}>Connecting...</div>
          </div>
        )}

        {messages.map((message) => {
        // Format tool calls nicely
        const toolDisplay = message.toolName 
          ? formatToolCall(message.toolName, message.toolArgs || {})
          : null;
        
        return (
        <div
          key={message.id}
          className={
            message.sender === "user" ? styles.messageUser : 
            message.sender === "tool" ? styles.messageTool :
            styles.messageAssistant
          }
          onClick={message.toolName ? () => toggleTool(message.id) : undefined}
          title={message.toolName ? "Click to show details" : undefined}
        >
          {toolDisplay ? (
            <>
              <span className={styles.toolIcon}>{toolDisplay.icon}</span>
              <span>{toolDisplay.description}</span>
              {expandedTools.has(message.id) && (
                <div className={styles.toolArgs}>{message.text}</div>
              )}
            </>
          ) : message.sender === "assistant" ? (
            <>
              <img src="/icon-32.png" alt="" className={styles.assistantIcon} />
              <div className={styles.assistantBody}><Markdown remarkPlugins={[remarkGfm]}>{message.text}</Markdown></div>
            </>
          ) : (
            <>
              {message.text}
              {message.images && message.images.length > 0 && (
                <div className={styles.attachmentContainer}>
                  {message.images.map((img, idx) => (
                    <div key={idx}>
                      <img src={img.dataUrl} alt={img.name} className={styles.attachmentThumbnail} />
                      <div className={styles.attachmentBadge}>📎 {img.name}</div>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      );
        })}

        {isTyping && (
          <div className={styles.messageAssistant}>
            <img src="/icon-32.png" alt="" className={styles.assistantIcon} />
            <div className={styles.assistantBody}>
              {streamingText ? (
                <Markdown remarkPlugins={[remarkGfm]}>{streamingText}</Markdown>
              ) : (
                <>
                  <span className={styles.streamingIndicator}>
                    {currentActivity || "Thinking"}
                    <StreamingDots />
                    <ElapsedTime />
                  </span>
                  <div className="activity-progress-bar"><div className="activity-progress-fill" /></div>
                </>
              )}
              <TrafficCounter />
              {debugEvents && debugEvents.length > 0 && (
                <div style={{
                  marginTop: '8px',
                  maxHeight: '120px',
                  overflowY: 'auto',
                  fontSize: '10px',
                  fontFamily: 'monospace',
                  lineHeight: '1.6',
                  color: 'var(--oc-text-faint, #999)',
                  backgroundColor: 'var(--oc-bg-soft, #f5f5f5)',
                  borderRadius: '8px',
                  padding: '6px 8px',
                  border: '1px solid var(--oc-border, #e5e5e5)',
                }}>
                  {debugEvents.map((ev, i) => (
                    <div key={i} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      <span style={{ color: 'var(--oc-accent, #0078d4)' }}>{ev.type}</span>
                      {ev.preview && <span style={{ opacity: 0.7 }}> {ev.preview}</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        <div ref={chatEndRef} />
      </div>
    </div>
  );
};
