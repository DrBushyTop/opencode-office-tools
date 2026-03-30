import * as React from "react";
import { useRef, useEffect, useLayoutEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { z } from "zod";
import { trafficStats } from "../lib/opencode-events";
import { getOfficeToolUi } from "../../shared/office-tool-registry";

const RecordValueSchema: z.ZodType<Record<string, unknown>> = z.record(z.string(), z.unknown());
const MessageImageSchema = z.object({
  dataUrl: z.string().min(1),
  name: z.string().min(1),
});
const MessageSchema = z.object({
  id: z.string().min(1),
  text: z.string(),
  sender: z.enum(["user", "assistant", "tool", "thinking"]),
  timestamp: z.date(),
  startedAt: z.date().optional(),
  finishedAt: z.date().optional(),
  toolName: z.string().optional(),
  toolArgs: RecordValueSchema.optional(),
  toolResult: z.unknown().optional(),
  toolError: z.string().optional(),
  toolMeta: RecordValueSchema.optional(),
  toolStatus: z.enum(["running", "completed", "error"]).optional(),
  images: z.array(MessageImageSchema).optional(),
});
const DebugEventSchema = z.object({
  type: z.string().min(1),
  preview: z.string(),
  timestamp: z.number(),
});
const TaskToolArgsSchema = z.object({
  subagent_type: z.string().optional(),
  description: z.string().optional(),
  task_id: z.string().optional(),
}).catchall(z.unknown());
const SessionMessagePartSchema = z.object({
  type: z.string().optional(),
  tool: z.string().optional(),
  state: z.object({
    status: z.string().optional(),
    input: RecordValueSchema.optional(),
    time: z.object({
      start: z.number().optional(),
      end: z.number().optional(),
    }).partial().optional(),
  }).passthrough().optional(),
}).passthrough();
const SessionMessageItemSchema = z.object({
  info: z.object({
    role: z.string().optional(),
  }).optional(),
  parts: z.array(SessionMessagePartSchema).optional(),
}).passthrough();

export type Message = z.infer<typeof MessageSchema>;
export type DebugEvent = z.infer<typeof DebugEventSchema>;

function summarizeTaskTool(args: Record<string, unknown>) {
  const parsed = TaskToolArgsSchema.safeParse(args);
  const subagentType = parsed.success && parsed.data.subagent_type ? parsed.data.subagent_type : "subagent";
  const description = parsed.success && parsed.data.description ? parsed.data.description : "Working";
  const prefix = /fresh[ -]?eyes|verif|review|qa/i.test(description) ? "Verification" : "Subagent";
  return `${prefix}: ${subagentType}: ${description}`;
}

function taskSessionId(message: Message) {
  const meta = message.toolMeta?.sessionId;
  if (typeof meta === "string" && meta) return meta;

  const input = message.toolArgs?.task_id;
  if (typeof input === "string" && input) return input;

  if (typeof message.toolResult !== "string") return "";
  const match = message.toolResult.match(/task_id:\s+([^\s]+)/);
  return match?.[1] || "";
}

function countTools(items: z.infer<typeof SessionMessageItemSchema>[]) {
  return items.reduce((sum, item) => {
    if (item.info?.role !== "assistant" || !Array.isArray(item.parts)) return sum;
    return sum + item.parts.filter((part) => part.type === "tool").length;
  }, 0);
}

function lastDoneTool(items: z.infer<typeof SessionMessageItemSchema>[]) {
  const parts = items
    .flatMap((item) => item.info?.role === "assistant" && Array.isArray(item.parts) ? item.parts : [])
    .filter((part) => part.type === "tool" && !!part.tool)
    .filter((part) => part.state?.status === "completed" || part.state?.status === "error")
    .sort((a, b) => (a.state?.time?.end || a.state?.time?.start || 0) - (b.state?.time?.end || b.state?.time?.start || 0));
  return parts[parts.length - 1];
}

function durationText(ms: number) {
  const total = Math.max(0, Math.floor(ms / 1000));
  const hours = Math.floor(total / 3600);
  const minutes = Math.floor((total % 3600) / 60);
  const seconds = total % 60;
  if (hours > 0) return `${hours}h ${minutes}m ${seconds}s`;
  if (minutes > 0) return `${minutes}m ${seconds}s`;
  return `${seconds}s`;
}

function toolLine(toolName: string, args: Record<string, unknown>) {
  return formatToolCall(toolName, args).description;
}

function toolCountText(count: number, running: boolean) {
  if (count === 0) return running ? "Waiting for first tool call" : "No tool calls used";
  const label = `${count} tool call${count === 1 ? "" : "s"}`;
  return running ? `${label} so far` : label;
}

function useTaskInfo(sessionId: string, active: boolean) {
  const [state, setState] = useState<{ count: number | null; last: string }>({ count: null, last: "" });

  useEffect(() => {
    if (!sessionId) {
      setState({ count: null, last: "" });
      return;
    }

    let cancelled = false;

    const load = async () => {
      try {
        const response = await fetch(`/api/opencode/session/${encodeURIComponent(sessionId)}/messages`);
        if (!response.ok) return;
        const items = z.array(SessionMessageItemSchema).catch([]).parse(await response.json());
        const part = lastDoneTool(items);
        if (!cancelled) {
          setState({
            count: countTools(items),
            last: part?.tool ? toolLine(part.tool, part.state?.input || {}) : "",
          });
        }
      } catch {
        if (!cancelled) setState({ count: null, last: "" });
      }
    };

    void load();
    if (!active) {
      return () => {
        cancelled = true;
      };
    }

    const timer = window.setInterval(() => {
      void load();
    }, 1200);

    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [sessionId, active]);

  return state;
}

interface MessageListProps {
  messages: Message[];
  liveMessages?: Message[];
  isTyping: boolean;
  isConnecting?: boolean;
  currentActivity?: string;
  debugEvents?: DebugEvent[];
  hostLabel?: string;
  showThinking?: boolean;
  showToolResponses?: boolean;
}

const toolConfig: Record<string, { icon: string; format: (args: Record<string, unknown>) => string }> = {
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
  const config = toolConfig[toolName] || getOfficeToolUi(toolName);
  if (config) {
    return { icon: config.icon, description: config.format(args) };
  }
  // Fallback for unknown tools
  return { icon: "🔧", description: toolName.replace(/_/g, " ") };
}

const useStyles = makeStyles({
  chatContainer: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    padding: "22px 0 0",
    display: "flex",
    flexDirection: "column",
    gap: "18px",
    scrollbarColor: "var(--oc-text-faint) transparent",
    scrollbarWidth: "thin",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      backgroundColor: "transparent",
    },
    "&::-webkit-scrollbar-thumb": {
      backgroundColor: "var(--oc-text-faint)",
      borderRadius: "999px",
      border: "2px solid transparent",
      backgroundClip: "content-box",
    },
    "&::-webkit-scrollbar-thumb:hover": {
      backgroundColor: "var(--oc-text-muted)",
    },
  },
  content: {
    width: "100%",
    maxWidth: "760px",
    margin: "0 auto",
    padding: "0 14px 24px",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    gap: "18px",
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
    color: "var(--text-weak)",
    gap: "10px",
  },
  emptyTitle: {
    fontSize: "26px",
    lineHeight: "1.2",
    color: "var(--text-strong)",
    fontWeight: "500",
  },
  emptyMeta: {
    fontSize: "13px",
    color: "var(--text-base)",
  },
  assistantIcon: {
    display: "none",
  },
  messageUser: {
    alignSelf: "stretch",
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: "8px",
    width: "100%",
    color: "var(--text-strong)",
  },
  userBody: {
    width: "fit-content",
    maxWidth: "min(82%, 64ch)",
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: "8px",
  },
  userText: {
    display: "inline-block",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowWrap: "anywhere",
    background: "var(--oc-user-bubble)",
    border: "1px solid var(--oc-user-border)",
    padding: "8px 12px",
    borderRadius: "6px",
    maxWidth: "100%",
  },
  messageAssistant: {
    alignSelf: "stretch",
    width: "100%",
    minWidth: 0,
    boxSizing: "border-box",
    color: "var(--text-strong)",
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
    boxSizing: "border-box",
    lineHeight: "1.6",
    color: "var(--text-strong)",
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
  messageThinking: {
    alignSelf: "stretch",
    width: "100%",
    minWidth: 0,
    boxSizing: "border-box",
    color: "var(--text-weak)",
    fontSize: "14px",
  },
  thinkingHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
    fontSize: "14px",
    color: "var(--text-weak)",
  },
  thinkingTitle: {
    fontWeight: 400,
    color: "var(--text-base)",
    overflowWrap: "anywhere",
  },
  thinkingBody: {
    minWidth: 0,
    lineHeight: "1.6",
    color: "var(--text-weak)",
    overflowWrap: "anywhere",
    "& p:first-child": {
      marginTop: 0,
    },
    "& p:last-child": {
      marginBottom: 0,
    },
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
    alignSelf: "stretch",
    width: "100%",
  },
  messageTask: {
    alignSelf: "stretch",
    width: "100%",
  },
  toolCard: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  toolTrigger: {
    display: "flex",
    alignItems: "baseline",
    gap: "8px",
    width: "100%",
    minWidth: 0,
    cursor: "pointer",
  },
  toolMain: {
    display: "flex",
    alignItems: "baseline",
    gap: "8px",
    minWidth: 0,
    flexWrap: "wrap",
    color: "var(--text-base)",
    fontSize: "14px",
    lineHeight: "1.5",
  },
  toolTitleText: {
    color: "var(--text-strong)",
    fontWeight: 500,
  },
  toolMetaText: {
    color: "var(--text-base)",
    minWidth: 0,
    overflowWrap: "anywhere",
  },
  toolStatus: {
    marginLeft: "auto",
    color: "var(--text-weak)",
    fontSize: "12px",
    whiteSpace: "nowrap",
  },
  taskHead: {
    display: "contents",
  },
  taskBody: {
    display: "contents",
  },
  taskTitle: {
    display: "contents",
  },
  taskTitleText: {
    color: "var(--text-strong)",
    fontWeight: 500,
  },
  taskMeta: {
    display: "contents",
  },
  taskCount: {
    color: "var(--text-base)",
    fontWeight: 400,
  },
  taskBadge: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    padding: "2px 8px",
    borderRadius: "999px",
    fontSize: "11px",
    fontWeight: 600,
    background: "var(--oc-bg-soft)",
    color: "var(--text-weak)",
    border: "1px solid var(--oc-border)",
    whiteSpace: "nowrap",
  },
  taskBadgeDone: {
    color: "var(--oc-success)",
  },
  taskBadgeError: {
    color: "var(--oc-danger-text)",
  },
  taskSpinner: {
    width: "10px",
    height: "10px",
    borderRadius: "50%",
    border: "2px solid var(--oc-border)",
    borderTopColor: "var(--oc-accent)",
    animationName: {
      from: { transform: "rotate(0deg)" },
      to: { transform: "rotate(360deg)" },
    },
    animationDuration: "0.8s",
    animationTimingFunction: "linear",
    animationIterationCount: "infinite",
  },
  toolIcon: {
    fontSize: "14px",
    flexShrink: 0,
  },
  toolArgs: {
    fontSize: "12px",
    fontFamily: '"IBM Plex Mono", "SFMono-Regular", "Consolas", monospace',
    whiteSpace: "pre-wrap",
    color: "var(--text-base)",
    overflowWrap: "anywhere",
  },
  toolDetail: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    width: "100%",
    marginLeft: "22px",
    padding: "12px",
    borderRadius: "12px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
  },
  toolLabel: {
    fontSize: "10px",
    fontWeight: 700,
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    color: "var(--text-weak)",
  },
  attachmentContainer: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    justifyContent: "flex-end",
  },
  attachmentThumbnail: {
    width: "64px",
    height: "64px",
    borderRadius: "10px",
    objectFit: "cover",
    border: "1px solid var(--oc-border)",
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
    color: "var(--text-weak)",
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
            background-color: var(--text-weak, #666);
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
            background: var(--oc-bg-soft, #e0e0e0);
            overflow: hidden;
            margin-top: 6px;
          }
          .activity-progress-fill {
            height: 100%;
            width: 25%;
            border-radius: 1px;
            background: var(--oc-accent, #0078d4);
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
    <span style={{ fontSize: '11px', color: 'var(--text-weak, #999)', marginLeft: '6px' }}>
      {elapsed}s
    </span>
  );
};

function formatBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(1)} MB`;
}

function cleanThinking(value: string) {
  return value
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\([^\)]+\)/g, "$1")
    .replace(/[*_~]+/g, "")
    .trim();
}

function thinkingHeading(text: string) {
  const markdown = text.replace(/\r\n?/g, "\n");
  const html = markdown.match(/<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/i);
  if (html?.[1]) {
    const value = cleanThinking(html[1].replace(/<[^>]+>/g, " "));
    if (value) return value;
  }
  const atx = markdown.match(/^\s{0,3}#{1,6}[ \t]+(.+?)(?:[ \t]+#+[ \t]*)?$/m);
  if (atx?.[1]) {
    const value = cleanThinking(atx[1]);
    if (value) return value;
  }
  const setext = markdown.match(/^([^\n]+)\n(?:=+|-+)\s*$/m);
  if (setext?.[1]) {
    const value = cleanThinking(setext[1]);
    if (value) return value;
  }
  const strong = markdown.match(/^\s*(?:\*\*|__)(.+?)(?:\*\*|__)\s*$/m);
  if (strong?.[1]) {
    const value = cleanThinking(strong[1]);
    if (value) return value;
  }
  return "";
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
      color: 'var(--text-weak, #999)',
      marginLeft: '8px',
      transition: 'color 0.2s',
    }}>
      <span style={{ color: flash ? 'var(--oc-accent, #0078d4)' : undefined, transition: 'color 0.2s' }}>
        ↓{formatBytes(stats.bytesIn)}
      </span>
      <span>↑{formatBytes(stats.bytesOut)}</span>
    </span>
  );
};

function useNow(active: boolean) {
  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    if (!active) return;
    setNow(Date.now());
    const timer = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, [active]);

  return now;
}

const TaskToolMessage: React.FC<{
  message: Message;
  expanded: boolean;
  toggle: () => void;
  showToolResponses: boolean;
}> = ({ message, expanded, toggle, showToolResponses }) => {
  const styles = useStyles();
  const sessionId = taskSessionId(message);
  const running = message.toolStatus === "running";
  const done = message.toolStatus === "completed";
  const failed = message.toolStatus === "error";
  const info = useTaskInfo(sessionId, running);
  const now = useNow(running);
  const started = message.startedAt?.getTime() || message.timestamp.getTime();
  const finished = message.finishedAt?.getTime() || (running ? now : message.timestamp.getTime());
  const elapsed = durationText(Math.max(0, finished - started));
  const meta = running
    ? info.last || "Starting subagent work"
    : [toolCountText(info.count ?? 0, false), elapsed].filter(Boolean).join(" • ");

  return (
    <div className={styles.messageTask}>
      <div className={styles.toolCard}>
        <div className={styles.toolTrigger} onClick={toggle} title="Click to show details">
          <span className={styles.toolIcon}>🧠</span>
          <div className={styles.toolMain}>
            <span className={styles.taskTitleText}>{summarizeTaskTool(message.toolArgs || {})}</span>
            <span className={styles.toolMetaText}>{meta}</span>
          </div>
          <div className={`${styles.taskBadge} ${done ? styles.taskBadgeDone : ""} ${failed ? styles.taskBadgeError : ""}`.trim()}>
            {running && <span className={styles.taskSpinner} />}
            <span>{running ? "Running" : failed ? "Error" : "Done"}</span>
          </div>
        </div>
      </div>

      {expanded && (
        <div className={styles.toolDetail}>
          <div>
            <div className={styles.toolLabel}>Input</div>
            <div className={styles.toolArgs}>{message.text}</div>
          </div>
          {showToolResponses && typeof message.toolResult !== "undefined" && (
            <div>
              <div className={styles.toolLabel}>Response</div>
              <div className={styles.toolArgs}>{typeof message.toolResult === "string" ? message.toolResult : JSON.stringify(message.toolResult, null, 2)}</div>
            </div>
          )}
          {showToolResponses && message.toolError && (
            <div>
              <div className={styles.toolLabel}>Error</div>
              <div className={styles.toolArgs}>{message.toolError}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

const ToolMessage: React.FC<{
  message: Message;
  expanded: boolean;
  toggle: () => void;
  showToolResponses: boolean;
}> = ({ message, expanded, toggle, showToolResponses }) => {
  const styles = useStyles();
  const toolDisplay = message.toolName
    ? formatToolCall(message.toolName, message.toolArgs || {})
    : null;
  const toolTitle = message.toolName ? message.toolName.replace(/_/g, " ") : "tool";

  if (message.toolName === "task") {
    return <TaskToolMessage message={message} expanded={expanded} toggle={toggle} showToolResponses={showToolResponses} />;
  }

  if (!toolDisplay) return null;

  return (
    <div className={styles.messageTool}>
      <div className={styles.toolCard}>
        <div className={styles.toolTrigger} onClick={toggle} title="Click to show details">
          <span className={styles.toolIcon}>{toolDisplay.icon}</span>
          <div className={styles.toolMain}>
            <span className={styles.toolTitleText}>{toolTitle}</span>
            <span className={styles.toolMetaText}>{toolDisplay.description}</span>
          </div>
          <span className={styles.toolStatus}>{message.toolStatus === "running" ? "Running" : message.toolStatus === "error" ? "Error" : "Done"}</span>
        </div>
      </div>
      {expanded && (
        <div className={styles.toolDetail}>
          <div>
            <div className={styles.toolLabel}>Input</div>
            <div className={styles.toolArgs}>{message.text}</div>
          </div>
          {showToolResponses && typeof message.toolResult !== "undefined" && (
            <div>
              <div className={styles.toolLabel}>Response</div>
              <div className={styles.toolArgs}>{typeof message.toolResult === "string" ? message.toolResult : JSON.stringify(message.toolResult, null, 2)}</div>
            </div>
          )}
          {showToolResponses && message.toolError && (
            <div>
              <div className={styles.toolLabel}>Error</div>
              <div className={styles.toolArgs}>{message.toolError}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export const MessageList: React.FC<MessageListProps> = ({
  messages,
  liveMessages = [],
  isTyping,
  isConnecting,
  currentActivity,
  debugEvents,
  hostLabel,
  showThinking = true,
  showToolResponses = false,
}) => {
  const styles = useStyles();
  const chatRef = useRef<HTMLDivElement>(null);
  const stickRef = useRef(true);
  const [expandedTools, setExpandedTools] = useState<Set<string>>(new Set());
  const safeMessages = React.useMemo(() => z.array(MessageSchema).catch([]).parse(messages), [messages]);
  const safeLiveMessages = React.useMemo(() => z.array(MessageSchema).catch([]).parse(liveMessages), [liveMessages]);
  const safeDebugEvents = React.useMemo(
    () => (debugEvents ? z.array(DebugEventSchema).catch([]).parse(debugEvents) : undefined),
    [debugEvents],
  );

  useEffect(() => {
    const el = chatRef.current;
    if (!el) return;

    const near = () => el.scrollHeight - el.scrollTop - el.clientHeight <= 24;
    const onScroll = () => {
      stickRef.current = near();
    };

    stickRef.current = near();
    el.addEventListener("scroll", onScroll);
    return () => el.removeEventListener("scroll", onScroll);
  }, []);

  useLayoutEffect(() => {
    if (!stickRef.current) return;
    const el = chatRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [liveMessages, messages]);

  const toggleTool = (id: string) => {
    setExpandedTools(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const visibleHistory = React.useMemo(
    () => safeMessages.filter((message) => showThinking || message.sender !== "thinking"),
    [safeMessages, showThinking],
  );
  const visibleLive = React.useMemo(
    () => safeLiveMessages.filter((message) => showThinking || message.sender !== "thinking"),
    [safeLiveMessages, showThinking],
  );

  return (
    <div ref={chatRef} className={styles.chatContainer}>
      <div className={styles.content}>
        {safeMessages.length === 0 && !isConnecting && (
          <div className={styles.emptyState}>
            <div className={styles.emptyTitle}>What can I do for you?</div>
            <div className={styles.emptyMeta}>{hostLabel ? `Connected to ${hostLabel}` : "Connected to your document"}</div>
          </div>
        )}

        {isConnecting && (
          <div className={styles.emptyState}>
            <div className={styles.emptyTitle}>Connecting...</div>
          </div>
        )}

        {[...visibleHistory, ...visibleLive].map((message) => {
        return (
        <div
          key={message.id}
          className={
            message.sender === "user" ? styles.messageUser : 
             message.sender === "tool" ? undefined :
              message.sender === "thinking" ? styles.messageThinking :
              styles.messageAssistant
            }
        >
          {message.sender === "tool" ? (
            <ToolMessage
              message={message}
              expanded={expandedTools.has(message.id)}
              toggle={() => toggleTool(message.id)}
              showToolResponses={showToolResponses}
            />
          ) : message.sender === "assistant" ? (
            <div className={styles.assistantBody}><Markdown remarkPlugins={[remarkGfm]}>{message.text}</Markdown></div>
          ) : message.sender === "thinking" ? (
            <>
              <div className={styles.thinkingHeader}>
                <span>Thinking</span>
                <span className={styles.thinkingTitle}>{thinkingHeading(message.text) || currentActivity || "Reasoning"}</span>
              </div>
              <div className={styles.thinkingBody}><Markdown remarkPlugins={[remarkGfm]}>{message.text}</Markdown></div>
            </>
          ) : (
            <div className={styles.userBody}>
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
              <div className={styles.userText}>{message.text}</div>
            </div>
          )}
        </div>
      );
        })}

        {isTyping && visibleLive.length === 0 && (
          <div className={styles.messageAssistant}>
            <div className={styles.assistantBody}>
              <>
                <span className={styles.streamingIndicator}>
                  {currentActivity || "Thinking"}
                  <StreamingDots />
                  <ElapsedTime />
                </span>
                <div className="activity-progress-bar"><div className="activity-progress-fill" /></div>
              </>
              <TrafficCounter />
                {safeDebugEvents && safeDebugEvents.length > 0 && (
                 <div style={{
                   marginTop: '8px',
                   maxHeight: '120px',
                   overflowY: 'auto',
                   fontSize: '10px',
                   fontFamily: 'monospace',
                   lineHeight: '1.6',
                   color: 'var(--text-weak, #999)',
                   backgroundColor: 'var(--oc-bg-soft, #f5f5f5)',
                   borderRadius: '8px',
                   padding: '6px 8px',
                   border: '1px solid var(--oc-border, #e5e5e5)',
                 }}>
                  {safeDebugEvents.map((ev, i) => (
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

        <div />
      </div>
    </div>
  );
};
