import * as React from "react";
import { useRef, useEffect, useLayoutEffect, useState } from "react";
import { makeStyles, mergeClasses } from "@fluentui/react-components";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { z } from "zod";
import { trafficStats } from "../lib/opencode-events";
import { coalesceTranscriptMessages } from "../lib/opencode-session-history";
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
    overflowX: "hidden",
    padding: "14px 0 0",
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
    flexShrink: 0,
  },
  contentEmpty: {
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
    overflowX: "hidden",
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
    overflowWrap: "anywhere",
    wordBreak: "break-word",
    lineHeight: "1.6",
    color: "var(--text-strong)",
    "& pre": {
      background: "var(--oc-bg-strong)",
      border: "1px solid var(--oc-border)",
      borderRadius: "10px",
      padding: "10px 12px",
      overflowX: "auto",
    },
    "& :not(pre) > code": {
      background: "var(--oc-bg-soft)",
      borderRadius: "6px",
      padding: "1px 4px",
    },
    "& table": {
      display: "block",
      overflowX: "auto",
      maxWidth: "100%",
    },
    "& a": {
      wordBreak: "break-all",
    },
  },
  messageThinking: {
    alignSelf: "stretch",
    width: "100%",
    minWidth: 0,
    boxSizing: "border-box",
    padding: "12px 14px",
    borderRadius: "14px",
    background: "var(--oc-thinking-bg)",
    border: "1px solid var(--oc-thinking-border)",
    color: "var(--text-weak)",
    fontSize: "14px",
  },
  thinkingHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
    fontSize: "12px",
    letterSpacing: "0.02em",
    color: "var(--oc-thinking-label)",
  },
  thinkingLabel: {
    color: "var(--oc-thinking-label)",
  },
  thinkingTitle: {
    fontWeight: 500,
    color: "var(--oc-thinking-title)",
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
    "& :not(pre) > code": {
      background: "var(--oc-bg-soft)",
      borderRadius: "6px",
      padding: "1px 4px",
    },
    "& h1, & h2, & h3, & h4, & h5, & h6": {
      color: "var(--oc-thinking-title)",
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
  },
  toolToggle: {
    appearance: "none",
    background: "transparent",
    border: "none",
    padding: 0,
    margin: 0,
    width: "100%",
    textAlign: "left",
    cursor: "pointer",
    color: "inherit",
    font: "inherit",
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
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    padding: "2px 8px",
    borderRadius: "999px",
    color: "var(--text-weak)",
    fontSize: "11px",
    fontWeight: 600,
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    whiteSpace: "nowrap",
  },
  toolStatusRunning: {
    color: "var(--oc-accent)",
    background: "color-mix(in srgb, var(--oc-accent) 12%, var(--oc-bg-soft) 88%)",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 28%, var(--oc-border) 72%)",
  },
  toolStatusDone: {
    color: "var(--oc-success)",
  },
  toolStatusError: {
    color: "var(--oc-danger-text)",
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
  taskBadgeRunning: {
    color: "var(--oc-accent)",
    background: "color-mix(in srgb, var(--oc-accent) 12%, var(--oc-bg-soft) 88%)",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 28%, var(--oc-border) 72%)",
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
    width: "calc(100% - 22px)",
    minWidth: 0,
    boxSizing: "border-box",
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
  liveStatusDock: {
    position: "sticky",
    bottom: "12px",
    zIndex: 1,
  },
  liveStatusCard: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px 14px",
    borderRadius: "14px",
    background: "color-mix(in srgb, var(--oc-bg-strong) 78%, transparent)",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 22%, var(--oc-border) 78%)",
    boxShadow: "0 12px 28px rgba(0, 0, 0, 0.18)",
    backdropFilter: "blur(12px)",
  },
  liveStatusHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  liveStatusDot: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    background: "var(--oc-accent)",
    animationName: {
      "0%": { transform: "scale(1)", opacity: 0.95 },
      "50%": { transform: "scale(1.4)", opacity: 0.35 },
      "100%": { transform: "scale(1)", opacity: 0.95 },
    },
    animationDuration: "1.2s",
    animationTimingFunction: "ease-in-out",
    animationIterationCount: "infinite",
  },
  liveStatusLabel: {
    fontSize: "11px",
    fontWeight: 700,
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    color: "var(--oc-accent)",
  },
  liveStatusTitle: {
    fontSize: "14px",
    fontWeight: 500,
    color: "var(--text-strong)",
    overflowWrap: "anywhere",
  },
  liveStatusMeta: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
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

function leadingThinkingHeading(text: string) {
  const markdown = text.replace(/\r\n?/g, "\n");
  const html = markdown.match(/^\s*<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>\s*(?:\n+)?/i);
  if (html?.[1]) {
    const value = cleanThinking(html[1].replace(/<[^>]+>/g, " "));
    if (value) return { title: value, length: html[0].length };
  }
  const atx = markdown.match(/^\s{0,3}#{1,6}[ \t]+(.+?)(?:[ \t]+#+[ \t]*)?\s*(?:\n+)?/);
  if (atx?.[1]) {
    const value = cleanThinking(atx[1]);
    if (value) return { title: value, length: atx[0].length };
  }
  const setext = markdown.match(/^([^\n]+)\n(?:=+|-+)\s*\n*(?:\n+)?/);
  if (setext?.[1]) {
    const value = cleanThinking(setext[1]);
    if (value) return { title: value, length: setext[0].length };
  }
  const strong = markdown.match(/^\s*(?:\*\*|__)(.+?)(?:\*\*|__)\s*\n*(?:\n+)?/);
  if (strong?.[1]) {
    const value = cleanThinking(strong[1]);
    if (value) return { title: value, length: strong[0].length };
  }
  return null;
}

function thinkingHeading(text: string) {
  return leadingThinkingHeading(text)?.title || "";
}

function stripThinkingHeading(text: string) {
  const markdown = text.replace(/\r\n?/g, "\n");
  const match = leadingThinkingHeading(markdown);
  if (!match) return markdown;
  return markdown.slice(match.length).trim();
}

function liveStatusText(messages: Message[], activity?: string) {
  if (activity?.trim()) return activity.trim();
  const item = [...messages].reverse().find((message) => message.sender !== "user");
  if (!item) return "Generating response";
  if (item.sender === "thinking") return thinkingHeading(item.text) || "Reasoning";
  if (item.sender === "tool" && item.toolName) return formatToolCall(item.toolName, item.toolArgs || {}).description;
  if (item.sender === "assistant") return item.text.trim() ? "Drafting response" : "Generating response";
  return "Working";
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
        <button type="button" className={styles.toolToggle} onClick={toggle} aria-expanded={expanded} title="Click to show details">
          <div className={styles.toolTrigger}>
            <span className={styles.toolIcon}>🧠</span>
            <div className={styles.toolMain}>
              <span className={styles.taskTitleText}>{summarizeTaskTool(message.toolArgs || {})}</span>
              <span className={styles.toolMetaText}>{meta}</span>
            </div>
            <div className={`${styles.taskBadge} ${running ? styles.taskBadgeRunning : ""} ${done ? styles.taskBadgeDone : ""} ${failed ? styles.taskBadgeError : ""}`.trim()}>
              {running && <span className={styles.taskSpinner} />}
              <span>{running ? "Running" : failed ? "Error" : "Done"}</span>
            </div>
          </div>
        </button>
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
  const running = message.toolStatus === "running";
  const failed = message.toolStatus === "error";

  if (message.toolName === "task") {
    return <TaskToolMessage message={message} expanded={expanded} toggle={toggle} showToolResponses={showToolResponses} />;
  }

  if (!toolDisplay) return null;

  return (
    <div className={styles.messageTool}>
      <div className={styles.toolCard}>
        <button type="button" className={styles.toolToggle} onClick={toggle} aria-expanded={expanded} title="Click to show details">
          <div className={styles.toolTrigger}>
            <span className={styles.toolIcon}>{toolDisplay.icon}</span>
            <div className={styles.toolMain}>
              <span className={styles.toolTitleText}>{toolTitle}</span>
              <span className={styles.toolMetaText}>{toolDisplay.description}</span>
            </div>
            <span className={`${styles.toolStatus} ${running ? styles.toolStatusRunning : failed ? styles.toolStatusError : styles.toolStatusDone}`.trim()}>
              {running && <span className={styles.taskSpinner} />}
              <span>{running ? "Running" : failed ? "Error" : "Done"}</span>
            </span>
          </div>
        </button>
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
  const liveStatus = React.useMemo(() => liveStatusText(safeLiveMessages, currentActivity), [safeLiveMessages, currentActivity]);

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
  const visibleMessages = React.useMemo(
    () => coalesceTranscriptMessages(visibleHistory, visibleLive),
    [visibleHistory, visibleLive],
  );

  return (
    <div ref={chatRef} className={styles.chatContainer}>
      <div className={mergeClasses(styles.content, visibleMessages.length === 0 ? styles.contentEmpty : undefined)}>
        {visibleMessages.length === 0 && !isConnecting && (
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

        {visibleMessages.map((message) => {
          const summary = message.sender === "thinking" ? thinkingHeading(message.text) : "";
          const thinkingContent = message.sender === "thinking" ? stripThinkingHeading(message.text) : message.text;

          return (
            <div
              key={message.id}
              className={
                message.sender === "user"
                  ? styles.messageUser
                  : message.sender === "tool"
                    ? undefined
                    : message.sender === "thinking"
                      ? styles.messageThinking
                      : styles.messageAssistant
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
                <div className={styles.assistantBody}><Markdown remarkPlugins={[remarkGfm]} disallowedElements={["img"]}>{message.text}</Markdown></div>
              ) : message.sender === "thinking" ? (
                <>
                  <div className={styles.thinkingHeader}>
                    <span className={styles.thinkingLabel}>Thinking</span>
                    <span className={styles.thinkingTitle}>{summary || currentActivity || "Reasoning"}</span>
                  </div>
                  {thinkingContent ? <div className={styles.thinkingBody}><Markdown remarkPlugins={[remarkGfm]} disallowedElements={["img"]}>{thinkingContent}</Markdown></div> : null}
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

        {isTyping && (
          <div className={styles.liveStatusDock}>
            <div className={styles.liveStatusCard}>
              <div className={styles.liveStatusHeader}>
                <span className={styles.liveStatusDot} />
                <span className={styles.liveStatusLabel}>OpenCode is running</span>
                <span className={styles.liveStatusTitle}>{liveStatus}</span>
              </div>
              <div className={styles.liveStatusMeta}>
                <span className={styles.streamingIndicator}>
                  Streaming
                  <StreamingDots />
                  <ElapsedTime />
                </span>
                <TrafficCounter />
              </div>
              <div className="activity-progress-bar"><div className="activity-progress-fill" /></div>
              {safeDebugEvents && safeDebugEvents.length > 0 && (
                <div style={{
                  marginTop: '2px',
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
