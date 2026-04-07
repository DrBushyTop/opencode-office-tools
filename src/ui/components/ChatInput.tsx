import * as React from "react";
import { useRef, useEffect, useState } from "react";
import { Textarea, Button, Combobox, Option, Tooltip, makeStyles } from "@fluentui/react-components";
import { Send24Regular, Dismiss24Regular, Stop24Regular, ChevronDown20Regular, ChevronUp20Regular } from "@fluentui/react-icons";
import { z } from "zod";
import { filterModels } from "../lib/model-search";
import type { ModelInfo, TodoItem } from "../lib/opencode-client";
import { modelInfoSchema } from "../lib/opencode-schemas";
import type { SlashCommand } from "../lib/opencode-schemas";
import { trafficStats } from "../lib/opencode-events";
import { liveStatusText, type Message } from "./MessageList";
import { insertMention, mentionQuery } from "../lib/file-mentions";
import { MentionPopover } from "./MentionPopover";
import { SlashPopover, filterCommands } from "./SlashPopover";

const ImageAttachmentSchema = z.object({
  id: z.string().min(1),
  dataUrl: z.string().min(1),
  name: z.string().min(1),
});

export type ImageAttachment = z.infer<typeof ImageAttachmentSchema>;

interface ChatInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: () => void;
  onStop?: () => void;
  disabled?: boolean;
  isRunning?: boolean;
  images?: ImageAttachment[];
  onImagesChange?: (images: ImageAttachment[]) => void;
  selectedModel: string;
  onModelChange: (model: string) => void;
  models: ModelInfo[];
  selectedVariant?: string;
  onVariantChange: (variant: string | undefined) => void;
  liveMessages?: Message[];
  currentActivity?: string;
  todos?: TodoItem[];
  slashCommands?: SlashCommand[];
  onSearchMentions?: (query: string) => Promise<string[]>;
}

const BYTE_UNITS = ["B", "KB", "MB", "GB"] as const;

function humanizeBytes(raw: number): string {
  let value = Math.max(0, raw);
  let idx = 0;
  while (value >= 1024 && idx < BYTE_UNITS.length - 1) {
    value /= 1024;
    idx++;
  }
  return idx === 0
    ? `${Math.round(value)} ${BYTE_UNITS[idx]}`
    : `${value.toFixed(1)} ${BYTE_UNITS[idx]}`;
}

const RunTimer: React.FC = () => {
  const [seconds, setSeconds] = useState(0);
  const origin = useRef(Date.now());

  useEffect(() => {
    origin.current = Date.now();
    setSeconds(0);
    const handle = setInterval(() => {
      setSeconds(Math.floor((Date.now() - origin.current) / 1000));
    }, 1000);
    return () => clearInterval(handle);
  }, []);

  if (seconds < 3) return null;
  return (
    <span style={{ fontSize: "11px", color: "var(--text-weak, #999)" }}>
      {seconds}s
    </span>
  );
};

const BandwidthMeter: React.FC = () => {
  const [rx, setRx] = useState(0);
  const [tx, setTx] = useState(0);

  useEffect(() => {
    const tick = setInterval(() => {
      setRx(trafficStats.bytesIn);
      setTx(trafficStats.bytesOut);
    }, 500);
    return () => clearInterval(tick);
  }, []);

  if (rx === 0 && tx === 0) return null;
  return (
    <span style={{ fontSize: "11px", color: "var(--text-weak, #999)" }}>
      {"\u2193"}{humanizeBytes(rx)} {"\u2191"}{humanizeBytes(tx)}
    </span>
  );
};

const statusIcons: Record<string, string> = {
  completed: "\u2713",
  in_progress: "\u25CF",
  pending: "\u25CB",
  cancelled: "\u2715",
};

const TodoDock: React.FC<{ todos: TodoItem[] }> = ({ todos }) => {
  const [expanded, setExpanded] = useState(false);
  const done = todos.filter((t) => t.status === "completed").length;
  const active = todos.find((t) => t.status === "in_progress");

  return (
    <div style={{
      borderRadius: "10px 10px 0 0",
      border: "1px solid var(--oc-border)",
      borderBottom: "none",
      background: "var(--oc-bg-soft)",
      marginBottom: "-1px",
      position: "relative",
      zIndex: 1,
      overflow: "hidden",
    }}>
      <button
        type="button"
        onClick={() => setExpanded((v) => !v)}
        style={{
          display: "flex",
          alignItems: "center",
          gap: "6px",
          width: "100%",
          padding: "6px 10px",
          background: "none",
          border: "none",
          cursor: "pointer",
          color: "var(--text-strong)",
          fontSize: "13px",
          lineHeight: "1",
          fontFamily: "inherit",
          textAlign: "left",
        }}
      >
        <span style={{ fontSize: "12px", lineHeight: "1", color: "var(--text-weak)", fontWeight: 600 }}>
          {done}/{todos.length}
        </span>
        {active && (
          <span style={{
            flex: 1,
            minWidth: 0,
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
            color: "var(--text-base, var(--text-strong))",
          }}>
            {active.content}
          </span>
        )}
        {!active && (
          <span style={{ flex: 1, color: "var(--text-weak)" }}>Tasks</span>
        )}
        {expanded
          ? <ChevronDown20Regular style={{ fontSize: "14px", color: "var(--text-weak)", flexShrink: 0 }} />
          : <ChevronUp20Regular style={{ fontSize: "14px", color: "var(--text-weak)", flexShrink: 0 }} />}
      </button>
      {expanded && (
        <div style={{
          padding: "0 10px 8px",
          maxHeight: "140px",
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: "2px",
        }}>
          {todos.map((todo, i) => (
            <div
              key={i}
              style={{
                display: "flex",
                alignItems: "flex-start",
                gap: "6px",
                fontSize: "13px",
                lineHeight: "1.4",
                color: todo.status === "completed" || todo.status === "cancelled"
                  ? "var(--text-weak)"
                  : "var(--text-strong)",
                opacity: todo.status === "pending" ? 0.8 : 1,
              }}
            >
              <span style={{
                flexShrink: 0,
                width: "14px",
                textAlign: "center",
                color: todo.status === "in_progress" ? "var(--oc-accent)" : "var(--text-weak)",
                fontSize: "12px",
                lineHeight: "1.4",
              }}>
                {statusIcons[todo.status] || statusIcons.pending}
              </span>
              <span style={{
                textDecoration: todo.status === "completed" || todo.status === "cancelled"
                  ? "line-through" : "none",
                minWidth: 0,
                wordBreak: "break-word",
              }}>
                {todo.content}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

const useStyles = makeStyles({
  dock: {
    width: "min(calc(100% - 24px), 760px)",
    margin: "0 auto 12px",
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    boxSizing: "border-box",
  },
  statusBar: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "6px 12px",
    borderRadius: "10px 10px 0 0",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 22%, var(--oc-border) 78%)",
    borderBottom: "none",
    background: "color-mix(in srgb, var(--oc-bg-strong) 85%, transparent)",
    backdropFilter: "blur(8px)",
    marginBottom: "-1px",
    position: "relative",
    zIndex: 1,
    fontSize: "11px",
    lineHeight: "1",
  },
  statusDot: {
    width: "7px",
    height: "7px",
    borderRadius: "50%",
    background: "var(--oc-accent)",
    flexShrink: 0,
    animationName: {
      "0%": { opacity: 1 },
      "50%": { opacity: 0.4 },
      "100%": { opacity: 1 },
    },
    animationDuration: "1.2s",
    animationTimingFunction: "ease-in-out",
    animationIterationCount: "infinite",
  },
  statusLabel: {
    fontWeight: 600,
    letterSpacing: "0.03em",
    textTransform: "uppercase",
    color: "var(--oc-accent)",
    flexShrink: 0,
    lineHeight: "1",
  },
  statusActivity: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    color: "var(--text-strong)",
    fontWeight: 500,
    fontSize: "12px",
    lineHeight: "1",
  },
  statusMeta: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    flexShrink: 0,
  },
  tray: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "6px 10px",
    borderRadius: "12px 12px 0 0",
    border: "1px solid var(--oc-border)",
    borderBottom: "none",
    background: "var(--oc-bg-soft)",
    marginBottom: "-1px",
    position: "relative",
    zIndex: 1,
  },
  trayHasStatus: {
    borderTopLeftRadius: "0",
    borderTopRightRadius: "0",
  },
  trayField: {
    flex: "1 1 0",
    minWidth: 0,
  },
  variantField: {
    flex: "0 0 100px",
    minWidth: "80px",
  },
  control: {
    minWidth: 0,
    width: "100%",
    height: "28px",
    fontSize: "11px",
    borderRadius: "8px",
    background: "var(--oc-bg)",
    border: "1px solid var(--oc-border) !important",
    padding: "0 8px",
    boxSizing: "border-box",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
    },
  },
  shell: {
    display: "flex",
    flexDirection: "column",
    borderRadius: "12px",
    backgroundColor: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    position: "relative",
    zIndex: 2,
  },
  shellHasAbove: {
    borderTopLeftRadius: "0",
    borderTopRightRadius: "0",
  },
  body: {
    display: "flex",
    flexDirection: "column",
    padding: "10px 12px 8px",
  },
  input: {
    flex: 1,
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
    minHeight: "52px",
    padding: "0",
    borderRadius: "0",
    border: "none !important",
    backgroundColor: "transparent !important",
    outline: "none !important",
    boxShadow: "none !important",
    color: "var(--text-strong)",
    fontSize: "14px",
    lineHeight: "1.5",
    "::after": {
      display: "none !important",
    },
    "& textarea": {
      padding: "0",
      minHeight: "52px",
    },
  },
  inputWrap: {
    flex: 1,
    width: "100%",
    minWidth: 0,
    padding: "0 0 4px",
  },
  footer: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    padding: "0 12px 10px",
    gap: "6px",
  },
  sendButton: {
    width: "34px",
    height: "34px",
    minWidth: "34px",
    padding: "0",
    backgroundColor: "var(--oc-accent)",
    border: "none",
    borderRadius: "10px",
    color: "var(--text-on-interactive-base, white)",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  stopButton: {
    backgroundColor: "var(--oc-bg-soft)",
    color: "var(--text-strong)",
    border: "1px solid var(--oc-border)",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
  attachmentStrip: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    padding: "8px 12px 0",
  },
  thumbnailFrame: {
    position: "relative",
    width: "56px",
    height: "56px",
    borderRadius: "8px",
    overflow: "hidden",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
  },
  thumbnailCover: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
  },
  thumbnailDismiss: {
    position: "absolute",
    top: "3px",
    right: "3px",
    minWidth: "18px",
    width: "18px",
    height: "18px",
    padding: "0",
    backgroundColor: "var(--oc-bg)",
    border: "1px solid var(--oc-border)",
    borderRadius: "50%",
    cursor: "pointer",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
});

export const ChatInput: React.FC<ChatInputProps> = ({
  value,
  onChange,
  onSend,
  onStop,
  disabled = false,
  isRunning = false,
  images = [],
  onImagesChange,
  selectedModel,
  onModelChange,
  models,
  selectedVariant,
  onVariantChange,
  liveMessages = [],
  currentActivity = "",
  todos = [],
  slashCommands = [],
  onSearchMentions,
}) => {
  const styles = useStyles();
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const safeImages = React.useMemo(() => z.array(ImageAttachmentSchema).catch([]).parse(images), [images]);
  const safeModels = React.useMemo(() => z.array(modelInfoSchema).catch([]).parse(models), [models]);
  const [modelValue, setModelValue] = React.useState("");
  const [modelOpen, setModelOpen] = React.useState(false);
  const selectedLabel = React.useMemo(
    () => safeModels.find((item) => item.key === selectedModel)?.label || selectedModel,
    [safeModels, selectedModel],
  );
  const modelItems = React.useMemo(
    () => filterModels(safeModels, modelOpen ? modelValue : ""),
    [safeModels, modelOpen, modelValue],
  );
  const modelVariants = React.useMemo(() => {
    const current = safeModels.find((item) => item.key === selectedModel);
    return current?.variants ?? [];
  }, [safeModels, selectedModel]);

  const activity = React.useMemo(
    () => liveStatusText(liveMessages, currentActivity),
    [liveMessages, currentActivity],
  );

  /* ---- slash command autocomplete state ---- */
  const slashMatch = React.useMemo(() => {
    // Trigger: input starts with "/" and has no whitespace (just the command name so far)
    const m = value.match(/^\/(\S*)$/);
    return m ? m[1] : null;
  }, [value]);

  const showSlash = slashMatch !== null && slashCommands.length > 0;
  const [slashIndex, setSlashIndex] = useState(0);
  const [mentionIndex, setMentionIndex] = useState(0);
  const [mentionItems, setMentionItems] = useState<string[]>([]);
  const [mentionLoading, setMentionLoading] = useState(false);
  const [caret, setCaret] = useState(0);

  // Reset highlight when filter changes
  useEffect(() => {
    setSlashIndex(0);
  }, [slashMatch]);

  const filteredSlashCount = React.useMemo(
    () => (showSlash ? filterCommands(slashCommands, slashMatch || "").length : 0),
    [showSlash, slashCommands, slashMatch],
  );
  const currentMention = React.useMemo(() => mentionQuery(value, caret), [value, caret]);
  const mentionText = currentMention?.query || "";
  const showMention = !showSlash && !!currentMention && !!onSearchMentions;

  // Clamp index when filtered results shrink
  useEffect(() => {
    setSlashIndex((prev) =>
      filteredSlashCount === 0 ? 0 : Math.min(prev, filteredSlashCount - 1),
    );
  }, [filteredSlashCount]);

  useEffect(() => {
    setMentionIndex(0);
  }, [mentionText]);

  useEffect(() => {
    if (!showMention || !currentMention || !onSearchMentions) {
      setMentionItems([]);
      setMentionLoading(false);
      return;
    }

    let cancelled = false;
    setMentionLoading(true);
    onSearchMentions(mentionText)
      .then((items) => {
        if (cancelled) return;
        setMentionItems(items);
      })
      .catch(() => {
        if (cancelled) return;
        setMentionItems([]);
      })
      .finally(() => {
        if (!cancelled) setMentionLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [mentionText, onSearchMentions, showMention]);

  useEffect(() => {
    setMentionIndex((prev) => (mentionItems.length === 0 ? 0 : Math.min(prev, mentionItems.length - 1)));
  }, [mentionItems]);

  const handleSlashSelect = React.useCallback(
    (cmd: SlashCommand) => {
      // Insert "/<name> " so the user can add arguments
      onChange(`/${cmd.name} `);
      setSlashIndex(0);
    },
    [onChange],
  );

  const handleMentionSelect = React.useCallback((path: string) => {
    const next = insertMention(value, caret, path);
    if (!next) return;
    onChange(next.value);
    setMentionItems([]);
    setMentionIndex(0);
    requestAnimationFrame(() => {
      inputRef.current?.focus();
      inputRef.current?.setSelectionRange(next.caret, next.caret);
      setCaret(next.caret);
    });
  }, [caret, onChange, value]);

  useEffect(() => {
    if (value === "") {
      inputRef.current?.focus();
    }
  }, [value]);

  useEffect(() => {
    if (!modelOpen) {
      setModelValue(selectedLabel);
    }
  }, [modelOpen, selectedLabel]);

  const onInputKeyDown = (e: React.KeyboardEvent) => {
    if (showMention && mentionItems.length > 0) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        setMentionIndex((prev) => (prev + 1) % mentionItems.length);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        setMentionIndex((prev) => (prev - 1 + mentionItems.length) % mentionItems.length);
        return;
      }
      if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey)) {
        e.preventDefault();
        const selected = mentionItems[mentionIndex];
        if (selected) handleMentionSelect(selected);
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        setMentionItems([]);
        return;
      }
    }
    if (showSlash && filteredSlashCount > 0) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        setSlashIndex((prev) => (prev + 1) % filteredSlashCount);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        setSlashIndex((prev) => (prev - 1 + filteredSlashCount) % filteredSlashCount);
        return;
      }
      if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey)) {
        e.preventDefault();
        const filtered = filterCommands(slashCommands, slashMatch || "");
        const selected = filtered[slashIndex];
        if (selected) {
          handleSlashSelect(selected);
        }
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        onChange("");
        return;
      }
    }
    // When popover is showing with no matches, block Enter to avoid sending
    // a raw "/..." as a normal message
    if (showSlash && filteredSlashCount === 0 && e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      return;
    }
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSend();
    }
  };

  const onClipboardPaste = async (e: React.ClipboardEvent) => {
    const clipboardItems = e.clipboardData?.items;
    if (!clipboardItems || !onImagesChange) return;

    const files = Array.from(clipboardItems)
      .filter((entry) => entry.type.startsWith("image/"))
      .map((entry) => entry.getAsFile())
      .filter((file): file is File => file !== null);

    if (files.length === 0) return;
    e.preventDefault();

    const pending = files.map(
      (file, idx) =>
        new Promise<ImageAttachment | null>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => {
            const result = typeof reader.result === "string" ? reader.result : "";
            const validated = ImageAttachmentSchema.safeParse({
              id: crypto.randomUUID(),
              dataUrl: result,
              name: `clipboard-${Date.now()}-${idx}.png`,
            });
            resolve(validated.success ? validated.data : null);
          };
          reader.onerror = () => resolve(null);
          reader.readAsDataURL(file);
        }),
    );

    const resolved = (await Promise.all(pending)).filter(
      (attachment): attachment is ImageAttachment => attachment !== null,
    );

    if (resolved.length > 0) {
      onImagesChange([...safeImages, ...resolved]);
    }
  };

  const discardAttachment = (id: string) => {
    if (onImagesChange) {
      onImagesChange(safeImages.filter((attachment) => attachment.id !== id));
    }
  };

  const hasTray = safeModels.length > 0;
  const hasStatus = isRunning;
  const hasTodos = todos.length > 0;
  const hasAbove = hasTray || hasStatus || hasTodos;

  return (
    <div className={styles.dock} style={{ position: "relative" }}>
      {showSlash && (
        <SlashPopover
          filter={slashMatch || ""}
          commands={slashCommands}
          selectedIndex={slashIndex}
          onSelect={handleSlashSelect}
          onHighlight={setSlashIndex}
        />
      )}
      {showMention && (
        <MentionPopover
          items={mentionItems}
          selectedIndex={mentionIndex}
          onSelect={handleMentionSelect}
          onHighlight={setMentionIndex}
          loading={mentionLoading}
        />
      )}
      {hasTodos && <TodoDock todos={todos} />}

      {hasStatus && (
        <div
          className={styles.statusBar}
          style={hasTodos ? { borderTopLeftRadius: 0, borderTopRightRadius: 0 } : undefined}
        >
          <span className={styles.statusDot} />
          <span className={styles.statusLabel}>Running</span>
          <span className={styles.statusActivity}>{activity}</span>
          <span className={styles.statusMeta}>
            <RunTimer />
            <BandwidthMeter />
          </span>
        </div>
      )}

      {hasTray && (
        <div
          className={`${styles.tray} ${hasStatus || hasTodos ? styles.trayHasStatus : ""}`.trim()}
        >
          <div className={styles.trayField}>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              freeform
              placeholder="Search models"
              aria-label="Model"
              value={modelValue}
              onChange={(event) => setModelValue((event.target as HTMLInputElement).value)}
              onOpenChange={(_, data) => {
                setModelOpen(data.open);
                setModelValue(data.open ? "" : selectedLabel);
              }}
              onOptionSelect={(_, data) => {
                const nextModel = data.optionValue;
                if (nextModel && nextModel !== selectedModel) {
                  onModelChange(nextModel);
                }
                setModelOpen(false);
                setModelValue(data.optionText || selectedLabel);
              }}
            >
              {modelItems.map((model) => (
                <Option key={model.key} value={model.key} text={model.label}>
                  {model.label}
                </Option>
              ))}
            </Combobox>
          </div>
          <div className={styles.variantField}>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              placeholder="Effort"
              aria-label="Model effort"
              value={selectedVariant ?? "default"}
              disabled={modelVariants.length === 0}
              onOptionSelect={(_, data) => onVariantChange(data.optionValue || undefined)}
            >
              <Option value="" text="default">
                default
              </Option>
              {modelVariants.map((variant) => (
                <Option key={variant} value={variant} text={variant}>
                  {variant}
                </Option>
              ))}
            </Combobox>
          </div>
        </div>
      )}
      <div className={`${styles.shell} ${hasAbove ? styles.shellHasAbove : ""}`.trim()}>
        {safeImages.length > 0 && (
          <div className={styles.attachmentStrip}>
            {safeImages.map((image) => (
              <div key={image.id} className={styles.thumbnailFrame}>
                <img src={image.dataUrl} alt="Preview" className={styles.thumbnailCover} />
                <button
                  type="button"
                  className={styles.thumbnailDismiss}
                  onClick={() => discardAttachment(image.id)}
                  title="Remove image"
                  aria-label={`Remove image ${image.name}`}
                >
                  <Dismiss24Regular style={{ fontSize: '12px' }} />
                </button>
              </div>
            ))}
          </div>
        )}
        <div className={styles.body}>
          <div className={styles.inputWrap}>
            <Textarea
              ref={inputRef}
              className={styles.input}
              value={value}
              onChange={(e, data) => {
                onChange(data.value);
                const next = e.target as HTMLTextAreaElement;
                setCaret(next.selectionStart ?? data.value.length);
              }}
              onKeyDown={onInputKeyDown}
              onPaste={onClipboardPaste}
              onClick={(event) => setCaret((event.target as HTMLTextAreaElement).selectionStart ?? value.length)}
              onKeyUp={(event) => setCaret((event.target as HTMLTextAreaElement).selectionStart ?? value.length)}
              onSelect={(event) => setCaret((event.target as HTMLTextAreaElement).selectionStart ?? value.length)}
              placeholder="Ask OpenCode to work on the current document..."
              rows={2}
              disabled={disabled}
            />
          </div>
        </div>
        <div className={styles.footer}>
          {isRunning && (
            <Tooltip content="Stop response" relationship="label">
              <Button
                appearance="secondary"
                icon={<Stop24Regular />}
                onClick={onStop}
                disabled={disabled || !onStop}
                aria-label="Stop response"
                className={`${styles.sendButton} ${styles.stopButton}`.trim()}
              />
            </Tooltip>
          )}
          <Tooltip content="Send message" relationship="label">
            <Button
              appearance="primary"
              icon={<Send24Regular />}
              onClick={onSend}
              disabled={disabled || (!value.trim() && safeImages.length === 0)}
              aria-label="Send message"
              className={styles.sendButton}
            />
          </Tooltip>
        </div>
      </div>
    </div>
  );
};
