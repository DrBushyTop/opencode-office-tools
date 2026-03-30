import * as React from "react";
import {
  Button,
  Combobox,
  Option,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Switch,
  Tooltip,
  makeStyles,
} from "@fluentui/react-components";
import { Compose24Regular, History24Regular, Settings24Regular } from "@fluentui/react-icons";
import { z } from "zod";
import { filterModels } from "../lib/model-search";
import type { ModelInfo } from "../lib/opencode-client";

const ModelTypeSchema = z.string();
const ModelInfoSchema = z.object({
  key: z.string().min(1),
  label: z.string().min(1),
  providerID: z.string().min(1),
  modelID: z.string().min(1),
  limitContext: z.number().nonnegative().optional(),
}) satisfies z.ZodType<ModelInfo>;

export type ModelType = z.infer<typeof ModelTypeSchema>;

export interface HeaderThemeOption {
  id: string;
  label: string;
  isDefault?: boolean;
}

export interface HeaderConnectionStatus {
  state: "connecting" | "connected" | "disconnected";
  label: string;
}

interface HeaderBarProps {
  onNewChat: () => void;
  onShowHistory: () => void;
  selectedModel: ModelType;
  onModelChange: (model: ModelType) => void;
  models: ModelInfo[];
  debugEnabled: boolean;
  onDebugChange: (value: boolean) => void;
  showThinking: boolean;
  onShowThinkingChange: (value: boolean) => void;
  showToolResponses: boolean;
  onShowToolResponsesChange: (value: boolean) => void;
  qaSubagentModel: ModelType;
  onQaSubagentModelChange: (model: ModelType) => void;
  themes: HeaderThemeOption[];
  selectedThemeId: string;
  onThemeChange: (themeId: string) => void;
  connectionStatus?: HeaderConnectionStatus;
  subtitle?: string;
  contextLabel?: string;
  usageSummary?: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    padding: "12px 16px 8px",
    gap: "6px",
    minHeight: "64px",
    borderBottom: "1px solid var(--oc-border)",
    background: "linear-gradient(180deg, color-mix(in srgb, var(--oc-bg-elevated, var(--oc-bg)) 92%, transparent), var(--oc-bg))",
  },
  topRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
    width: "100%",
    minWidth: 0,
    flexWrap: "nowrap",
  },
  bottomRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    width: "100%",
    minWidth: 0,
  },
  left: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
    flex: 1,
  },
  subtitle: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    flex: "1 1 auto",
    minWidth: 0,
  },
  dropdown: {
    minWidth: 0,
    width: "100%",
    height: "34px",
    fontSize: "12px",
    borderRadius: "10px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border) !important",
    padding: "0 10px",
    boxSizing: "border-box",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
    },
  },
  group: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    flexWrap: "wrap",
    gap: "8px",
    flexShrink: 0,
    minWidth: 0,
  },
  statusPill: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    color: "var(--oc-text-muted)",
    fontSize: "11px",
    lineHeight: 1,
    whiteSpace: "nowrap",
    flexShrink: 0,
  },
  statusDot: {
    width: "7px",
    height: "7px",
    borderRadius: "999px",
    background: "currentColor",
    boxShadow: "0 0 0 2px color-mix(in srgb, currentColor 18%, transparent)",
    color: "inherit",
    transform: "translateY(-0.5px)",
  },
  statusConnecting: {
    color: "var(--oc-warning, #d4a72c)",
  },
  statusConnected: {
    color: "var(--oc-success, #4db56a)",
  },
  statusDisconnected: {
    color: "var(--oc-danger, #d95c5c)",
  },
  usage: {
    fontSize: "11px",
    color: "var(--oc-text-muted)",
    padding: "6px 10px",
    borderRadius: "999px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
    whiteSpace: "nowrap",
  },
  icon: {
    minWidth: "34px",
    width: "34px",
    height: "34px",
    padding: "0",
    borderRadius: "10px",
    color: "var(--oc-text-muted)",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
      color: "var(--oc-text)",
      border: "1px solid var(--oc-border)",
    },
  },
  primary: {
    backgroundColor: "var(--oc-accent)",
    color: "white",
    borderRadius: "10px",
    padding: "4px",
    width: "34px",
    height: "34px",
    minWidth: "34px",
    border: "none",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  contextChip: {
    fontSize: "11px",
    color: "var(--oc-text-muted)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    marginLeft: "auto",
    minWidth: 0,
    textAlign: "right",
    flexShrink: 1,
  },
  menu: {
    width: "320px",
    display: "flex",
    flexDirection: "column",
    gap: "14px",
    padding: "14px",
    background: "var(--oc-bg-elevated, var(--oc-bg, #1b1818))",
    color: "var(--oc-text, #f1ecec)",
    border: "1px solid var(--oc-border, rgba(255,255,255,0.10))",
    borderRadius: "16px",
    boxShadow: "var(--oc-shadow, 0 0 0 1px rgba(255,255,255,0.06), 0 24px 64px rgba(0,0,0,0.32))",
    position: "relative",
    zIndex: 20,
  },
  menuHeader: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    paddingBottom: "2px",
  },
  menuTitle: {
    fontSize: "13px",
    fontWeight: "700",
    color: "var(--oc-text)",
  },
  menuHint: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    lineHeight: "1.5",
  },
  item: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    borderRadius: "14px",
    border: "1px solid var(--oc-border)",
    background: "color-mix(in srgb, var(--oc-bg-soft) 88%, transparent)",
  },
  row: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
  },
  label: {
    fontSize: "12px",
    fontWeight: "600",
    color: "var(--oc-text)",
  },
  valueText: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "120px",
    textAlign: "right",
  },
  help: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    lineHeight: "1.5",
  },
  sectionLabel: {
    fontSize: "11px",
    fontWeight: "700",
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    color: "var(--oc-text-faint)",
  },
});

function label(models: ModelInfo[], model: ModelType) {
  if (!model) return "Same as primary model";
  return models.find((item) => item.key === model)?.label || model;
}

export const HeaderBar: React.FC<HeaderBarProps> = ({
  onNewChat,
  onShowHistory,
  selectedModel,
  onModelChange,
  models,
  debugEnabled,
  onDebugChange,
  showThinking,
  onShowThinkingChange,
  showToolResponses,
  onShowToolResponsesChange,
  qaSubagentModel,
  onQaSubagentModelChange,
  themes,
  selectedThemeId,
  onThemeChange,
  connectionStatus,
  subtitle,
  contextLabel,
  usageSummary,
}) => {
  const styles = useStyles();
  const [value, setValue] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const [qaValue, setQaValue] = React.useState("");
  const [qaOpen, setQaOpen] = React.useState(false);
  const selectedLabel = label(models, selectedModel);
  const qaLabel = label(models, qaSubagentModel);
  const selectedTheme = React.useMemo(
    () => themes.find((theme) => theme.id === selectedThemeId) ?? themes.find((theme) => theme.isDefault) ?? themes[0],
    [selectedThemeId, themes],
  );
  const safeModels = React.useMemo(() => z.array(ModelInfoSchema).catch([]).parse(models), [models]);
  const items = React.useMemo(() => filterModels(safeModels, open ? value : ""), [safeModels, open, value]);
  const qaItems = React.useMemo(() => filterModels(safeModels, qaOpen ? qaValue : ""), [safeModels, qaOpen, qaValue]);
  const statusClassName = React.useMemo(() => {
    switch (connectionStatus?.state) {
      case "connecting":
        return styles.statusConnecting;
      case "connected":
        return styles.statusConnected;
      case "disconnected":
        return styles.statusDisconnected;
      default:
        return undefined;
    }
  }, [connectionStatus?.state, styles.statusConnected, styles.statusConnecting, styles.statusDisconnected]);

  React.useEffect(() => {
    if (!open) setValue(selectedLabel);
  }, [open, selectedLabel]);

  React.useEffect(() => {
    if (!qaOpen) setQaValue(qaLabel);
  }, [qaLabel, qaOpen]);

  return (
    <div className={styles.header}>
      <div className={styles.topRow}>
        <div className={styles.left}>
          <Combobox
            className={styles.dropdown}
            appearance="filled-darker"
          freeform
          placeholder="Search models"
          aria-label="Model"
          value={value}
            onChange={(event) => setValue((event.target as HTMLInputElement).value)}
            onOpenChange={(_, data) => {
              setOpen(data.open);
              setValue(data.open ? "" : selectedLabel);
            }}
            onOptionSelect={(_, data) => {
              const nextModel = ModelTypeSchema.safeParse(data.optionValue);
              if (nextModel.success && nextModel.data !== selectedModel) onModelChange(nextModel.data);
              setOpen(false);
              setValue(data.optionText || selectedLabel);
            }}
          >
            {items.map((model) => (
              <Option key={model.key} value={model.key} text={model.label}>
                {model.label}
              </Option>
            ))}
          </Combobox>
        </div>

        <div className={styles.group}>
          {usageSummary && <div className={styles.usage}>{usageSummary}</div>}
          <Popover positioning="below-end" inline>
            <PopoverTrigger disableButtonEnhancement>
              <Button icon={<Settings24Regular />} appearance="subtle" aria-label="Options" className={styles.icon} />
            </PopoverTrigger>
            <PopoverSurface className={styles.menu}>
            <div className={styles.menuHeader}>
              <div className={styles.menuTitle}>Chat settings</div>
              <div className={styles.menuHint}>Tune local UI behavior, choose a theme, and set the QA subagent model.</div>
            </div>
            <div className={styles.sectionLabel}>Appearance</div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Theme</div>
                {selectedTheme && <div className={styles.valueText}>{selectedTheme.label}</div>}
              </div>
              <Combobox
                className={styles.dropdown}
                appearance="filled-darker"
                placeholder="Select theme"
                aria-label="Theme"
                value={selectedTheme?.label ?? ""}
                onOptionSelect={(_, data) => {
                  const nextThemeId = data.optionValue;
                  if (nextThemeId && nextThemeId !== selectedThemeId) onThemeChange(nextThemeId);
                }}
              >
                {themes.map((theme) => (
                  <Option key={theme.id} value={theme.id} text={theme.label}>
                    {theme.label}
                    {theme.isDefault ? " (Default)" : ""}
                  </Option>
                ))}
              </Combobox>
            </div>
            <div className={styles.sectionLabel}>Display</div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Thinking</div>
                <Switch aria-label="Show Thinking" checked={showThinking} onChange={(_, data) => onShowThinkingChange(data.checked)} />
              </div>
            </div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Raw Tool Responses in Expand</div>
                <Switch aria-label="Show Raw Tool Responses in Expand" checked={showToolResponses} onChange={(_, data) => onShowToolResponsesChange(data.checked)} />
              </div>
            </div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Debug Events</div>
                <Switch aria-label="Show Debug Events" checked={debugEnabled} onChange={(_, data) => onDebugChange(data.checked)} />
              </div>
            </div>
            <div className={styles.sectionLabel}>Models</div>
            <div className={styles.item}>
              <div className={styles.label}>QA Subagent Model</div>
              <Combobox
                className={styles.dropdown}
                appearance="filled-darker"
                freeform
                placeholder="Same as primary model"
                aria-label="QA Subagent Model"
                value={qaValue}
                onChange={(event) => setQaValue((event.target as HTMLInputElement).value)}
                onOpenChange={(_, data) => {
                  setQaOpen(data.open);
                  setQaValue(data.open ? "" : qaLabel);
                }}
                onOptionSelect={(_, data) => {
                  const nextModel = ModelTypeSchema.catch("").parse(data.optionValue);
                  onQaSubagentModelChange(nextModel);
                  setQaOpen(false);
                  setQaValue(data.optionText || qaLabel);
                }}
              >
                <Option value="" text="Same as primary model">
                  Same as primary model
                </Option>
                {qaItems.map((model) => (
                  <Option key={model.key} value={model.key} text={model.label}>
                    {model.label}
                  </Option>
                ))}
              </Combobox>
              <div className={styles.help}>Uses the selected model list from OpenCode. Leaving this on the default follows the active chat model.</div>
            </div>
            <div className={styles.menuHint}>UI options are kept locally in the add-in. The QA subagent model is stored in OpenCode config so it persists across sessions.</div>
            </PopoverSurface>
          </Popover>
          <Tooltip content="History" relationship="label">
            <Button
              icon={<History24Regular />}
              appearance="subtle"
              onClick={onShowHistory}
              aria-label="History"
              className={styles.icon}
            />
          </Tooltip>
          <Tooltip content="New chat" relationship="label">
            <Button
              icon={<Compose24Regular />}
              onClick={onNewChat}
              aria-label="New chat"
              className={styles.primary}
            />
          </Tooltip>
        </div>
      </div>

      {(connectionStatus || subtitle || contextLabel) && (
        <div className={styles.bottomRow}>
          {connectionStatus && (
            <div className={`${styles.statusPill} ${statusClassName ?? ""}`.trim()}>
              <span className={styles.statusDot} />
              <span>{connectionStatus.label}</span>
            </div>
          )}
          {subtitle && <div className={styles.subtitle}>{subtitle}</div>}
          {contextLabel && <div className={styles.contextChip}>{contextLabel}</div>}
        </div>
      )}
    </div>
  );
};
