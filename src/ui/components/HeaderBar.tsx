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
import { Compose20Regular, History20Regular, Settings20Regular } from "@fluentui/react-icons";
import { z } from "zod";
import { filterModels } from "../lib/model-search";
import type { ModelInfo } from "../lib/opencode-client";
import type { ThemePreference } from "../lib/ui-theme";

const ModelTypeSchema = z.string();
const ModelInfoSchema = z.object({
  key: z.string().min(1),
  label: z.string().min(1),
  providerID: z.string().min(1),
  modelID: z.string().min(1),
  limitContext: z.number().nonnegative().optional(),
  variants: z.array(z.string()).optional(),
}) satisfies z.ZodType<ModelInfo>;

export type ModelType = z.infer<typeof ModelTypeSchema>;

export interface HeaderThemeOption {
  id: string;
  label: string;
  isDefault?: boolean;
}

export interface HeaderThemeModeOption {
  id: ThemePreference;
  label: string;
}

export interface HeaderConnectionStatus {
  state: "connecting" | "connected" | "disconnected";
  label: string;
}

interface HeaderBarProps {
  onNewChat: () => void;
  onShowHistory: () => void;
  historyOpen: boolean;
  selectedModel: ModelType;
  onModelChange: (model: ModelType) => void;
  models: ModelInfo[];
  selectedVariant?: string;
  onVariantChange: (variant: string | undefined) => void;
  debugEnabled: boolean;
  onDebugChange: (value: boolean) => void;
  showThinking: boolean;
  onShowThinkingChange: (value: boolean) => void;
  showToolResponses: boolean;
  onShowToolResponsesChange: (value: boolean) => void;
  qaSubagentModel: ModelType;
  onQaSubagentModelChange: (model: ModelType) => void;
  qaSubagentVariant: string | undefined;
  onQaSubagentVariantChange: (variant: string | undefined) => void;
  themes: HeaderThemeOption[];
  selectedThemeId: string;
  onThemeChange: (themeId: string) => void;
  themeModes: HeaderThemeModeOption[];
  selectedThemeMode: ThemePreference;
  onThemeModeChange: (mode: ThemePreference) => void;
  connectionStatus?: HeaderConnectionStatus;
  subtitle?: string;
  contextLabel?: string;
  usageSummary?: string;
}

const useStyles = makeStyles({
  shell: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "14px 14px 10px",
    borderBottom: "1px solid var(--oc-border)",
    background: "linear-gradient(180deg, color-mix(in srgb, var(--oc-bg-soft) 76%, transparent), transparent)",
    flexShrink: 0,
  },
  primaryCard: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "14px",
    borderRadius: "18px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
    boxShadow: "var(--oc-shadow)",
  },
  topRow: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: "12px",
    flexWrap: "wrap",
  },
  identityCluster: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
    flex: "1 1 280px",
  },
  brandRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
    flexWrap: "wrap",
  },
  brandPill: {
    display: "inline-flex",
    alignItems: "center",
    padding: "4px 10px",
    borderRadius: "999px",
    background: "color-mix(in srgb, var(--oc-accent) 14%, transparent)",
    color: "var(--oc-accent)",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 34%, transparent)",
    fontSize: "11px",
    fontWeight: "700",
    letterSpacing: "0.06em",
    textTransform: "uppercase",
  },
  subtitle: {
    fontSize: "13px",
    lineHeight: "18px",
    color: "var(--oc-text-muted)",
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  contextLabel: {
    fontSize: "12px",
    color: "var(--oc-text-faint)",
    lineHeight: "16px",
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  actionCluster: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
    flex: "0 1 auto",
  },
  statusPill: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px 10px",
    borderRadius: "999px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
    color: "var(--oc-text-muted)",
    fontSize: "11px",
    fontWeight: "600",
  },
  statusDot: {
    width: "7px",
    height: "7px",
    borderRadius: "999px",
    background: "currentColor",
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
    padding: "4px 10px",
    borderRadius: "999px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
    whiteSpace: "nowrap",
  },
  icon: {
    minWidth: "32px",
    width: "32px",
    height: "32px",
    padding: "0",
    borderRadius: "10px",
    color: "var(--oc-text-muted)",
    background: "transparent",
    border: "1px solid transparent",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
      color: "var(--oc-text)",
      border: "1px solid var(--oc-border)",
    },
  },
  iconActive: {
    background: "var(--oc-bg-soft)",
    color: "var(--oc-text)",
    border: "1px solid var(--oc-border)",
  },
  primary: {
    backgroundColor: "var(--oc-accent)",
    color: "white",
    borderRadius: "10px",
    padding: "0",
    width: "32px",
    height: "32px",
    minWidth: "32px",
    border: "none",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  controlRow: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) 132px",
    gap: "10px",
    alignItems: "end",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "1fr",
    },
  },
  controlBlock: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
  },
  controlLabel: {
    fontSize: "11px",
    fontWeight: "700",
    letterSpacing: "0.06em",
    textTransform: "uppercase",
    color: "var(--oc-text-faint)",
  },
  control: {
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
  menu: {
    width: "340px",
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
  sectionLabel: {
    fontSize: "11px",
    fontWeight: "700",
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    color: "var(--oc-text-faint)",
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
  variantBar: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flexWrap: "wrap",
  },
  variantChip: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "2px 8px",
    borderRadius: "999px",
    fontSize: "11px",
    fontWeight: "500",
    cursor: "pointer",
    whiteSpace: "nowrap",
    lineHeight: "18px",
    border: "1px solid var(--oc-border)",
    background: "transparent",
    color: "var(--oc-text-muted)",
    transition: "all 0.12s ease",
    textTransform: "capitalize",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
      color: "var(--oc-text)",
    },
  },
  variantChipActive: {
    background: "color-mix(in srgb, var(--oc-accent) 15%, transparent)",
    border: "1px solid color-mix(in srgb, var(--oc-accent) 40%, transparent)",
    color: "var(--oc-accent)",
    fontWeight: "600",
  },
});

function labelForModel(models: ModelInfo[], model: ModelType) {
  if (!model) return "Choose a model";
  return models.find((item) => item.key === model)?.label || model;
}

export const HeaderBar: React.FC<HeaderBarProps> = ({
  onNewChat,
  onShowHistory,
  historyOpen,
  selectedModel,
  onModelChange,
  models,
  selectedVariant,
  onVariantChange,
  debugEnabled,
  onDebugChange,
  showThinking,
  onShowThinkingChange,
  showToolResponses,
  onShowToolResponsesChange,
  qaSubagentModel,
  onQaSubagentModelChange,
  qaSubagentVariant,
  onQaSubagentVariantChange,
  themes,
  selectedThemeId,
  onThemeChange,
  themeModes,
  selectedThemeMode,
  onThemeModeChange,
  connectionStatus,
  subtitle,
  contextLabel,
  usageSummary,
}) => {
  const styles = useStyles();
  const safeModels = React.useMemo(() => z.array(ModelInfoSchema).catch([]).parse(models), [models]);
  const [modelValue, setModelValue] = React.useState("");
  const [modelOpen, setModelOpen] = React.useState(false);
  const [qaValue, setQaValue] = React.useState("");
  const [qaOpen, setQaOpen] = React.useState(false);

  const selectedModelLabel = React.useMemo(
    () => labelForModel(safeModels, selectedModel),
    [safeModels, selectedModel],
  );
  const qaLabel = React.useMemo(
    () => (qaSubagentModel ? labelForModel(safeModels, qaSubagentModel) : "Same as primary model"),
    [qaSubagentModel, safeModels],
  );
  const modelItems = React.useMemo(
    () => filterModels(safeModels, modelOpen ? modelValue : ""),
    [modelOpen, modelValue, safeModels],
  );
  const qaItems = React.useMemo(
    () => filterModels(safeModels, qaOpen ? qaValue : ""),
    [qaOpen, qaValue, safeModels],
  );
  const modelVariants = React.useMemo(() => {
    const current = safeModels.find((model) => model.key === selectedModel);
    return current?.variants ?? [];
  }, [safeModels, selectedModel]);
  const qaModelVariants = React.useMemo(() => {
    const key = qaSubagentModel || selectedModel;
    const current = safeModels.find((model) => model.key === key);
    return current?.variants ?? [];
  }, [qaSubagentModel, safeModels, selectedModel]);
  const selectedTheme = React.useMemo(
    () => themes.find((theme) => theme.id === selectedThemeId) ?? themes.find((theme) => theme.isDefault) ?? themes[0],
    [selectedThemeId, themes],
  );
  const selectedThemeModeOption = React.useMemo(
    () => themeModes.find((mode) => mode.id === selectedThemeMode) ?? themeModes[0],
    [selectedThemeMode, themeModes],
  );
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
    if (!modelOpen) setModelValue(selectedModelLabel);
  }, [modelOpen, selectedModelLabel]);

  React.useEffect(() => {
    if (!qaOpen) setQaValue(qaLabel);
  }, [qaLabel, qaOpen]);

  return (
    <div className={styles.shell}>
      <div className={styles.primaryCard}>
        <div className={styles.topRow}>
          <div className={styles.identityCluster}>
            <div className={styles.brandRow}>
              <span className={styles.brandPill}>OpenCode</span>
              {subtitle && <span className={styles.subtitle}>{subtitle}</span>}
            </div>
            {contextLabel && <span className={styles.contextLabel}>{contextLabel}</span>}
          </div>
          <div className={styles.actionCluster}>
            {connectionStatus && (
              <div className={`${styles.statusPill} ${statusClassName ?? ""}`.trim()}>
                <span className={styles.statusDot} />
                <span>{connectionStatus.label}</span>
              </div>
            )}
            {usageSummary && <div className={styles.usage}>{usageSummary}</div>}
            <Popover positioning="below-end" inline>
              <PopoverTrigger disableButtonEnhancement>
                <Button icon={<Settings20Regular />} appearance="subtle" aria-label="Options" className={styles.icon} size="small" />
              </PopoverTrigger>
              <PopoverSurface className={styles.menu}>
                <div className={styles.menuHeader}>
                  <div className={styles.menuTitle}>Workspace settings</div>
                  <div className={styles.menuHint}>Adjust theme behavior, transcript detail, and QA defaults without leaving the current session.</div>
                </div>
                <div className={styles.sectionLabel}>Appearance</div>
                <div className={styles.item}>
                  <div className={styles.row}>
                    <div className={styles.label}>Theme palette</div>
                    {selectedTheme && <div className={styles.valueText}>{selectedTheme.label}</div>}
                  </div>
                  <Combobox
                    className={styles.control}
                    appearance="filled-darker"
                    aria-label="Theme palette"
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
                <div className={styles.item}>
                  <div className={styles.row}>
                    <div className={styles.label}>Color mode</div>
                    <div className={styles.valueText}>{selectedThemeModeOption?.label}</div>
                  </div>
                  <Combobox
                    className={styles.control}
                    appearance="filled-darker"
                    aria-label="Theme mode"
                    value={selectedThemeModeOption?.label ?? ""}
                    onOptionSelect={(_, data) => {
                      const nextMode = data.optionValue as ThemePreference | undefined;
                      if (nextMode && nextMode !== selectedThemeMode) onThemeModeChange(nextMode);
                    }}
                  >
                    {themeModes.map((mode) => (
                      <Option key={mode.id} value={mode.id} text={mode.label}>
                        {mode.label}
                      </Option>
                    ))}
                  </Combobox>
                </div>
                <div className={styles.sectionLabel}>Display</div>
                <div className={styles.item}>
                  <div className={styles.row}>
                    <div className={styles.label}>Show thinking</div>
                    <Switch aria-label="Show Thinking" checked={showThinking} onChange={(_, data) => onShowThinkingChange(data.checked)} />
                  </div>
                </div>
                <div className={styles.item}>
                  <div className={styles.row}>
                    <div className={styles.label}>Show raw tool responses in expand</div>
                    <Switch aria-label="Show Raw Tool Responses in Expand" checked={showToolResponses} onChange={(_, data) => onShowToolResponsesChange(data.checked)} />
                  </div>
                </div>
                <div className={styles.item}>
                  <div className={styles.row}>
                    <div className={styles.label}>Show debug events</div>
                    <Switch aria-label="Show Debug Events" checked={debugEnabled} onChange={(_, data) => onDebugChange(data.checked)} />
                  </div>
                </div>
                <div className={styles.sectionLabel}>QA Model</div>
                <div className={styles.item}>
                  <div className={styles.label}>Subagent model</div>
                  <Combobox
                    className={styles.control}
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
                  <div className={styles.help}>The QA subagent follows the active chat model unless you pin a different one here.</div>
                </div>
                {qaModelVariants.length > 0 && (
                  <div className={styles.item}>
                    <div className={styles.label}>QA reasoning effort</div>
                    <div className={styles.variantBar}>
                      <button
                        type="button"
                        className={`${styles.variantChip} ${qaSubagentVariant === undefined ? styles.variantChipActive : ""}`.trim()}
                        onClick={() => onQaSubagentVariantChange(undefined)}
                      >
                        default
                      </button>
                      {qaModelVariants.map((variant) => (
                        <button
                          type="button"
                          key={variant}
                          className={`${styles.variantChip} ${qaSubagentVariant === variant ? styles.variantChipActive : ""}`.trim()}
                          onClick={() => onQaSubagentVariantChange(variant)}
                        >
                          {variant}
                        </button>
                      ))}
                    </div>
                    <div className={styles.help}>Controls the reasoning effort level for visual QA runs.</div>
                  </div>
                )}
              </PopoverSurface>
            </Popover>
            <Tooltip content={historyOpen ? "Hide history" : "Show history"} relationship="label">
              <Button
                icon={<History20Regular />}
                appearance="subtle"
                onClick={onShowHistory}
                aria-label="History"
                className={`${styles.icon} ${historyOpen ? styles.iconActive : ""}`.trim()}
                size="small"
              />
            </Tooltip>
            <Tooltip content="New chat" relationship="label">
              <Button
                icon={<Compose20Regular />}
                onClick={onNewChat}
                aria-label="New chat"
                className={styles.primary}
                size="small"
              />
            </Tooltip>
          </div>
        </div>

        <div className={styles.controlRow}>
          <div className={styles.controlBlock}>
            <div className={styles.controlLabel}>Primary model</div>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              freeform
              placeholder="Search models"
              aria-label="Primary model"
              value={modelValue}
              onChange={(event) => setModelValue((event.target as HTMLInputElement).value)}
              onOpenChange={(_, data) => {
                setModelOpen(data.open);
                setModelValue(data.open ? "" : selectedModelLabel);
              }}
              onOptionSelect={(_, data) => {
                const nextModel = data.optionValue;
                if (nextModel && nextModel !== selectedModel) {
                  onModelChange(nextModel);
                }
                setModelOpen(false);
                setModelValue(data.optionText || selectedModelLabel);
              }}
            >
              {modelItems.map((model) => (
                <Option key={model.key} value={model.key} text={model.label}>
                  {model.label}
                </Option>
              ))}
            </Combobox>
          </div>
          <div className={styles.controlBlock}>
            <div className={styles.controlLabel}>Effort</div>
            <Combobox
              className={styles.control}
              appearance="filled-darker"
              placeholder="default"
              aria-label="Model effort"
              value={selectedVariant ?? "default"}
              disabled={modelVariants.length === 0}
              onOptionSelect={(_, data) => onVariantChange(data.optionValue || undefined)}
            >
              <Option value="" text="default">default</Option>
              {modelVariants.map((variant) => (
                <Option key={variant} value={variant} text={variant}>
                  {variant}
                </Option>
              ))}
            </Combobox>
          </div>
        </div>
      </div>
    </div>
  );
};
