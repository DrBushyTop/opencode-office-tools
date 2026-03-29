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
import { filterModels } from "../lib/model-search";
import type { ModelInfo } from "../lib/opencode-client";

export type ModelType = string;

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
  subtitle?: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 12px",
    gap: "10px",
    minHeight: "52px",
    borderBottom: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
  },
  left: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
    flex: 1,
  },
  title: {
    fontSize: "12px",
    fontWeight: "600",
    color: "var(--oc-text)",
  },
  subtitle: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  dropdown: {
    minWidth: "160px",
    fontSize: "12px",
    borderRadius: "8px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border) !important",
    padding: "0 8px",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
    },
  },
  group: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    flexShrink: 0,
  },
  icon: {
    minWidth: "30px",
    width: "30px",
    height: "30px",
    padding: "0",
    borderRadius: "8px",
    color: "var(--oc-text-muted)",
    background: "transparent",
    border: "1px solid transparent",
    ":hover": {
      background: "var(--oc-bg-soft)",
      color: "var(--oc-text)",
      border: "1px solid var(--oc-border)",
    },
  },
  primary: {
    backgroundColor: "var(--oc-accent)",
    color: "white",
    borderRadius: "8px",
    padding: "4px",
    width: "30px",
    height: "30px",
    minWidth: "30px",
    border: "none",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ":hover": {
      backgroundColor: "var(--oc-accent-strong)",
    },
  },
  menu: {
    width: "280px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    background: "var(--oc-bg, #1b1818)",
    backgroundColor: "var(--oc-bg, #1b1818)",
    color: "var(--oc-text, #f1ecec)",
    border: "1px solid var(--oc-border, rgba(255,255,255,0.10))",
    borderRadius: "12px",
    boxShadow: "var(--oc-shadow, 0 0 0 1px rgba(255,255,255,0.06), 0 16px 48px rgba(0,0,0,0.24))",
    position: "relative",
    zIndex: 20,
  },
  menuTitle: {
    fontSize: "12px",
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
    gap: "6px",
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
  help: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    lineHeight: "1.5",
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
  subtitle,
}) => {
  const styles = useStyles();
  const [value, setValue] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const [qaValue, setQaValue] = React.useState("");
  const [qaOpen, setQaOpen] = React.useState(false);
  const selectedLabel = label(models, selectedModel);
  const qaLabel = label(models, qaSubagentModel);
  const items = React.useMemo(() => filterModels(models, open ? value : ""), [models, open, value]);
  const qaItems = React.useMemo(() => filterModels(models, qaOpen ? qaValue : ""), [models, qaOpen, qaValue]);

  React.useEffect(() => {
    if (!open) setValue(selectedLabel);
  }, [open, selectedLabel]);

  React.useEffect(() => {
    if (!qaOpen) setQaValue(qaLabel);
  }, [qaLabel, qaOpen]);

  return (
    <div className={styles.header}>
      <div className={styles.left}>
        <div className={styles.title}>OpenCode</div>
        <Combobox
          className={styles.dropdown}
          appearance="filled-darker"
          freeform
          placeholder="Search models"
          value={value}
          onChange={(event) => setValue((event.target as HTMLInputElement).value)}
          onOpenChange={(_, data) => {
            setOpen(data.open);
            setValue(data.open ? "" : selectedLabel);
          }}
          onOptionSelect={(_, data) => {
            if (data.optionValue && data.optionValue !== selectedModel) onModelChange(data.optionValue as ModelType);
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
        {subtitle && <div className={styles.subtitle}>{subtitle}</div>}
      </div>

      <div className={styles.group}>
        <Popover positioning="below-end" inline>
          <PopoverTrigger disableButtonEnhancement>
            <Button icon={<Settings24Regular />} appearance="subtle" aria-label="Options" className={styles.icon} />
          </PopoverTrigger>
          <PopoverSurface className={styles.menu}>
            <div className={styles.menuTitle}>Options</div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Thinking</div>
                <Switch checked={showThinking} onChange={(_, data) => onShowThinkingChange(data.checked)} />
              </div>
            </div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Raw Tool Responses in Expand</div>
                <Switch checked={showToolResponses} onChange={(_, data) => onShowToolResponsesChange(data.checked)} />
              </div>
            </div>
            <div className={styles.item}>
              <div className={styles.row}>
                <div className={styles.label}>Show Debug Events</div>
                <Switch checked={debugEnabled} onChange={(_, data) => onDebugChange(data.checked)} />
              </div>
            </div>
            <div className={styles.item}>
              <div className={styles.label}>QA Subagent Model</div>
              <Combobox
                className={styles.dropdown}
                appearance="filled-darker"
                freeform
                placeholder="Same as primary model"
                value={qaValue}
                onChange={(event) => setQaValue((event.target as HTMLInputElement).value)}
                onOpenChange={(_, data) => {
                  setQaOpen(data.open);
                  setQaValue(data.open ? "" : qaLabel);
                }}
                onOptionSelect={(_, data) => {
                  onQaSubagentModelChange((data.optionValue as ModelType) || "");
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
  );
};
