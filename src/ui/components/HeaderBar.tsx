import * as React from "react";
import { Button, Tooltip, Switch, makeStyles, Combobox, Option } from "@fluentui/react-components";
import { Compose24Regular, History24Regular } from "@fluentui/react-icons";
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
  onDebugChange: (v: boolean) => void;
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
  leftSection: {
    display: "flex",
    flexDirection: "column",
    gap: "3px",
    minWidth: 0,
    flex: 1,
  },
  titleRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
  },
  title: {
    fontSize: "12px",
    fontWeight: "600",
    color: "var(--oc-text)",
  },
  debugRow: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: "var(--oc-text-muted)",
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
  buttonGroup: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    flexShrink: 0,
  },
  iconButton: {
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
  clearButton: {
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
});

export const HeaderBar: React.FC<HeaderBarProps> = ({
  onNewChat,
  onShowHistory,
  selectedModel,
  onModelChange,
  models,
  debugEnabled,
  onDebugChange,
  subtitle,
}) => {
  const styles = useStyles();
  const [value, setValue] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const selectedLabel = models.find(m => m.key === selectedModel)?.label || selectedModel;
  const items = React.useMemo(() => filterModels(models, open ? value : ""), [models, open, value]);

  React.useEffect(() => {
    if (open) return;
    setValue(selectedLabel);
  }, [open, selectedModel]);

  return (
    <div className={styles.header}>
      <div className={styles.leftSection}>
        <div className={styles.titleRow}>
          <div className={styles.title}>OpenCode</div>
        </div>
        <Combobox
          className={styles.dropdown}
          appearance="filled-darker"
          freeform
          placeholder="Search models"
          value={value}
          onChange={(event) => {
            setValue((event.target as HTMLInputElement).value);
          }}
          onOpenChange={(_, data) => {
            setOpen(data.open);
            setValue(data.open ? "" : selectedLabel);
          }}
          onOptionSelect={(_, data) => {
            if (data.optionValue && data.optionValue !== selectedModel) {
              onModelChange(data.optionValue as ModelType);
            }
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
        {/* Debug toggle — hidden by default, enable via localStorage: opencode-debug-visible=true */}
        {localStorage.getItem("opencode-debug-visible") === "true" && (
          <div className={styles.debugRow}>
            <Switch
              checked={debugEnabled}
              onChange={(_, data) => onDebugChange(data.checked)}
              label="Debug"
              style={{ fontSize: "11px" }}
            />
          </div>
        )}
        {subtitle && <div className={styles.subtitle}>{subtitle}</div>}
      </div>
      <div className={styles.buttonGroup}>
        <Tooltip content="History" relationship="label">
          <Button
            icon={<History24Regular />}
            appearance="subtle"
            onClick={onShowHistory}
            aria-label="History"
            className={styles.iconButton}
          />
        </Tooltip>
        <Tooltip content="New chat" relationship="label">
          <Button
            icon={<Compose24Regular />}
            onClick={onNewChat}
            aria-label="New chat"
            className={styles.clearButton}
          />
        </Tooltip>
      </div>
    </div>
  );
};
