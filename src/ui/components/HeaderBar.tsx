import * as React from "react";
import { Button, Tooltip, Switch, makeStyles, Combobox, Option, tokens } from "@fluentui/react-components";
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
    padding: "8px 12px",
    paddingRight: "40px",
    gap: "8px",
    minHeight: "40px",
  },
  leftSection: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    minWidth: 0,
    flex: 1,
  },
  debugRow: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  subtitle: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  dropdown: {
    minWidth: "120px",
    opacity: 0.6,
    fontSize: "12px",
    borderBottom: "none",
    ":hover": {
      opacity: 1,
    },
  },
  buttonGroup: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flexShrink: 0,
  },
  iconButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
  },
  clearButton: {
    backgroundColor: "#0078d4",
    color: "white",
    borderRadius: "4px",
    padding: "4px",
    width: "28px",
    height: "28px",
    minWidth: "28px",
    border: "none",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ":hover": {
      backgroundColor: "#106ebe",
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
        <Combobox
          className={styles.dropdown}
          appearance="underline"
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
