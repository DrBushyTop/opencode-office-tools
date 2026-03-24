import * as React from "react";
import {
  Button,
  Switch,
  makeStyles,
  tokens,
  Tooltip,
} from "@fluentui/react-components";
import {
  Dismiss16Regular,
  ChevronLeft16Regular,
  Folder16Regular,
} from "@fluentui/react-icons";
import type { PermissionRule } from "../lib/permissionService";

interface PermissionManagerProps {
  cwd: string | null;
  onCwdChange: (cwd: string) => void;
  rules: PermissionRule[];
  onRemoveRule: (index: number) => void;
  onClearRules: () => void;
  allowAll: boolean;
  onAllowAllChange: (v: boolean) => void;
  onClose: () => void;
}

const KIND_COLORS: Record<string, string> = {
  shell: "#d13438",
  write: "#ca5010",
  read: "#0078d4",
  mcp: "#8764b8",
};

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  title: {
    fontWeight: 600,
    fontSize: "14px",
  },
  body: {
    flex: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionTitle: {
    fontSize: "12px",
    fontWeight: 600,
    textTransform: "uppercase",
    color: tokens.colorNeutralForeground3,
    letterSpacing: "0.5px",
  },
  cwdRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: "13px",
  },
  cwdPath: {
    fontFamily: "monospace",
    fontSize: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "4px 8px",
    borderRadius: "4px",
    flex: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  ruleRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 0",
    fontSize: "12px",
  },
  kindBadge: {
    fontSize: "11px",
    fontWeight: 600,
    padding: "1px 6px",
    borderRadius: "3px",
    color: "white",
    textTransform: "uppercase",
    minWidth: "42px",
    textAlign: "center",
  },
  rulePath: {
    flex: 1,
    fontFamily: "monospace",
    fontSize: "11px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  emptyText: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  folderBrowser: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "4px",
    overflow: "hidden",
  },
  folderHeader: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    padding: "6px 8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: "11px",
    fontFamily: "monospace",
    overflow: "hidden",
  },
  folderHeaderPath: {
    flex: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  folderList: {
    maxHeight: "180px",
    overflowY: "auto",
    padding: "2px 0",
  },
  folderItem: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px 8px",
    fontSize: "12px",
    cursor: "pointer",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  folderItemName: {
    flex: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  folderActions: {
    display: "flex",
    gap: "8px",
    padding: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    justifyContent: "flex-end",
  },
  allowAllRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  allowAllLabel: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  allowAllTitle: {
    fontSize: "13px",
    fontWeight: 500,
  },
  allowAllDesc: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
});

export const PermissionManager: React.FC<PermissionManagerProps> = ({
  cwd,
  onCwdChange,
  rules,
  onRemoveRule,
  onClearRules,
  allowAll,
  onAllowAllChange,
  onClose,
}) => {
  const styles = useStyles();
  const [browsing, setBrowsing] = React.useState(false);
  const [browseDir, setBrowseDir] = React.useState<string>("");
  const [browseDirs, setBrowseDirs] = React.useState<string[]>([]);
  const [browseParent, setBrowseParent] = React.useState<string | null>(null);
  const [browseLoading, setBrowseLoading] = React.useState(false);

  const loadDir = React.useCallback(async (dir?: string) => {
    setBrowseLoading(true);
    try {
      const param = dir ? `?path=${encodeURIComponent(dir)}` : "";
      const r = await fetch(`/api/browse${param}`);
      const data = await r.json();
      if (data.error) return;
      setBrowseDir(data.path);
      setBrowseDirs(data.dirs || []);
      setBrowseParent(data.parent);
    } catch { /* ignore */ }
    finally { setBrowseLoading(false); }
  }, []);

  const handleBrowse = React.useCallback(() => {
    setBrowsing(true);
    // Start from current cwd, or server env
    if (cwd) {
      loadDir(cwd);
    } else {
      fetch("/api/env").then(r => r.json()).then(data => {
        loadDir(data.cwd || data.home);
      }).catch(() => loadDir());
    }
  }, [cwd, loadDir]);

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <span className={styles.title}>Permissions & Settings</span>
        <Button
          appearance="subtle"
          icon={<Dismiss16Regular />}
          onClick={onClose}
          aria-label="Close"
        />
      </div>

      <div className={styles.body}>
        {/* Working Directory */}
        <div className={styles.section}>
          <span className={styles.sectionTitle}>Working Directory</span>
          <div className={styles.cwdRow}>
            <span>{cwd ? "📂" : "⚠️"}</span>
            {cwd ? (
              <Tooltip content={cwd} relationship="label">
                <span className={styles.cwdPath}>{cwd}</span>
              </Tooltip>
            ) : (
              <span className={styles.emptyText}>Not set</span>
            )}
          </div>

          {!browsing ? (
            <Button size="small" onClick={handleBrowse}>Browse…</Button>
          ) : (
            <div className={styles.folderBrowser}>
              <div className={styles.folderHeader}>
                {browseParent && (
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ChevronLeft16Regular />}
                    onClick={() => loadDir(browseParent!)}
                    aria-label="Go up"
                    style={{ minWidth: "24px", padding: 0 }}
                  />
                )}
                <span className={styles.folderHeaderPath}>{browseDir}</span>
              </div>
              <div className={styles.folderList}>
                {browseLoading ? (
                  <div style={{ padding: "8px", fontSize: "12px", color: tokens.colorNeutralForeground3 }}>Loading…</div>
                ) : browseDirs.length === 0 ? (
                  <div style={{ padding: "8px", fontSize: "12px", color: tokens.colorNeutralForeground3, fontStyle: "italic" }}>No subdirectories</div>
                ) : (
                  browseDirs.map((name) => (
                    <div
                      key={name}
                      className={styles.folderItem}
                      onClick={() => loadDir(browseDir + "/" + name)}
                    >
                      <Folder16Regular />
                      <span className={styles.folderItemName}>{name}</span>
                    </div>
                  ))
                )}
              </div>
              <div className={styles.folderActions}>
                <Button size="small" appearance="subtle" onClick={() => setBrowsing(false)}>
                  Cancel
                </Button>
                <Button
                  size="small"
                  appearance="primary"
                  onClick={() => {
                    onCwdChange(browseDir);
                    setBrowsing(false);
                  }}
                >
                  Select "{browseDir.split("/").pop() || browseDir}"
                </Button>
              </div>
            </div>
          )}
        </div>

        {/* Allow All Toggle */}
        <div className={styles.section}>
          <span className={styles.sectionTitle}>Trust Mode</span>
          <div className={styles.allowAllRow}>
            <div className={styles.allowAllLabel}>
              <span className={styles.allowAllTitle}>Allow All</span>
              <span className={styles.allowAllDesc}>
                Auto-approve all operations under the working directory
              </span>
            </div>
            <Switch
              checked={allowAll}
              onChange={(_, data) => onAllowAllChange(data.checked)}
            />
          </div>
        </div>

        {/* Permission Rules */}
        <div className={styles.section}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span className={styles.sectionTitle}>Saved Rules</span>
            {rules.length > 0 && (
              <Button size="small" appearance="subtle" onClick={onClearRules}>
                Clear all
              </Button>
            )}
          </div>
          {rules.length === 0 ? (
            <span className={styles.emptyText}>
              No saved rules. Click "Always Allow" on a permission prompt to save one.
            </span>
          ) : (
            rules.map((rule, i) => (
              <div key={`${rule.kind}-${rule.pathPrefix}-${i}`} className={styles.ruleRow}>
                <span
                  className={styles.kindBadge}
                  style={{ backgroundColor: KIND_COLORS[rule.kind] || "#666" }}
                >
                  {rule.kind}
                </span>
                <Tooltip content={rule.pathPrefix} relationship="label">
                  <span className={styles.rulePath}>{rule.pathPrefix}</span>
                </Tooltip>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<Dismiss16Regular />}
                  onClick={() => onRemoveRule(i)}
                  aria-label="Remove rule"
                />
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
};
