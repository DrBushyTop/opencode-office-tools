import * as React from "react";
import { Button, makeStyles, tokens } from "@fluentui/react-components";
import type { OfficePermissionRequest } from "../../shared/office-permissions";
import { isReadOnlyOfficeTool } from "../../shared/office-tool-registry";

export type PermissionDecision = "allow" | "deny" | "always";

interface PermissionDialogProps {
  request: OfficePermissionRequest;
  cwd: string | null;
  onDecision: (decision: PermissionDecision) => void;
}

const KIND_META: Record<string, { icon: string; label: string; color: string }> = {
  shell: { icon: "⚡", label: "Run Shell Command", color: "#d13438" },
  write: { icon: "✏️", label: "Write File", color: "#ca5010" },
  read: { icon: "📖", label: "Read File", color: "#0078d4" },
  mcp: { icon: "🔌", label: "MCP Server Call", color: "#8764b8" },
};

const useStyles = makeStyles({
  overlay: {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(0,0,0,0.4)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 9999,
    padding: "16px",
  },
  dialog: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: "8px",
    maxWidth: "420px",
    width: "100%",
    boxShadow: tokens.shadow16,
    overflow: "hidden",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  icon: {
    fontSize: "20px",
  },
  kindLabel: {
    fontWeight: 600,
    fontSize: "13px",
  },
  intention: {
    padding: "8px 16px",
    fontSize: "13px",
    color: tokens.colorNeutralForeground2,
  },
  details: {
    padding: "0 16px 12px",
    maxHeight: "200px",
    overflowY: "auto",
  },
  codeBlock: {
    fontFamily: "monospace",
    fontSize: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "8px",
    borderRadius: "4px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-all",
    lineHeight: "1.4",
  },
  actions: {
    display: "flex",
    gap: "8px",
    padding: "12px 16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    justifyContent: "flex-end",
  },
  denyBtn: {
    color: "#d13438",
  },
  allowBtn: {
    backgroundColor: "#107c10",
    color: "white",
    ":hover": { backgroundColor: "#0b6a0b" },
  },
  alwaysBtn: {
    fontWeight: 600,
  },
});

function getDetail(request: OfficePermissionRequest): string {
  return JSON.stringify({
    permission: request.permission,
    tool: request.metadata.tool || request.patterns[0],
    input: request.metadata.input,
    patterns: request.patterns,
  }, null, 2);
}

export const PermissionDialog: React.FC<PermissionDialogProps> = ({
  request,
  onDecision,
}) => {
  const styles = useStyles();
  const requestedTool = String(request.metadata.tool || request.patterns[0] || "");
  const meta = request.permission === "doom_loop"
    ? { icon: "🛑", label: "Repeated Tool Call", color: "#d13438" }
    : isReadOnlyOfficeTool(requestedTool)
      ? KIND_META.read
      : KIND_META.write;
  const detail = getDetail(request);

  return (
    <div className={styles.overlay}>
      <div className={styles.dialog}>
        <div className={styles.header}>
          <span className={styles.icon}>{meta.icon}</span>
          <span className={styles.kindLabel} style={{ color: meta.color }}>
            {meta.label}
          </span>
        </div>

        <div className={styles.intention}>
          {request.permission === "doom_loop"
            ? "OpenCode wants confirmation before repeating the same tool call again."
            : `OpenCode wants permission to use ${String(request.metadata.tool || request.patterns[0] || "this tool")}.`}
        </div>

        {detail && (
          <div className={styles.details}>
            <div className={styles.codeBlock}>{detail}</div>
          </div>
        )}

        <div className={styles.actions}>
          <Button
            appearance="subtle"
            className={styles.denyBtn}
            onClick={() => onDecision("deny")}
          >
            Deny
          </Button>
          <Button
            appearance="primary"
            className={styles.allowBtn}
            onClick={() => onDecision("allow")}
          >
            Allow
          </Button>
          <Button
            appearance="outline"
            className={styles.alwaysBtn}
            style={{ borderColor: meta.color, color: meta.color }}
            onClick={() => onDecision("always")}
          >
            Always Allow
          </Button>
        </div>
      </div>
    </div>
  );
};
