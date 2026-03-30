import * as React from "react";
import { Button, makeStyles } from "@fluentui/react-components";
import { z } from "zod";
import type { OfficePermissionRequest } from "../../shared/office-permissions";
import { permissionKind, permissionTarget } from "../../shared/office-permissions";

const PermissionDecisionSchema = z.enum(["allow", "deny", "always"]);
const OfficePermissionRequestSchema = z.object({
  id: z.string().min(1),
  sessionID: z.string().min(1),
  permission: z.string().min(1),
  patterns: z.array(z.string()),
  metadata: z.record(z.string(), z.unknown()),
  always: z.array(z.string()),
  tool: z.object({
    messageID: z.string().min(1),
    callID: z.string().min(1),
  }).optional(),
}) satisfies z.ZodType<OfficePermissionRequest>;

export type PermissionDecision = z.infer<typeof PermissionDecisionSchema>;

interface PermissionDialogProps {
  request: OfficePermissionRequest;
  cwd: string | null;
  sessionTitle?: string | null;
  onDecision: (decision: PermissionDecision) => void;
}

const KIND_META: Record<string, { icon: string; label: string; color: string }> = {
  shell: { icon: "⚡", label: "Run Shell Command", color: "#d13438" },
  write: { icon: "✏️", label: "Write File", color: "#ca5010" },
  read: { icon: "📖", label: "Read File", color: "#0078d4" },
  subagent: { icon: "🧠", label: "Launch Subagent", color: "#5c2dbd" },
  mcp: { icon: "🔌", label: "MCP Server Call", color: "#8764b8" },
  generic: { icon: "🔐", label: "Permission Request", color: "#4f6bed" },
  danger: { icon: "🛑", label: "Repeated Tool Call", color: "#d13438" },
};

const useStyles = makeStyles({
  overlay: {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(8, 8, 8, 0.56)",
    backdropFilter: "blur(10px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 9999,
    padding: "16px",
  },
  dialog: {
    background: "var(--oc-bg)",
    color: "var(--oc-text)",
    borderRadius: "16px",
    maxWidth: "460px",
    width: "100%",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "16px 18px 12px",
    borderBottom: "1px solid var(--oc-border)",
  },
  icon: {
    width: "32px",
    height: "32px",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "16px",
    borderRadius: "10px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
  },
  kindLabel: {
    fontWeight: 700,
    fontSize: "13px",
    letterSpacing: "0.01em",
  },
  intention: {
    padding: "14px 18px 10px",
    fontSize: "13px",
    color: "var(--oc-text-muted)",
    lineHeight: "1.6",
  },
  details: {
    padding: "0 18px 16px",
    maxHeight: "220px",
    overflowY: "auto",
  },
  codeBlock: {
    fontFamily: '"SFMono-Regular", "Consolas", "Menlo", monospace',
    fontSize: "12px",
    background: "var(--oc-bg-strong)",
    color: "var(--oc-text-muted)",
    padding: "12px",
    borderRadius: "12px",
    border: "1px solid var(--oc-border)",
    whiteSpace: "pre-wrap",
    wordBreak: "break-all",
    lineHeight: "1.5",
  },
  actions: {
    display: "flex",
    gap: "8px",
    padding: "14px 18px 18px",
    borderTop: "1px solid var(--oc-border)",
    justifyContent: "flex-end",
    background: "var(--oc-bg)",
  },
  denyBtn: {
    color: "var(--oc-danger-text)",
    borderRadius: "10px",
    background: "transparent",
    border: "1px solid transparent",
    ":hover": {
      background: "var(--oc-danger-bg)",
      border: "1px solid var(--oc-danger-border)",
    },
  },
  allowBtn: {
    backgroundColor: "var(--oc-accent)",
    color: "var(--text-on-interactive-base, #fcfcfc)",
    borderRadius: "10px",
    border: "none",
    ":hover": { backgroundColor: "var(--oc-accent-strong)" },
  },
  alwaysBtn: {
    fontWeight: 600,
    borderRadius: "10px",
    background: "transparent",
  },
});

function getDetail(request: OfficePermissionRequest): string {
  const safeRequest = OfficePermissionRequestSchema.catch(request).parse(request);
  return JSON.stringify({
    permission: safeRequest.permission,
    target: permissionTarget(safeRequest),
    input: safeRequest.metadata.input,
    patterns: safeRequest.patterns,
    metadata: safeRequest.metadata,
  }, null, 2);
}

export const PermissionDialog: React.FC<PermissionDialogProps> = ({
  request,
  sessionTitle,
  onDecision,
}) => {
  const styles = useStyles();
  const safeRequest = OfficePermissionRequestSchema.catch(request).parse(request);
  const target = permissionTarget(safeRequest);
  const meta = KIND_META[permissionKind(safeRequest)] || KIND_META.generic;
  const detail = getDetail(safeRequest);
  const requester = sessionTitle && sessionTitle.trim() ? sessionTitle.trim() : null;

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
          {safeRequest.permission === "doom_loop"
            ? "OpenCode wants confirmation before repeating the same tool call again."
            : `OpenCode wants permission to use ${target || "this tool"}.`}
          {requester ? ` Requested by ${requester}.` : ""}
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
