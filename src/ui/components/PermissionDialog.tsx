import * as React from "react";
import { Button, makeStyles, tokens } from "@fluentui/react-components";
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
