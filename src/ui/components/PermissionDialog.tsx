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

/* ------------------------------------------------------------------ */
/*  Visual metadata per permission kind                               */
/* ------------------------------------------------------------------ */

const KIND_BADGES: Record<string, { glyph: string; title: string; accent: string }> = {
  shell: { glyph: "⚡", title: "Run Shell Command", accent: "#d13438" },
  write: { glyph: "✏️", title: "Write File", accent: "#ca5010" },
  read: { glyph: "📖", title: "Read File", accent: "#0078d4" },
  subagent: { glyph: "🧠", title: "Launch Subagent", accent: "#5c2dbd" },
  mcp: { glyph: "🔌", title: "MCP Server Call", accent: "#8764b8" },
  generic: { glyph: "🔐", title: "Permission Request", accent: "#4f6bed" },
  danger: { glyph: "🛑", title: "Repeated Tool Call", accent: "#d13438" },
};

/* ------------------------------------------------------------------ */
/*  Styles                                                            */
/* ------------------------------------------------------------------ */

const useStyles = makeStyles({
  backdrop: {
    position: "fixed",
    inset: "0",
    backgroundColor: "rgba(8, 8, 8, 0.56)",
    backdropFilter: "blur(10px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 9999,
    padding: "16px",
  },
  card: {
    background: "var(--oc-bg)",
    color: "var(--oc-text)",
    borderRadius: "16px",
    maxWidth: "460px",
    width: "100%",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
  },
  titleRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "16px 18px 12px",
    borderBottom: "1px solid var(--oc-border)",
  },
  badge: {
    width: "32px",
    height: "32px",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "16px",
    borderRadius: "10px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    flexShrink: 0,
  },
  titleText: {
    fontWeight: 700,
    fontSize: "13px",
    letterSpacing: "0.01em",
  },
  prompt: {
    padding: "14px 18px 10px",
    fontSize: "13px",
    color: "var(--oc-text-muted)",
    lineHeight: "1.6",
  },
  detailWrap: {
    padding: "0 18px 16px",
    maxHeight: "220px",
    overflowY: "auto",
  },
  mono: {
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
  buttonBar: {
    display: "flex",
    gap: "8px",
    padding: "14px 18px 18px",
    borderTop: "1px solid var(--oc-border)",
    justifyContent: "flex-end",
    background: "var(--oc-bg)",
  },
  btnDeny: {
    color: "var(--oc-danger-text)",
    borderRadius: "10px",
    background: "transparent",
    border: "1px solid transparent",
    ":hover": {
      background: "var(--oc-danger-bg)",
      border: "1px solid var(--oc-danger-border)",
    },
  },
  btnAllow: {
    backgroundColor: "var(--oc-accent)",
    color: "var(--text-on-interactive-base, #fcfcfc)",
    borderRadius: "10px",
    border: "none",
    ":hover": { backgroundColor: "var(--oc-accent-strong)" },
  },
  btnPersist: {
    fontWeight: 600,
    borderRadius: "10px",
    background: "transparent",
  },
});

/* ------------------------------------------------------------------ */
/*  Helpers                                                           */
/* ------------------------------------------------------------------ */

function buildDetailJson(req: OfficePermissionRequest): string {
  const safe = OfficePermissionRequestSchema.catch(req).parse(req);
  return JSON.stringify(
    {
      permission: safe.permission,
      target: permissionTarget(safe),
      input: safe.metadata.input,
      patterns: safe.patterns,
      metadata: safe.metadata,
    },
    null,
    2,
  );
}

function intentionCopy(req: OfficePermissionRequest, target: string | null): string {
  if (req.permission === "doom_loop") {
    return "Confirmation needed — the same tool call is about to repeat.";
  }
  return `Requesting access to ${target || "a tool"}.`;
}

/* ------------------------------------------------------------------ */
/*  Component                                                         */
/* ------------------------------------------------------------------ */

export const PermissionDialog: React.FC<PermissionDialogProps> = ({
  request,
  sessionTitle,
  onDecision,
}) => {
  const cls = useStyles();
  const safe = OfficePermissionRequestSchema.catch(request).parse(request);
  const target = permissionTarget(safe);
  const badge = KIND_BADGES[permissionKind(safe)] ?? KIND_BADGES.generic;
  const detailJson = buildDetailJson(safe);
  const caller = sessionTitle?.trim() || null;

  return (
    <div className={cls.backdrop}>
      <div className={cls.card} role="dialog" aria-label={badge.title}>
        {/* ---- title ---- */}
        <header className={cls.titleRow}>
          <span className={cls.badge} aria-hidden>{badge.glyph}</span>
          <span className={cls.titleText} style={{ color: badge.accent }}>
            {badge.title}
          </span>
        </header>

        {/* ---- description ---- */}
        <p className={cls.prompt}>
          {intentionCopy(safe, target)}
          {caller && <>{" "}(from <strong>{caller}</strong>)</>}
        </p>

        {/* ---- json detail ---- */}
        {detailJson && (
          <section className={cls.detailWrap}>
            <pre className={cls.mono}>{detailJson}</pre>
          </section>
        )}

        {/* ---- actions ---- */}
        <footer className={cls.buttonBar}>
          <Button
            appearance="subtle"
            className={cls.btnDeny}
            onClick={() => onDecision("deny")}
          >
            Deny
          </Button>
          <Button
            appearance="outline"
            className={cls.btnPersist}
            style={{ borderColor: badge.accent, color: badge.accent }}
            onClick={() => onDecision("always")}
          >
            Always Allow
          </Button>
          <Button
            appearance="primary"
            className={cls.btnAllow}
            onClick={() => onDecision("allow")}
          >
            Allow
          </Button>
        </footer>
      </div>
    </div>
  );
};
