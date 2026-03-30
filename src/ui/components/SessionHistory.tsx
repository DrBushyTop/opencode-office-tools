import * as React from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";
import { Delete24Regular, ArrowLeft24Regular } from "@fluentui/react-icons";
import { z } from "zod";
import type { OfficeHost } from "../sessionStorage";
import { deleteSession, listSessions, type OpencodeSessionInfo } from "../lib/opencode-session-history";

const SessionInfoSchema = z.object({
  id: z.string().min(1),
  title: z.string(),
  directory: z.string(),
  time: z.object({
    created: z.number(),
    updated: z.number(),
  }),
}) satisfies z.ZodType<OpencodeSessionInfo>;

interface SessionHistoryProps {
  host: OfficeHost;
  shared: boolean;
  onSharedChange: (shared: boolean) => void;
  onSelectSession: (session: OpencodeSessionInfo) => void;
  onClose: () => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    background: "var(--oc-bg)",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "14px 14px 12px",
    borderBottom: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
  },
  headerTitle: {
    fontWeight: "700",
    fontSize: "13px",
    color: "var(--oc-text)",
    letterSpacing: "0.01em",
  },
  filterRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "10px",
    padding: "12px 14px",
    borderBottom: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
  },
  filterLabel: {
    fontSize: "11px",
    fontWeight: "700",
    color: "var(--oc-text-faint)",
    textTransform: "uppercase",
    letterSpacing: "0.08em",
  },
  filterGroup: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px",
    borderRadius: "999px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
  },
  filterButton: {
    minWidth: "76px",
    padding: "0 12px",
    fontSize: "12px",
    borderRadius: "999px",
    height: "28px",
  },
  backButton: {
    minWidth: "30px",
    width: "30px",
    height: "30px",
    padding: "4px",
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
  list: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    padding: "10px",
    scrollbarColor: "var(--oc-text-faint) transparent",
    scrollbarWidth: "thin",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      backgroundColor: "transparent",
    },
    "&::-webkit-scrollbar-thumb": {
      backgroundColor: "var(--oc-text-faint)",
      borderRadius: "999px",
      border: "2px solid transparent",
      backgroundClip: "content-box",
      opacity: 0.45,
    },
    "&::-webkit-scrollbar-thumb:hover": {
      backgroundColor: "var(--oc-text-muted)",
    },
  },
  emptyState: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    color: "var(--oc-text-faint)",
    fontSize: "13px",
    lineHeight: "1.6",
    padding: "24px",
    textAlign: "center",
  },
  sessionItem: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "12px",
    borderRadius: "12px",
    cursor: "pointer",
    marginBottom: "8px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    transition: "transform 0.15s ease, border-color 0.15s ease, background 0.15s ease, box-shadow 0.15s ease",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
      border: "1px solid var(--oc-border-strong)",
      transform: "translateY(-1px)",
      boxShadow: "0 10px 24px rgba(0,0,0,0.08)",
    },
  },
  sessionContent: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
  },
  sessionTitle: {
    fontSize: "13px",
    fontWeight: "600",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    color: "var(--oc-text)",
  },
  sessionMeta: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    marginTop: "4px",
    display: "flex",
    gap: "8px",
  },
  deleteButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
    borderRadius: "8px",
    color: "var(--oc-text-faint)",
    ":hover": {
      color: "var(--oc-danger-text)",
      backgroundColor: "var(--oc-danger-bg)",
    },
  },
});

function formatDate(dateString: string): string {
  const date = new Date(dateString);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return "Just now";
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;
  
  return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

export const SessionHistory: React.FC<SessionHistoryProps> = ({
  host,
  shared,
  onSharedChange,
  onSelectSession,
  onClose,
}) => {
  const styles = useStyles();
  const [sessions, setSessions] = React.useState<OpencodeSessionInfo[]>([]);

  React.useEffect(() => {
    listSessions(host, shared)
      .then((items) => setSessions(z.array(SessionInfoSchema).catch([]).parse(items)))
      .catch(() => setSessions([]));
  }, [host, shared]);

  const handleDelete = (e: React.MouseEvent, sessionId: string) => {
    e.stopPropagation();
    deleteSession(sessionId)
      .then(() => listSessions(host, shared))
      .then((items) => setSessions(z.array(SessionInfoSchema).catch([]).parse(items)))
      .catch(() => setSessions([]));
  };

  const hostLabel = host === "powerpoint" ? "PowerPoint" : host === "word" ? "Word" : host === "excel" ? "Excel" : "OneNote";

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button
          icon={<ArrowLeft24Regular />}
          appearance="subtle"
          className={styles.backButton}
          onClick={onClose}
          aria-label="Back"
        />
        <Text className={styles.headerTitle}>{hostLabel} History</Text>
      </div>
      <div className={styles.filterRow}>
        <Text className={styles.filterLabel}>Show</Text>
        <div className={styles.filterGroup}>
          <Button
            appearance={shared ? "subtle" : "primary"}
            className={styles.filterButton}
            onClick={() => onSharedChange(false)}
          >
            This app
          </Button>
          <Button
            appearance={shared ? "primary" : "subtle"}
            className={styles.filterButton}
            onClick={() => onSharedChange(true)}
          >
            All history
          </Button>
        </div>
      </div>

      <div className={styles.list}>
        {sessions.length === 0 ? (
          <div className={styles.emptyState}>
            No saved conversations yet.<br />
            Start chatting to create one!
          </div>
        ) : (
          sessions.map((session) => (
            <div
              key={session.id}
              className={styles.sessionItem}
              onClick={() => onSelectSession(session)}
            >
              <div className={styles.sessionContent}>
                <div className={styles.sessionTitle}>{session.title}</div>
                <div className={styles.sessionMeta}>
                  <span>{formatDate(new Date(session.time.updated).toISOString())}</span>
                  <span>•</span>
                  <span>{session.directory.split(/[\\/]/).pop() || session.directory}</span>
                </div>
              </div>
              <Button
                icon={<Delete24Regular />}
                appearance="subtle"
                className={styles.deleteButton}
                onClick={(e) => handleDelete(e, session.id)}
                aria-label="Delete session"
              />
            </div>
          ))
        )}
      </div>
    </div>
  );
};
