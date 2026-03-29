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
    height: "100%",
    backgroundColor: "var(--colorNeutralBackground2)",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
    borderBottom: "1px solid var(--colorNeutralStroke2)",
  },
  headerTitle: {
    fontWeight: "600",
    fontSize: "14px",
  },
  filterRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    padding: "10px 12px",
    borderBottom: "1px solid var(--colorNeutralStroke2)",
    backgroundColor: "var(--colorNeutralBackground1)",
  },
  filterLabel: {
    fontSize: "11px",
    fontWeight: "600",
    color: "var(--colorNeutralForeground3)",
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  filterGroup: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  filterButton: {
    minWidth: "76px",
    padding: "0 10px",
    fontSize: "12px",
    borderRadius: "999px",
  },
  backButton: {
    minWidth: "32px",
    padding: "4px",
  },
  list: {
    flex: 1,
    overflowY: "auto",
    padding: "8px",
  },
  emptyState: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    color: "var(--colorNeutralForeground3)",
    fontSize: "14px",
    padding: "20px",
    textAlign: "center",
  },
  sessionItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "10px 12px",
    borderRadius: "6px",
    cursor: "pointer",
    marginBottom: "4px",
    backgroundColor: "var(--colorNeutralBackground1)",
    border: "1px solid var(--colorNeutralStroke2)",
    transition: "all 0.15s ease",
    ":hover": {
      backgroundColor: "var(--colorNeutralBackground1Hover)",
      border: "1px solid var(--colorNeutralStroke1Hover)",
    },
  },
  sessionContent: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
  },
  sessionTitle: {
    fontSize: "13px",
    fontWeight: "500",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    color: "var(--colorNeutralForeground1)",
  },
  sessionMeta: {
    fontSize: "11px",
    color: "var(--colorNeutralForeground3)",
    marginTop: "2px",
    display: "flex",
    gap: "8px",
  },
  deleteButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
    color: "var(--colorNeutralForeground3)",
    ":hover": {
      color: "var(--colorPaletteRedForeground1)",
      backgroundColor: "var(--colorPaletteRedBackground1)",
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
