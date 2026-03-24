import * as React from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";
import { Delete24Regular, ArrowLeft24Regular } from "@fluentui/react-icons";
import type { SavedSession, OfficeHost } from "../sessionStorage";
import { getSavedSessions, deleteSession } from "../sessionStorage";

interface SessionHistoryProps {
  host: OfficeHost;
  onSelectSession: (session: SavedSession) => void;
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
      borderColor: "var(--colorNeutralStroke1Hover)",
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
  onSelectSession,
  onClose,
}) => {
  const styles = useStyles();
  const [sessions, setSessions] = React.useState<SavedSession[]>([]);

  React.useEffect(() => {
    setSessions(getSavedSessions(host));
  }, [host]);

  const handleDelete = (e: React.MouseEvent, sessionId: string) => {
    e.stopPropagation();
    deleteSession(host, sessionId);
    setSessions(getSavedSessions(host));
  };

  const hostLabel = host === "powerpoint" ? "PowerPoint" : host === "word" ? "Word" : "Excel";

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
                  <span>{formatDate(session.updatedAt)}</span>
                  <span>•</span>
                  <span>{session.messages.filter(m => m.sender === "user").length} messages</span>
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
