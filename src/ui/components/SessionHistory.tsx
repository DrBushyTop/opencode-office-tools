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
  entryRow: {
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
  entryBody: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
  },
  entryTitle: {
    fontSize: "13px",
    fontWeight: "600",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    color: "var(--oc-text)",
  },
  entryDetail: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    marginTop: "4px",
    display: "flex",
    gap: "8px",
  },
  removeAction: {
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

/** Convert an ISO date string into a compact relative label. */
function describeRecency(isoString: string): string {
  const then = new Date(isoString).getTime();
  const elapsed = Date.now() - then;

  if (elapsed < 60_000) return "Moments ago";

  const minutes = Math.floor(elapsed / 60_000);
  if (minutes < 60) return `${minutes} min ago`;

  const hours = Math.floor(elapsed / 3_600_000);
  if (hours < 24) return `${hours}h ago`;

  const days = Math.floor(elapsed / 86_400_000);
  if (days < 7) return `${days}d ago`;

  return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" }).format(new Date(isoString));
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

  const removeEntry = (e: React.MouseEvent, sessionId: string) => {
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
            Nothing here yet.<br />
            Your conversations will appear once you start chatting.
          </div>
        ) : (
          sessions.map((session) => (
            <div
              key={session.id}
              className={styles.entryRow}
              onClick={() => onSelectSession(session)}
            >
              <div className={styles.entryBody}>
                <div className={styles.entryTitle}>{session.title}</div>
                <div className={styles.entryDetail}>
                  <span>{describeRecency(new Date(session.time.updated).toISOString())}</span>
                  <span>•</span>
                  <span>{session.directory.split(/[\\/]/).pop() || session.directory}</span>
                </div>
              </div>
              <Button
                icon={<Delete24Regular />}
                appearance="subtle"
                className={styles.removeAction}
                onClick={(e) => removeEntry(e, session.id)}
                aria-label="Delete session"
              />
            </div>
          ))
        )}
      </div>
    </div>
  );
};
