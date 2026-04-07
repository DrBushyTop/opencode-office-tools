import * as React from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";
import { Delete24Regular, ArrowLeft24Regular } from "@fluentui/react-icons";
import { z } from "zod";
import type { OfficeHost } from "../sessionStorage";
import {
  deleteSession,
  listSessions,
  type OpencodeSessionInfo,
} from "../lib/opencode-session-history";

const SessionInfoSchema = z.object({
  id: z.string().min(1),
  title: z.string(),
  directory: z.string(),
  time: z.object({ created: z.number(), updated: z.number() }),
}) satisfies z.ZodType<OpencodeSessionInfo>;

/* ------------------------------------------------------------------ */
/*  Props                                                             */
/* ------------------------------------------------------------------ */

interface SessionHistoryProps {
  host: OfficeHost;
  shared: boolean;
  directory?: string;
  onSharedChange: (shared: boolean) => void;
  onSelectSession: (session: OpencodeSessionInfo) => void;
  onClose: () => void;
}

/* ------------------------------------------------------------------ */
/*  Styles                                                            */
/* ------------------------------------------------------------------ */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    background: "var(--oc-bg)",
  },
  toolbar: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "14px 14px 12px",
    borderBottom: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
  },
  toolbarTitle: {
    fontWeight: "700",
    fontSize: "13px",
    color: "var(--oc-text)",
    letterSpacing: "0.01em",
  },
  scopeBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "10px",
    padding: "12px 14px",
    borderBottom: "1px solid var(--oc-border)",
    background: "var(--oc-bg)",
  },
  scopeCaption: {
    fontSize: "11px",
    fontWeight: "700",
    color: "var(--oc-text-faint)",
    textTransform: "uppercase",
    letterSpacing: "0.08em",
  },
  pillGroup: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px",
    borderRadius: "999px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
  },
  pill: {
    minWidth: "76px",
    padding: "0 12px",
    fontSize: "12px",
    borderRadius: "999px",
    height: "28px",
  },
  navBtn: {
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
  feed: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    padding: "10px",
    scrollbarColor: "var(--oc-text-faint) transparent",
    scrollbarWidth: "thin",
    "&::-webkit-scrollbar": { width: "8px" },
    "&::-webkit-scrollbar-track": { backgroundColor: "transparent" },
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
  placeholder: {
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
  card: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "12px",
    borderRadius: "12px",
    cursor: "pointer",
    marginBottom: "8px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    transition:
      "transform 0.15s ease, border-color 0.15s ease, background 0.15s ease, box-shadow 0.15s ease",
    ":hover": {
      background: "var(--oc-bg-soft-hover)",
      border: "1px solid var(--oc-border-strong)",
      transform: "translateY(-1px)",
      boxShadow: "0 10px 24px rgba(0,0,0,0.08)",
    },
  },
  cardContent: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
  },
  cardTitle: {
    fontSize: "13px",
    fontWeight: "600",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    color: "var(--oc-text)",
  },
  cardMeta: {
    fontSize: "11px",
    color: "var(--oc-text-faint)",
    marginTop: "4px",
    display: "flex",
    gap: "8px",
  },
  trashBtn: {
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

/* ------------------------------------------------------------------ */
/*  Helpers                                                           */
/* ------------------------------------------------------------------ */

const HOST_LABELS: Record<OfficeHost, string> = {
  powerpoint: "PowerPoint",
  word: "Word",
  excel: "Excel",
  onenote: "OneNote",
};

/** Produce a compact relative-time label from epoch millis. */
function relativeTime(epochMs: number): string {
  const delta = Date.now() - epochMs;
  if (delta < 60_000) return "Just now";
  const mins = Math.floor(delta / 60_000);
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(delta / 3_600_000);
  if (hrs < 24) return `${hrs}h ago`;
  const days = Math.floor(delta / 86_400_000);
  if (days < 7) return `${days}d ago`;
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
  }).format(epochMs);
}

/** Extract the last path segment as a short folder label. */
function folderLabel(dir: string): string {
  return dir.split(/[\\/]/).pop() || dir;
}

/* ------------------------------------------------------------------ */
/*  Sub-component: single session row                                 */
/* ------------------------------------------------------------------ */

interface EntryProps {
  session: OpencodeSessionInfo;
  onSelect: () => void;
  onDelete: (e: React.MouseEvent) => void;
}

const SessionEntry: React.FC<EntryProps> = ({ session, onSelect, onDelete }) => {
  const cls = useStyles();
  return (
    <article className={cls.card} onClick={onSelect} role="button" tabIndex={0}>
      <div className={cls.cardContent}>
        <div className={cls.cardTitle}>{session.title}</div>
        <div className={cls.cardMeta}>
          <span>{relativeTime(session.time.updated)}</span>
          <span aria-hidden>&middot;</span>
          <span>{folderLabel(session.directory)}</span>
        </div>
      </div>
      <Button
        icon={<Delete24Regular />}
        appearance="subtle"
        className={cls.trashBtn}
        onClick={onDelete}
        aria-label="Remove"
      />
    </article>
  );
};

/* ------------------------------------------------------------------ */
/*  Main component                                                    */
/* ------------------------------------------------------------------ */

export const SessionHistory: React.FC<SessionHistoryProps> = ({
  host,
  shared,
  directory,
  onSharedChange,
  onSelectSession,
  onClose,
}) => {
  const cls = useStyles();
  const [sessions, setSessions] = React.useState<OpencodeSessionInfo[]>([]);

  const refresh = React.useCallback(
    () =>
      listSessions(host, shared, directory)
        .then((raw) => setSessions(z.array(SessionInfoSchema).catch([]).parse(raw)))
        .catch(() => setSessions([])),
    [directory, host, shared],
  );

  React.useEffect(() => { refresh(); }, [refresh]);

  const handleDelete = (e: React.MouseEvent, id: string) => {
    e.stopPropagation();
    const session = sessions.find((item) => item.id === id);
    deleteSession(id, session?.directory).then(refresh).catch(refresh);
  };

  return (
    <section className={cls.root}>
      {/* top bar */}
      <nav className={cls.toolbar}>
        <Button
          icon={<ArrowLeft24Regular />}
          appearance="subtle"
          className={cls.navBtn}
          onClick={onClose}
          aria-label="Go back"
        />
        <Text className={cls.toolbarTitle}>{HOST_LABELS[host]} History</Text>
      </nav>

      {/* scope toggle */}
      <div className={cls.scopeBar}>
        <Text className={cls.scopeCaption}>Scope</Text>
        <div className={cls.pillGroup}>
          <Button
            appearance={shared ? "subtle" : "primary"}
            className={cls.pill}
            onClick={() => onSharedChange(false)}
          >
            This folder
          </Button>
          <Button
            appearance={shared ? "primary" : "subtle"}
            className={cls.pill}
            onClick={() => onSharedChange(true)}
          >
            All folders
          </Button>
        </div>
      </div>

      {/* session list */}
      <div className={cls.feed}>
        {sessions.length === 0 ? (
          <div className={cls.placeholder}>
            No sessions yet.
            <br />
            Start a conversation and it will show up here.
          </div>
        ) : (
          sessions.map((s) => (
            <SessionEntry
              key={s.id}
              session={s}
              onSelect={() => onSelectSession(s)}
              onDelete={(e) => handleDelete(e, s.id)}
            />
          ))
        )}
      </div>
    </section>
  );
};
