import * as React from "react";
import { Button, Spinner, makeStyles } from "@fluentui/react-components";

type BrowseResponse = {
  path: string;
  parent: string | null;
  dirs: string[];
};

interface FolderPickerDialogProps {
  open: boolean;
  initialPath?: string;
  onClose: () => void;
  onSelect: (path: string) => void;
}

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
    width: "min(100%, 640px)",
    maxHeight: "min(80vh, 720px)",
    border: "1px solid var(--oc-border)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    padding: "16px 18px 12px",
    borderBottom: "1px solid var(--oc-border)",
  },
  title: {
    fontWeight: 700,
    fontSize: "14px",
  },
  hint: {
    fontSize: "12px",
    color: "var(--oc-text-faint)",
    lineHeight: "1.5",
  },
  body: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px 18px",
    minHeight: 0,
  },
  row: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  input: {
    width: "100%",
    borderRadius: "10px",
    border: "1px solid var(--oc-border)",
    background: "var(--oc-bg-soft)",
    color: "var(--oc-text)",
    padding: "10px 12px",
    fontSize: "12px",
    boxSizing: "border-box",
  },
  path: {
    fontSize: "12px",
    color: "var(--oc-text-muted)",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    borderRadius: "10px",
    padding: "10px 12px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  list: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    overflowY: "auto",
    minHeight: "180px",
    maxHeight: "360px",
  },
  item: {
    width: "100%",
    justifyContent: "flex-start",
    borderRadius: "10px",
    padding: "10px 12px",
    background: "var(--oc-bg-soft)",
    border: "1px solid var(--oc-border)",
    color: "var(--oc-text)",
  },
  empty: {
    fontSize: "12px",
    color: "var(--oc-text-faint)",
    padding: "12px",
    textAlign: "center",
    border: "1px dashed var(--oc-border)",
    borderRadius: "10px",
  },
  error: {
    fontSize: "12px",
    color: "var(--oc-danger-text)",
    background: "var(--oc-danger-bg)",
    border: "1px solid var(--oc-danger-border)",
    borderRadius: "10px",
    padding: "10px 12px",
  },
  footer: {
    display: "flex",
    gap: "8px",
    justifyContent: "flex-end",
    padding: "14px 18px 18px",
    borderTop: "1px solid var(--oc-border)",
  },
  grow: {
    flex: 1,
  },
});

function joinPath(base: string, name: string) {
  if (base === "/") return `/${name}`;
  return `${base.replace(/\/+$/, "")}/${name}`;
}

export const FolderPickerDialog: React.FC<FolderPickerDialogProps> = ({ open, initialPath, onClose, onSelect }) => {
  const styles = useStyles();
  const [path, setPath] = React.useState("");
  const [input, setInput] = React.useState("");
  const [parent, setParent] = React.useState<string | null>(null);
  const [dirs, setDirs] = React.useState<string[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState("");

  const load = React.useCallback(async (next?: string) => {
    setLoading(true);
    setError("");

    try {
      const query = next ? `?path=${encodeURIComponent(next)}` : "";
      const response = await fetch(`/api/browse${query}`);
      const data = await response.json();

      if (!response.ok) {
        throw new Error(typeof data?.error === "string" ? data.error : "Failed to browse folders");
      }

      const result = data as BrowseResponse;
      setPath(result.path);
      setInput(result.path);
      setParent(result.parent);
      setDirs(Array.isArray(result.dirs) ? result.dirs : []);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    if (!open) return;
    void load(initialPath);
  }, [initialPath, load, open]);

  React.useEffect(() => {
    if (!open) return;

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [onClose, open]);

  if (!open) return null;

  return (
    <div className={styles.backdrop} onClick={onClose}>
      <div className={styles.card} role="dialog" aria-label="Choose folder" onClick={(event) => event.stopPropagation()}>
        <header className={styles.header}>
          <div className={styles.title}>Choose default folder</div>
          <div className={styles.hint}>New chats will start in this folder. Existing sessions keep their own folder.</div>
        </header>
        <div className={styles.body}>
          <div className={styles.row}>
            <input
              className={styles.input}
              value={input}
              onChange={(event) => setInput(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  void load(input);
                }
              }}
              placeholder="Enter a folder path"
            />
            <Button appearance="secondary" onClick={() => void load(input)} disabled={loading}>Go</Button>
          </div>
          <div className={styles.row}>
            <Button appearance="secondary" onClick={() => parent && void load(parent)} disabled={loading || !parent}>Up</Button>
            <Button appearance="secondary" onClick={() => path && onSelect(path)} disabled={loading || !path}>Use current folder</Button>
          </div>
          <div className={styles.path}>{path || "Loading..."}</div>
          {error && <div className={styles.error}>{error}</div>}
          <div className={styles.list}>
            {loading ? (
              <div className={styles.empty}><Spinner size="tiny" /> Loading folders…</div>
            ) : dirs.length === 0 ? (
              <div className={styles.empty}>No subfolders found.</div>
            ) : (
              dirs.map((dir) => (
                <Button key={dir} appearance="subtle" className={styles.item} onClick={() => void load(joinPath(path, dir))}>
                  {dir}
                </Button>
              ))
            )}
          </div>
        </div>
        <footer className={styles.footer}>
          <div className={styles.grow} />
          <Button appearance="secondary" onClick={onClose}>Cancel</Button>
          <Button appearance="primary" onClick={() => path && onSelect(path)} disabled={loading || !path}>Choose folder</Button>
        </footer>
      </div>
    </div>
  );
};
