import * as React from "react";
import { useEffect, useMemo, useRef } from "react";
import { makeStyles } from "@fluentui/react-components";

export interface MentionPopoverProps {
  items: string[];
  selectedIndex: number;
  onSelect: (path: string) => void;
  onHighlight: (index: number) => void;
  loading?: boolean;
}

const useStyles = makeStyles({
  root: {
    position: "absolute",
    bottom: "100%",
    left: "0",
    right: "0",
    marginBottom: "4px",
    borderRadius: "10px",
    border: "1px solid var(--oc-border)",
    backgroundColor: "var(--oc-bg)",
    boxShadow: "var(--oc-shadow)",
    overflow: "hidden",
    zIndex: 100,
    maxHeight: "240px",
    overflowY: "auto",
  },
  list: {
    listStyleType: "none",
    margin: "0",
    padding: "4px",
  },
  item: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "7px 10px",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "13px",
    lineHeight: "1.3",
    ":hover": {
      backgroundColor: "var(--oc-bg-soft-hover)",
    },
  },
  selected: {
    backgroundColor: "var(--oc-bg-soft-hover)",
  },
  sigil: {
    fontWeight: 700,
    color: "var(--oc-accent)",
    flexShrink: 0,
  },
  path: {
    color: "var(--text-strong)",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  empty: {
    padding: "10px 12px",
    fontSize: "12px",
    color: "var(--text-weak)",
    textAlign: "center" as const,
  },
});

export const MentionPopover: React.FC<MentionPopoverProps> = ({ items, selectedIndex, onSelect, onHighlight, loading = false }) => {
  const styles = useStyles();
  const listRef = useRef<HTMLUListElement>(null);
  const safeItems = useMemo(() => items.slice(0, 20), [items]);

  useEffect(() => {
    const list = listRef.current;
    if (!list) return;
    const item = list.children[selectedIndex] as HTMLElement | undefined;
    item?.scrollIntoView({ block: "nearest" });
  }, [selectedIndex]);

  if (loading) {
    return <div className={styles.root}><div className={styles.empty}>Searching files…</div></div>;
  }

  if (safeItems.length === 0) {
    return <div className={styles.root}><div className={styles.empty}>No matching files or folders</div></div>;
  }

  return (
    <div className={styles.root} role="listbox" aria-label="File mentions">
      <ul className={styles.list} ref={listRef}>
        {safeItems.map((item, index) => (
          <li
            key={item}
            role="option"
            aria-selected={index === selectedIndex}
            className={`${styles.item} ${index === selectedIndex ? styles.selected : ""}`.trim()}
            onMouseEnter={() => onHighlight(index)}
            onMouseDown={(event) => {
              event.preventDefault();
              onSelect(item);
            }}
          >
            <span className={styles.sigil}>@</span>
            <span className={styles.path}>{item}</span>
          </li>
        ))}
      </ul>
    </div>
  );
};
