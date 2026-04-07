import * as React from "react";
import { useRef, useEffect, useMemo } from "react";
import { makeStyles } from "@fluentui/react-components";
import type { SlashCommand } from "../lib/opencode-schemas";

export interface SlashPopoverProps {
  /** Current text typed after the `/` trigger (e.g. if input is "/rev" this is "rev") */
  filter: string;
  /** All available slash commands fetched from the server */
  commands: SlashCommand[];
  /** Index of the currently highlighted item (controlled) */
  selectedIndex: number;
  /** Called when the user picks a command (Enter/click) */
  onSelect: (command: SlashCommand) => void;
  /** Called when highlight changes via mouse hover */
  onHighlight: (index: number) => void;
}

/* ---- fuzzy scoring (mirrors model-search.ts pattern) ---- */

function normalize(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function score(query: string, value: string) {
  const needle = normalize(query);
  const hay = normalize(value);
  if (!needle) return 0;
  if (!hay) return Number.NEGATIVE_INFINITY;

  let pos = -1;
  let total = 0;
  let streak = 0;

  for (const char of needle) {
    const next = hay.indexOf(char, pos + 1);
    if (next === -1) return Number.NEGATIVE_INFINITY;
    const gap = pos === -1 ? next : next - pos - 1;
    streak = next === pos + 1 ? streak + 1 : 1;
    total += 12 - Math.min(gap, 10) + streak * 6;
    if (next === 0) total += 10;
    pos = next;
  }

  total -= hay.length - needle.length;
  return total;
}

function filterCommands(commands: SlashCommand[], query: string): SlashCommand[] {
  const value = query.trim();
  if (!value) return commands;

  return commands
    .map((cmd) => {
      const nameScore = score(value, cmd.name);
      const descScore = cmd.description ? score(value, cmd.description) : Number.NEGATIVE_INFINITY;
      const best = Math.max(nameScore, descScore);
      return { cmd, best };
    })
    .filter((item) => item.best > Number.NEGATIVE_INFINITY)
    .sort((a, b) => b.best - a.best || a.cmd.name.localeCompare(b.cmd.name))
    .map((item) => item.cmd);
}

const SOURCE_LABELS: Record<string, string> = {
  command: "",
  mcp: "mcp",
  skill: "skill",
};

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
  name: {
    fontWeight: 600,
    color: "var(--text-strong)",
    whiteSpace: "nowrap",
  },
  description: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    color: "var(--text-weak)",
  },
  badge: {
    flexShrink: 0,
    fontSize: "10px",
    fontWeight: 600,
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    padding: "1px 5px",
    borderRadius: "4px",
    backgroundColor: "var(--oc-bg-soft)",
    color: "var(--text-weak)",
    border: "1px solid var(--oc-border)",
  },
  empty: {
    padding: "10px 12px",
    fontSize: "12px",
    color: "var(--text-weak)",
    textAlign: "center" as const,
  },
});

export const SlashPopover: React.FC<SlashPopoverProps> = ({
  filter,
  commands,
  selectedIndex,
  onSelect,
  onHighlight,
}) => {
  const styles = useStyles();
  const listRef = useRef<HTMLUListElement>(null);

  const filtered = useMemo(() => filterCommands(commands, filter), [commands, filter]);

  // scroll selected item into view
  useEffect(() => {
    const list = listRef.current;
    if (!list) return;
    const item = list.children[selectedIndex] as HTMLElement | undefined;
    item?.scrollIntoView({ block: "nearest" });
  }, [selectedIndex]);

  if (filtered.length === 0) {
    return (
      <div className={styles.root}>
        <div className={styles.empty}>No matching commands</div>
      </div>
    );
  }

  return (
    <div className={styles.root} role="listbox" aria-label="Slash commands">
      <ul className={styles.list} ref={listRef}>
        {filtered.map((cmd, i) => {
          const sourceLabel = SOURCE_LABELS[cmd.source || "command"];
          return (
            <li
              key={cmd.name}
              role="option"
              aria-selected={i === selectedIndex}
              className={`${styles.item} ${i === selectedIndex ? styles.selected : ""}`.trim()}
              onMouseEnter={() => onHighlight(i)}
              onMouseDown={(e) => {
                e.preventDefault(); // keep focus on textarea
                onSelect(cmd);
              }}
            >
              <span className={styles.name}>/{cmd.name}</span>
              {cmd.description && (
                <span className={styles.description}>{cmd.description}</span>
              )}
              {sourceLabel && <span className={styles.badge}>{sourceLabel}</span>}
            </li>
          );
        })}
      </ul>
    </div>
  );
};

/** Re-export the filter function so ChatInput can compute filtered length for index clamping */
export { filterCommands };
