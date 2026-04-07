const mention = /(^|[\s([{"'])@(?:"([^"]+)"|(\S+))/g;
const active = /(^|[\s([{"'])@([^\s@"]*)$/;

function trim(value: string) {
  return value.replace(/[.,!?;:)}\]"']+$/, "");
}

function encode(input: string) {
  let value = input.replace(/\\/g, "/");
  if (/^[A-Za-z]:/.test(value)) value = `/${value}`;
  return value
    .split("/")
    .map((part, index) => (index === 1 && /^[A-Za-z]:$/.test(part) ? part : encodeURIComponent(part)))
    .join("/");
}

function tail(path: string) {
  const parts = path.split(/[\\/]/).filter(Boolean);
  return parts[parts.length - 1] || path;
}

export function mentionQuery(text: string, caret: number) {
  const match = text.slice(0, caret).match(active);
  if (!match) return null;
  return {
    start: caret - (match[2] || "").length - 1,
    end: caret,
    query: match[2] || "",
  };
}

export function insertMention(text: string, caret: number, path: string) {
  const match = mentionQuery(text, caret);
  if (!match) return null;
  const token = path.includes(" ") ? `@"${path}" ` : `@${path} `;
  const next = `${text.slice(0, match.start)}${token}${text.slice(match.end)}`;
  return {
    value: next,
    caret: match.start + token.length,
  };
}

export function mentionPaths(text: string) {
  return Array.from(text.matchAll(mention)).flatMap((item) => {
    const value = item[2] || trim(item[3] || "");
    if (!value) return [];
    return [value];
  });
}

export function resolveMention(root: string, input: string) {
  if (input.startsWith("/")) return input;
  if (/^[A-Za-z]:[\\/]/.test(input) || /^[A-Za-z]:$/.test(input)) return input;
  if (input.startsWith("\\\\") || input.startsWith("//")) return input;
  return `${root.replace(/[\\/]+$/, "")}/${input}`;
}

export function mentionParts(text: string, root: string) {
  const seen = new Set<string>();
  return mentionPaths(text).flatMap((item) => {
    const path = resolveMention(root, item);
    if (seen.has(path)) return [];
    seen.add(path);
    return [{
      type: "file" as const,
      mime: "text/plain",
      url: `file://${encode(path)}`,
      filename: tail(item),
    }];
  });
}

export function expandMentions(text: string, root: string) {
  return text.replace(mention, (full, prefix, quoted, plain) => {
    const value = quoted || trim(plain || "");
    if (!value) return full;
    const path = resolveMention(root, value);
    const quotedPath = /\s/.test(path) ? `"${path}"` : path;
    return `${prefix}@${quotedPath}`;
  });
}
