import { describe, expect, it } from "vitest";
import { expandMentions, insertMention, mentionParts, mentionPaths, mentionQuery, resolveMention } from "./file-mentions";

describe("file mentions", () => {
  it("finds active mention queries", () => {
    expect(mentionQuery("check @src/App", 14)).toEqual({ start: 6, end: 14, query: "src/App" });
  });

  it("inserts quoted mentions for paths with spaces", () => {
    expect(insertMention("open @rea", 9, "docs/read me.md")).toEqual({
      value: 'open @"docs/read me.md" ',
      caret: 24,
    });
  });

  it("extracts plain and quoted mentions", () => {
    expect(mentionPaths('check @src/App.tsx and @"docs/read me.md".')).toEqual([
      "src/App.tsx",
      "docs/read me.md",
    ]);
  });

  it("resolves relative mentions against the active folder", () => {
    expect(resolveMention("/repo", "src/App.tsx")).toBe("/repo/src/App.tsx");
  });

  it("builds deduped file prompt parts", () => {
    expect(mentionParts("see @src/App.tsx and @src/App.tsx", "/repo")).toEqual([
      {
        type: "file",
        mime: "text/plain",
        url: "file:///repo/src/App.tsx",
        filename: "App.tsx",
      },
    ]);
  });

  it("expands slash-command mention arguments to absolute paths", () => {
    expect(expandMentions('review @src/App.tsx and @"docs/read me.md"', "/repo")).toBe(
      'review @/repo/src/App.tsx and @"/repo/docs/read me.md"',
    );
  });
});
