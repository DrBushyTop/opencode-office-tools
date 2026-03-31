---
name: deck-check
description: Review a PowerPoint deck for layout, readability, structure, and workflow risks using the repo's inspection and QA tools.
references:
  - review-checklist
---

# Deck Check

Use this skill when the task is to audit a deck before delivery, after edits, or before a larger refresh.

## Workflow

1. **Inspect context**
   - Use `get_presentation_overview` and `get_presentation_structure` first.
2. **Review visuals slide by slide**
   - Use `get_slide_image` for each slide under review so you are checking the current exact visual, not only structure or text.
3. **Inspect structure when needed**
   - Use `list_slide_shapes` when layout issues may actually be shape naming, grouping, placeholder misuse, or incorrect text containers.
   - Use `read_slide_text` for a specific shape only when text fidelity looks suspect.
   - Use `get_slide_animations` when animation order or timing is part of the review.
4. **Assess workflow safety**
   - Note slides that would be better handled by duplication first.
   - Note slides whose content should move to layout-based creation for future consistency.
   - Note stale-ref risk after any round-trip workflow, and state that returned refreshed slide ids or shape refs should be reused; otherwise `list_slide_shapes` should be rerun before the next precise edit.
5. **Report only actionable findings**
   - Prioritize broken layout, unreadable text, structural problems, and obvious consistency issues.

## Output

For each issue, include:

- slide reference
- severity
- category
- short description
- fix hint with the likely tool path

## References

- `references/review-checklist.md`
