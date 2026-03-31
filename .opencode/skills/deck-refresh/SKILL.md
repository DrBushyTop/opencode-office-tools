---
name: deck-refresh
description: Refresh an existing PowerPoint deck in place while preserving template fit, using inspect-first routing, slide-scoped edits, and visual QA.
references:
  - edit-routing
---

# Deck Refresh

Use this skill when the user wants to modernize, tighten, or retheme slides without rebuilding the whole deck.

## Workflow

1. **Inspect the current deck**
   - Use `get_presentation_overview` and `get_presentation_structure` first.
   - Use `get_slide_image` on target slides.
   - Use `list_slide_layouts` if new slides may be added.
2. **Choose the safest starting point**
   - Edit in place when the slide structure is already correct.
   - Use `duplicate_slide` first when the change is exploratory, high-risk, or should preserve an untouched original.
3. **Inspect slide structure before precision edits**
   - Use `list_slide_shapes` on each target slide.
4. **Route each change narrowly**
   - One text shape with formatting fidelity concerns → `read_slide_text` + `edit_slide_text`
   - Several related text shapes on one slide → `edit_slide_xml`
   - Chart refresh → `edit_slide_chart`
   - Master/theme adjustment → `edit_slide_master`
   - New slide that should fit an existing layout → `list_slide_layouts` + `create_slide_from_layout`
   - Shape-only cleanup such as position, fill, line, naming, grouping → `manage_slide_shapes`
   - Images/tables → `manage_slide_media` / `manage_slide_table`
5. **Refresh targeting after round-trips**
   - Reuse returned ids/refs when available.
   - Otherwise rerun `list_slide_shapes` before the next pinpoint edit.
6. **Run visual QA**
   - Review each changed slide with `get_slide_image` so validation is based on the current exact visual of the slide.
   - Use the `visual-qa` agent for a fresh-eyes check on changed slides.

## Working Rules

- Prefer slide-scoped edits over whole-slide regeneration.
- Prefer native layout and text tools before generic shape edits.
- Keep `add_slide_from_code` as a last resort.

## References

- `references/edit-routing.md`
