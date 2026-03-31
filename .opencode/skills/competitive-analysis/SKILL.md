---
name: competitive-analysis
description: Analyze an existing PowerPoint deck or provided reference material, identify reusable presentation patterns, and recommend the narrowest in-repo editing workflow for each change.
references:
  - ppt-patterns
---

# Competitive Analysis

Use this skill when the task is to study a deck, compare slide approaches, or extract reusable presentation patterns before editing.

## Goal

Turn observations into an execution plan that fits this repo's PowerPoint tool model.

## Workflow

1. **Inspect first**
   - Use `get_presentation_overview` for deck inventory.
   - Use `get_presentation_structure` for theme, slide size, and selection context.
   - Use `list_slide_layouts` if layout reuse or template fit matters.
2. **Sample the relevant slides**
   - Use `get_slide_image` for visual comparison.
   - Use `list_slide_shapes` when structure, grouping, or text-container patterns matter.
   - Use `get_presentation_content` for wording, title cadence, and repeated phrases.
3. **Identify repeatable patterns**
   - Layout pattern: title slide, section divider, two-column slide, quote, chart slide, comparison, timeline, summary.
   - Structural pattern: placeholder usage, repeated text box groupings, chart placement, image treatment, margin system.
   - Style pattern: color accents, typography scale, density, alignment, icon/image handling.
4. **Map each recommendation to a tool path**
   - Single text shape → `read_slide_text` + `edit_slide_text`
   - Multi-shape same-slide text → `edit_slide_xml`
   - Chart change → `edit_slide_chart`
   - Layout-based new slide → `list_slide_layouts` + `create_slide_from_layout`
   - Prototype variation → `duplicate_slide` + targeted edits
   - Geometry/fill/grouping cleanup → `manage_slide_shapes`
   - Use `add_slide_from_code` only if the native tool path is still insufficient.
5. **Call out reference freshness risks**
   - If the plan involves round-trip edits or duplication, note that slide ids and shape refs must be refreshed before later targeted edits.

## Output

Return a short working brief with:

- slides or layouts inspected
- repeated patterns worth reusing
- mismatches or gaps to fix
- recommended tool route for each requested change
- any ref-refresh or visual-QA checkpoints

## References

- `references/ppt-patterns.md`
