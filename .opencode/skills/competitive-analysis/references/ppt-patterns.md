# PowerPoint Pattern Notes

Use this checklist when turning deck observations into an editing plan.

## Pattern Types

- **Layout pattern**: recurring placeholder map or spatial arrangement.
- **Content pattern**: repeated headline length, bullet density, metric framing, caption style.
- **Structure pattern**: separate shapes for animated units, consistent grouping, predictable chart/text pairings.
- **Theme pattern**: accent colors, line weights, card fills, background treatment.

## Routing Hints

- Need to preserve rich text structure in one shape: `read_slide_text` then `edit_slide_text`.
- Need several coordinated text changes on one slide: `edit_slide_xml`.
- Need a new slide that fits the deck's template language: `list_slide_layouts` then `create_slide_from_layout`.
- Need a variation of an existing slide: `duplicate_slide`, then inspect the duplicate and edit it.
- Need only box geometry, fill, naming, or grouping cleanup: `manage_slide_shapes`.

## Watch For

- Repeated manual layouts that should become layout-driven creation.
- Text-heavy slides that should be split into separate shapes for animation or readability.
- Prototype edits performed on the wrong slide instead of a duplicate.
- Stale shape refs after duplicate or round-trip edits.
