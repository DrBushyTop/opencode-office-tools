# Deck Refresh Routing

## Quick Router

| Need | Preferred tool path |
|---|---|
| Understand deck and theme | `get_presentation_overview` + `get_presentation_structure` |
| Understand available layouts | `list_slide_layouts` |
| Inspect current slide structure | `list_slide_shapes` |
| Edit one text shape | `read_slide_text` + `edit_slide_text` |
| Edit several text shapes on one slide | `edit_slide_xml` in one batch |
| Single-slide diagrams or advanced OOXML formatting | `edit_slide_xml` |
| Edit chart | `edit_slide_chart` |
| Edit supported master/theme elements | `edit_slide_master` |
| Create slide from template layout | `list_slide_layouts` + `create_slide_from_layout` |
| Prototype from existing slide | `duplicate_slide` |
| Create or substantially arrange shapes | `execute_office_js` |
| Small sparse shape cleanup | `manage_slide_shapes` |
| Insert or replace image | `manage_slide_media` |
| Insert or update native table | `manage_slide_table` |
| Custom visualization, geometric shape work, or unsupported host operation | `execute_office_js` |

## Refresh Rule

After `edit_slide_text`, `edit_slide_xml`, `edit_slide_chart`, `edit_slide_master`, `duplicate_slide`, or Open XML animation/notes/transition edits, refresh slide targeting before the next precise mutation unless the tool already returned updated refs.

Do not use repeated `manage_slide_shapes` text replacements for multi-shape copy refreshes on one slide unless there is a specific reason to avoid the XML text tools.

For live slide authoring and coordinated shape layout, start with `execute_office_js` when Office.js can express the operation cleanly. Use `manage_slide_shapes` afterward only for small sparse patches.
