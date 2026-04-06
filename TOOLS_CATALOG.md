# Tools Catalog

This document lists all available tools that OpenCode can use when working with your Office documents.

## Word Tools

| Tool | Description |
|------|-------------|
| `get_document_overview` | Get a structural overview of the document including word count, heading hierarchy, table count, and list count. Use this first to understand the document structure. |
| `get_document_content` | Get the full HTML content of the Word document. |
| `get_document_part` | Read structural Word document parts by address, including `headers_footers`, `section[1]`, `section[1].header.primary`, and `table_of_contents`. |
| `get_document_section` | Read content of a specific section by heading name. More efficient than reading the entire document. |
| `set_document_content` | Replace the entire document body with new HTML content. |
| `set_document_part` | Update structural Word document parts by address for header/footer content, section page setup, and native table of contents insertion. |
| `get_selection` | Get the currently selected content as OOXML. |
| `get_selection_text` | Get the currently selected text as plain readable text. |
| `get_selection_html` | Get the currently selected content as HTML. |
| `get_document_range` | Read a generic Word target by address such as `selection`, `bookmark[Name]`, `content_control[id=12]`, `table[1]`, or `table[1].cell[2,3]`. |
| `set_document_range` | Update a generic Word target by address with HTML, text, or OOXML using replace/insert/clear. Clearing `table[n]` is rejected because it would delete the entire table. |
| `find_document_text` | Locate text without mutating the document, optionally scoped to a generic Word target address such as `selection`, `table[1]`, or `table[1].cell[2,3]`. |
| `get_document_targets` | Inspect tables, content controls, and bookmarks so later range operations can target them precisely. |
| `insert_content_at_selection` | Insert HTML content at the cursor position (before, after, or replace selection). |
| `find_and_replace` | Search and replace text with options for case sensitivity, whole word matching, and optional generic target scoping including `table[1].cell[2,3]`. |
| `insert_table` | Insert a formatted table at the cursor with header styling and grid/striped options. |
| `apply_style_to_selection` | Apply formatting to selected text (bold, italic, underline, font size, colors, highlighting). |
| `fetch_web_page` | Fetch content from a URL and convert the page to markdown through the local proxy. |

## PowerPoint Tools

| Tool | Description |
|------|-------------|
| `get_presentation_overview` | Get a quick overview of the presentation with slide count and content previews. Use this first. |
| `get_presentation_structure` | Inspect slide masters, layouts, themes, backgrounds, footer-like placeholders, and selection state; can also return structured template metadata. |
| `get_presentation_content` | Read text content from slides with support for chunked reading of large presentations. |
| `get_slide_image` | Capture a slide as a PNG image for visual inspection before making changes. |
| `add_slide_animation` | Add a slide animation through an Open XML fallback with timing control; supports motion paths, scale emphasis, and rotation, and replaces the slide in the deck. |
| `clear_slide_animations` | Remove all animations from a slide through an Open XML fallback; this replaces the slide in the deck. |
| `get_slide_animations` | Inspect the animation sequence on a slide, including effect types, targets, timing, and order. |
| `execute_office_js` | Office.js escape hatch for the live PowerPoint deck. Use only when the higher-level tools cannot express the operation cleanly. Do not use for batch shape creation or text formatting that `edit_slide_xml` or `manage_slide_shapes` can handle. |
| `get_slide_notes` | Read speaker notes by exporting slides through an Open XML fallback when the native PowerPoint API does not expose notes directly. |
| `get_slide_transition` | Inspect a slide transition through an Open XML fallback. |
| `list_slide_shapes` | List shapes on a slide and return stable shape refs you can reuse across the new PowerPoint text and chart workflow. |
| `read_slide_text` | Read the raw paragraph XML for one text shape by stable shape ref. Use this before fidelity-sensitive text edits. |
| `edit_slide_text` | Preferred text-editing tool for one existing text shape. Replaces paragraph XML in place while preserving the rest of the slide. |
| `edit_slide_xml` | General-purpose single-slide XML editor. Use it for batch text, diagrams, advanced formatting, and other OOXML edits against the exported `ppt/slides/slide1.xml` part. |
| `edit_slide_chart` | Create, update, or delete native PowerPoint charts on a slide using stable chart refs. |
| `edit_slide_master` | Update supported slide master theme colors and decorative master shapes. |
| `list_slide_layouts` | Catalog slide masters and layouts, including placeholder metadata, before creating new slides from layouts. |
| `duplicate_slide` | Duplicate a slide as a convenience wrapper for native slide duplication. |
| `create_slide_from_layout` | Create a new slide from a layout and optionally bind text, images, or tables into placeholders. |
| `manage_slide` | Create, delete, move, or clear slides with one generic slide-management tool. |
| `manage_slide_shapes` | Create, update, or delete shapes with generic geometry, styling, grouping, naming, and simple text controls. Prefer the XML text tools for rewriting copy in existing text boxes. |
| `manage_slide_media` | Insert, replace, or delete editable image shapes on a PowerPoint slide. |
| `manage_slide_table` | Create, update, or delete editable native PowerPoint tables. |
| `set_slide_notes` | Add or update speaker notes by round-tripping a slide through Open XML and replacing it in the deck; this may change slide identity. |
| `set_slide_transition` | Add, update, or clear a slide transition by round-tripping a slide through Open XML; this may change slide identity. |
| `fetch_web_page` | Fetch content from a URL and convert the page to markdown through the local proxy. |

## Excel Tools

| Tool | Description |
|------|-------------|
| `get_workbook_overview` | Get a structural overview of the workbook including sheets, visibility, protection, used ranges, tables, PivotTables, filters, frozen panes, named ranges, and chart counts. Use this first. |
| `get_workbook_info` | Get a lightweight workbook summary with worksheet names and the active sheet; prefer `get_workbook_overview` for structure. |
| `get_workbook_content` | Read cell values and formulas from a worksheet or specific range; detail mode also includes display text, number formats, validation, merged areas, and table/PivotTable overlap. |
| `set_workbook_content` | Write a 2D array of values or formulas to cells starting at a specific position, optionally clear first, and optionally create a table. |
| `get_selected_range` | Read the currently selected cells including values and formulas; detail mode also includes richer cell metadata. |
| `set_selected_range` | Write values or formulas to the currently selected range, expanding from a single selected cell when needed. |
| `find_and_replace_cells` | Search and replace text in cells with Excel's native replace behavior (ExcelApi 1.9+), preserving formulas better than value-only rewrites. |
| `manage_chart` | Create or update charts, including source data, title, type, placement, resizing, activation, and deletion. |
| `apply_cell_formatting` | Format cells with fonts, fills, borders, number formats, horizontal/vertical alignment, wrapping, merging, and row/column sizing. |
| `manage_named_range` | Create, update, rename, hide/show, or delete workbook-scoped named ranges. |
| `manage_range` | Perform generic range operations like clear, insert, delete, copy, fill, sort, and filter; provide `columnIndex` when applying filter criteria. |
| `manage_worksheet` | Create, rename, delete, move, change visibility, freeze/unfreeze, activate, protect, or unprotect worksheets. |
| `manage_table` | Create or update Excel tables, including style, totals, resizing, row appends/inserts, filter reset, conversion back to ranges, and deletion. |
| `fetch_web_page` | Fetch content from a URL and convert the page to markdown through the local proxy. |

## OneNote Tools

| Tool | Description |
|------|-------------|
| `get_notebook_overview` | Get a structural overview of the active OneNote notebook, including sections, section groups, page ids, and page client URLs. Use this first. |
| `get_page_content` | Read the active OneNote page as a summary, extracted text, or structured JSON. OneNote only exposes full page content for the active page. |
| `get_note_selection` | Read the current OneNote selection as plain text or a matrix of values. |
| `set_note_selection` | Write text, HTML, or an image to the current OneNote selection using OneNote's supported selection coercions. |
| `create_page` | Create a new page in the active section or before/after the current page, with optional initial HTML content. |
| `set_page_title` | Rename the active OneNote page. |
| `append_page_content` | Append limited supported HTML to the active OneNote page, reusing the last outline when possible. |
| `navigate_to_page` | Navigate OneNote to a target page by page id or client URL; provide exactly one target so active-page-only reads and edits can work. |
| `fetch_web_page` | Fetch content from a URL and convert the page to markdown through the local proxy. |

---

## Tool Usage Patterns

### Start with Overview Tools
Always begin by using the overview tool for your application:
- Word: `get_document_overview`
- PowerPoint: `get_presentation_overview`  
- Excel: `get_workbook_overview`
- OneNote: `get_notebook_overview`

This helps OpenCode understand your document structure before making targeted reads or edits.

### Runtime Notes
- Office tools are bundled inside the add-in's `.opencode/tools/` directory.
- The local add-in server exposes those tools to OpenCode and executes them inside the active Office host.
- Read-only inspection tools are auto-approved; mutating tools use the OpenCode permission flow.

### Surgical Edits vs Full Replacement
- **Surgical**: Use `set_document_range`, `insert_content_at_selection`, `find_document_text`, `find_and_replace`, `list_slide_shapes`, `read_slide_text`, `edit_slide_text`, `edit_slide_xml`, `edit_slide_chart`, `edit_slide_master`, and `manage_slide_shapes` for targeted changes
- **Full replacement**: Use `set_document_content`; for PowerPoint, inspect first and prefer layout-driven creation or slide-scoped native edits over broad slide regeneration

### Word: Generic Part Addressing
When working with advanced Word structure, prefer document-part addresses:
- `headers_footers` for a cross-section summary
- `section[1]` or `section[*]` for section-level page setup
- `section[1].header.primary` or `section[2].footer.firstPage` for boilerplate areas
- `table_of_contents` for native TOC insertion or inspection

### Word: Generic Target Addressing
When working with body content, prefer a small set of generic target primitives:
- `selection` for the current selection
- `bookmark[Name]` for bookmark-oriented reads and writes
- `content_control[id=12]` or `content_control[index=1]` for content-control targeting
- `table[1]` for an entire table range (requires WordApi 1.3; read, replace, or insert; clear is rejected to avoid deleting the full table)
- `table[1].cell[2,3]` for a specific table cell body (requires WordApi 1.3)

Suggested pattern:
1. Use `get_document_targets` to discover tables, content controls, and bookmarks
2. Use `get_document_range` or `find_document_text` to inspect the exact target
3. Use `set_document_range` for generic edits
4. Keep `set_document_part` for headers, footers, section setup, and native TOC work

### PowerPoint: Prefer Native Tools Before Code
Inspect first, then route to the narrowest native tool that matches the task. Use `list_slide_layouts` + `create_slide_from_layout` for layout-driven slide creation, `list_slide_shapes` + `read_slide_text`/`edit_slide_text`/`edit_slide_xml` for stable text-shape editing, `edit_slide_chart` for native charts, and `manage_slide`, `manage_slide_shapes`, `manage_slide_media`, and `manage_slide_table` for broader slide authoring. Keep `execute_office_js` as a last-resort escape hatch for operations the higher-level tools cannot express.

Suggested workflow:
1. Use `get_presentation_overview`, `get_presentation_structure`, and `list_slide_layouts` to understand the deck, slide size, theme, and available layouts.
2. Use `create_slide_from_layout` for layout-based creation, or `duplicate_slide` when you want a safe prototype branch before editing.
3. Use `list_slide_shapes` to get current shape refs before text or chart edits.
4. Prefer OOXML-first text edits for fidelity: `read_slide_text` + `edit_slide_text` for one shape, or `edit_slide_xml` for several text shapes on the same slide.
5. Use `edit_slide_xml` as the general single-slide OOXML editor for diagrams, advanced formatting, and structural XML edits on one slide.
6. If multiple existing text boxes on one slide need copy changes, do not default to one-by-one `manage_slide_shapes` updates.
7. Use `edit_slide_chart` for chart updates and `edit_slide_master` for supported master/theme changes.
8. Use `execute_office_js` only as a last resort for host-native operations the higher-level tools cannot express.
9. Keep `manage_slide_shapes`, `manage_slide_media`, and `manage_slide_table` for geometry, styling, native media, and native table operations.
7. Refresh slide ids, shape ids, or shape refs after round-trip mutations before making the next targeted edit.

### PowerPoint: Shape Design for Animation-Readiness
When building slides that will later be animated, keep shapes structured so each animatable element is a **separate shape** with a **descriptive name**:

- **One shape per animatable unit.** If bullet points should appear one by one, create each bullet as its own text box rather than a single multi-bullet shape. If a metric card has an icon and a label, keep them as separate shapes (or group them intentionally if they should animate together).
- **Use descriptive shape names.** Name shapes semantically — `"Bullet 1 - Revenue"`, `"Hero Image"`, `"Key Metric: Users"` — so the animation model can understand what each shape represents and target them precisely with `add_slide_animation`. Avoid generic names like `"TextBox 5"`.
- **Order shapes intentionally.** Shapes are animated by their index (creation order). Place shapes in the order they should naturally appear (e.g., title first, then subtitle, then content items top-to-bottom or left-to-right).
- **Review after creation.** After building a slide, verify that the shape structure matches the intended animation plan. Check that separate elements are not accidentally merged into one shape and that names are descriptive.

### Excel: Formatting After Data
When working with Excel data:
1. Use `set_workbook_content` to write data
2. Use `apply_cell_formatting` to style headers and cells
3. Use `manage_chart` to visualize or refine the data
4. Use `manage_named_range` for important data regions
5. Use `manage_range` for generic range-level cleanup, fill, sort, or filter operations

### OneNote: Active-Page-Centric Workflow
When working with OneNote:
1. Use `get_notebook_overview` to discover sections and page ids
2. Use `navigate_to_page` when the page you need is not currently active
3. Use `get_page_content` to inspect the active page before mutating it
4. Use `append_page_content`, `set_page_title`, or `set_note_selection` for edits
5. Keep HTML simple because OneNote supports only a limited subset and normalizes whitespace
