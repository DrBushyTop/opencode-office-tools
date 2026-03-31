# Deck Check Review Checklist

## Visual

- Text overflow or clipping
- Shapes outside slide bounds
- Misalignment or uneven spacing
- Low contrast or undersized text
- Inconsistent title/body treatment across related slides

## Structural

- Generic shape names where targeted edits or animations are likely
- Multiple concepts forced into one text box when separate shapes would be clearer
- Unnecessary ungrouped fragments or over-grouped content
- Layout mismatch that suggests a layout-based slide should be used instead

## Workflow

- Slide should be duplicated before a risky redesign
- Use returned refreshed slide ids or shape refs when available; otherwise rerun `list_slide_shapes` before the next pinpoint edit
- Native chart or layout tools would be a better fit than broad shape editing
