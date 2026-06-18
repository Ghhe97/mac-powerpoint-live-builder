# Mac PowerPoint Live Workflow Reference

## Live Sequence Contract

A deck build should be a list of small actions:

```json
[
  {"type": "create_presentation"},
  {"type": "add_slide", "layout": "blank"},
  {"type": "focus_slide", "slide_index": 1},
  {"type": "shape", "slide_index": 1, "shape_name": "s1_bg", "...": "..."},
  {"type": "text", "slide_index": 1, "shape_name": "s1_title", "text": "..."}
]
```

Use `pptx_run_live_sequence(steps=...)` for bulk live builds. For heavy decks,
batch by slide if one call becomes slow or hard to debug.

## Mac-Specific Pitfalls

- New presentations can start with zero slides. Always call `add_slide` before
  targeting slide 1.
- Newly created PowerPoint shape references can be unstable. The MCP server should
  name each shape, refetch it by `name`, then style it. If direct styling appears
  ignored, verify the server does this.
- PowerPoint color properties expect AppleScript 16-bit RGB channel lists. A robust
  server should convert `#RRGGBB` to `{R*257, G*257, B*257}`.
- PowerPoint is sandboxed. Thumbnail/PDF export should use the PowerPoint container
  temp path internally, then copy outputs to the requested path.
- macOS may prompt for Automation permission. If tools fail with permission or
  AppleEvent errors, ask the user to allow the launcher to control PowerPoint.

## Dense Deck Layout Patterns

Use 16:9 coordinates in points. A common canvas is `960 x 540`.

- Header: small section label, title, one-line thesis.
- KPI strip: 4-5 cards, 7-10 pt labels, 18-24 pt values.
- Main body: split into 2 panels or 1 large matrix + 1 insight callout.
- Footer: 6-7 pt source line with source family and date/definition caveats.
- Minimum practical font size for dense slides: 6.2 pt for footers, 7-8 pt for
  matrix body, 10-12 pt for panel titles, 22-28 pt for slide titles.

## Visual QA Checklist

Inspect each exported thumbnail:

- Slide is nonblank and saved as editable `.pptx`.
- No obvious overlap, clipping, or text outside panels.
- Important numbers and headings remain readable at overview scale.
- Colors match expected theme; sample pixels if a critical brand/research color
  looks wrong.
- All data-heavy slides include sources or definition caveats.
- Deck overview shows consistent hierarchy across slides.

## Suggested Output Files

For a user-facing deliverable, save:

- `outputs/<deck-name>.pptx`
- `outputs/<deck-name>-slide1.png`, one per slide
- `outputs/<deck-name>-overview.png`

If the user asks for iteration, update the same deck or create a versioned file
depending on whether they want history preserved.
