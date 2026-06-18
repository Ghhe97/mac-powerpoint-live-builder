---
name: mac-powerpoint-live-builder
description: Create, edit, visually QA, and save editable PowerPoint decks by driving the live Microsoft PowerPoint for Mac window through a local stdio MCP server. Use when the user wants Codex to generate PPT/PPTX slides on macOS with the presentation visibly building in the PowerPoint app, especially for dense research decks, structured business presentations, slide-by-slide iteration, or workflows that require exported slide thumbnails for visual verification.
---

# Mac PowerPoint Live Builder

Use this skill when the output must be an editable `.pptx` created in the running
Microsoft PowerPoint for Mac app, not a static image or file-only OOXML rewrite.

The skill includes a bundled local MCP server installer. It does not require an
API token; it does require macOS Automation permission for the process that
launches the MCP server.

## Workflow

1. Check for live `pptx_*` MCP tools in the current runtime. If they are missing,
   read `references/install.md` and run the bundled installer/check flow before
   attempting to create slides.
2. Confirm the user wants a live Mac PowerPoint build, then infer slide count,
   audience, language, density, and output tone from the prompt.
3. If the request depends on current market, legal, policy, financial, or product
   facts, browse and cite current sources before building.
4. Draft a slide plan first: one concise title, a high-density layout model, core
   facts, and source footer for each slide.
5. Generate a live sequence that starts with `create_presentation`, then explicitly
   adds one blank slide before building each slide. Some PowerPoint for Mac builds
   create new presentations with zero slides.
6. Use named shapes/text boxes for every element. Names are required for later
   edits and for reliable styling after creation.
7. Execute the sequence through MCP so the PowerPoint window visibly mutates.
8. Save the deck, export each slide thumbnail, and export an overview image.
9. Inspect thumbnails. Fix overlap, illegible text, weak contrast, bad colors,
   missing sources, and low information density before final delivery.

For exact sequence patterns, QA checks, and Mac-specific pitfalls, read
`references/workflow.md` before implementing a deck.

## Design Rules

- Prefer editable PowerPoint text boxes and shapes over rasterized slide images.
- Use dark or light professional systems with restrained accent colors. Keep source
  footers small but present on research slides.
- For high-density research slides, use KPI strips, matrices, small multiples,
  timeline rows, risk grids, and compact callouts instead of paragraphs.
- Keep the first visual pass truthful: if data is approximate, label it as
  approximate; if sources differ by definition, note the definition.
- Do not claim success until exported thumbnails have been visually inspected.

## Tooling

Use the user's configured MCP tools when available. If the tools are missing,
install from the bundled vendor package with:

```bash
python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-codex-config
```

If the current Agent product is not Codex, run the installer without
`--write-codex-config` and copy the printed stdio command into that product's MCP
settings.

`scripts/check_pptx_mcp.py` verifies that a server executable starts and exposes
the expected `pptx_*` live-build tools.

Expected core tools:

- `pptx_create_presentation`
- `pptx_add_slide`
- `pptx_focus_slide`
- `pptx_add_shape`
- `pptx_add_text_box`
- `pptx_run_live_sequence`
- `pptx_save_presentation`
- `pptx_get_slide_thumbnail`
- `pptx_get_deck_overview`

## Delivery

Return links to the saved `.pptx` and preview PNGs. Mention validation performed:
tool listing, live sequence completion, save path, thumbnail export, and any
residual caveats.
