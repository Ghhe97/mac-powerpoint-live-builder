---
name: mac-powerpoint-live-builder
description: Create, edit, visually QA, and save editable PowerPoint decks by driving the live Microsoft PowerPoint for Mac window through a local stdio MCP server. Use when the user wants Codex to generate PPT/PPTX slides on macOS with the presentation visibly building in the PowerPoint app, especially for dense research decks, structured business presentations, slide-by-slide iteration, or workflows that require exported slide thumbnails for visual verification.
---

# Mac PowerPoint Live Builder

Use this skill when the output must be an editable `.pptx` created in the running
Microsoft PowerPoint for Mac app, not a static image or file-only OOXML rewrite.

The skill includes a bundled local MCP server installer. It does not require an
API token; it does require macOS Automation permission for the process that
launches AppleScript. Some Agent products run MCP inside a sandbox that cannot
send AppleEvents even when the app appears in macOS Automation settings. For
those products, use bridge mode: run `scripts/powerpoint_bridge.py` outside the
Agent sandbox, then configure MCP with `POWERPOINT_LIVE_BRIDGE_URL` and
`POWERPOINT_LIVE_BRIDGE_TOKEN_FILE`.

## Workflow

1. Check for live `pptx_*` MCP tools in the current Agent runtime. Distinguish:
   installed MCP server, server tool inventory, and tools actually mounted in the
   current Agent session. If runtime tools are missing, read `references/install.md`;
   do not treat installer self-check success as proof that this session can call MCP.
2. Confirm the user wants a live Mac PowerPoint build, then infer slide count,
   audience, language, density, and output tone from the prompt.
3. If the request depends on current market, legal, policy, financial, or product
   facts, browse and cite current sources before building.
4. Before a large build, prefer a real smoke check through MCP when available:
   `pptx_create_presentation`, then close the empty presentation without saving. If
   only the bundled checker is available, run it with `--smoke-powerpoint`.
5. Draft a slide plan first: one concise title, a high-density layout model, core
   facts, and source footer for each slide.
6. Generate a live sequence that starts with `create_presentation`, then explicitly
   adds one blank slide before building each slide. Some PowerPoint for Mac builds
   create new presentations with zero slides.
7. Use named shapes/text boxes for every element. Names are required for later
   edits and for reliable styling after creation.
8. Execute the sequence through MCP so the PowerPoint window visibly mutates.
9. Save the deck, export each slide thumbnail, and export an overview image.
10. Inspect thumbnails. Fix overlap, illegible text, weak contrast, bad colors,
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
- If live PowerPoint control fails and another method is used, explicitly label the
  output as a non-live fallback. Never imply PowerPoint visibly built the deck unless
  MCP/AppleScript actually did so in the running app.

## Tooling

Use the user's configured MCP tools when available. If the tools are missing,
install from the bundled vendor package with:

```bash
python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-codex-config
```

If the current Agent product is not Codex, run the installer without
`--write-codex-config` and copy the printed stdio JSON or WorkBuddy block into
that product's MCP settings. Restart the Agent if it loads MCP servers only at
startup.

For WorkBuddy or another sandboxed Agent that fails live smoke tests with
`-10004` even after Automation permission is enabled, use bridge mode:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/start_bridge.command
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-workbuddy-config --bridge-mode
```

Then restart WorkBuddy and run:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint --bridge-mode
```

If bridge mode reports `osascript timed out after 60s`, the bridge process is
running but its launcher is probably not allowed to control PowerPoint. For a
bridge started by `start_bridge.command`, enable macOS Automation permission for
Terminal to control Microsoft PowerPoint, then restart the bridge.

`scripts/check_pptx_mcp.py` verifies that a server executable starts and exposes
the expected `pptx_*` live-build tools. Add `--smoke-powerpoint` only when you
want it to create and close a tiny PowerPoint presentation through MCP.

For diagnostics:

```bash
python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint
```

If `osascript` reports `-1708`, suspect foreground activation and retry with the
updated server that wraps `activate` in `try/end try`. If it reports `-10004` or
`not authorized`, check macOS Automation permission for the app that launched
AppleScript. If permission is already enabled and the Agent is sandboxed, switch
to bridge mode instead of reporting live success.

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
