"""Local MCP server for live Microsoft PowerPoint control on macOS.

This server speaks AppleScript that matches PowerPoint for Mac's dictionary and
drives the running PowerPoint app directly. It is intended for Agent workflows
that need editable `.pptx` output plus visible slide-by-slide construction.

Implementation notes:

  - The shapes collection is *not* safe for `repeat with X in <collection>` iteration —
    it hangs PowerPoint indefinitely. We use `repeat with i from 1 to N` everywhere.
  - The geometry property is `left position`, not `left` (despite `left` showing up in
    other PowerPoint AppleScript dialects).
  - `placeholder index` / `placeholder type` are reserved-keyword conflicts; address with
    pipe-quotes `|placeholder index|`. But: in this corporate template, shapes are
    freeform (no placeholder format at all), so the tools that need OOXML idx fall back
    to addressing by shape `name`.
  - `slide layout N of slide master` is not a valid reference form. The slide property
    is `layout` (one identifier), taking enum constants like `slide layout blank`.
  - `duplicate slide N of activePres` returns Parameter error -50 unconditionally;
    PowerPoint Mac AppleScript does not support slide duplication. Use python-pptx.
  - `replace tr what ... replacement ...` is not a valid PowerPoint AppleScript verb;
    we read-modify-write in Python (which loses styled runs).
  - PowerPoint is sandboxed; writing outside `~/Library/Containers/com.microsoft.Powerpoint/`
    triggers TCC prompts and may fail silently. Temporary I/O lives inside the
    container's `Data/tmp/` directory.
"""

from __future__ import annotations

import glob
import io
import json
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.request
from typing import Any

from mcp.server.fastmcp import FastMCP, Image
from PIL import Image as PILImage, ImageDraw, ImageFont

logger = logging.getLogger("pptx_mcp")

mcp = FastMCP("powerpoint-live-mcp")

# Sandbox-writable directory for any artifacts (PDF exports, thumbnails) that the
# PowerPoint process needs to produce. Writing outside this directory triggers TCC
# prompts and may fail silently.
POWERPOINT_SANDBOX_TMP = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/pptx-live-mcp"
)

BRIDGE_URL_ENV = "POWERPOINT_LIVE_BRIDGE_URL"
BRIDGE_TOKEN_ENV = "POWERPOINT_LIVE_BRIDGE_TOKEN"
BRIDGE_TOKEN_FILE_ENV = "POWERPOINT_LIVE_BRIDGE_TOKEN_FILE"


# --- AppleScript runner ---------------------------------------------------

def _escape_applescript_string(s: str) -> str:
    """Escape a Python string so it can be embedded inside an AppleScript string literal."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _osascript_error_message(returncode: int, detail: str) -> str:
    hint = ""
    if "-1708" in detail:
        hint = (
            " Hint: PowerPoint rejected 'activate'. This usually means the "
            "launcher cannot bring PowerPoint to the foreground; wrap activate "
            "in try/end try or continue without foreground focus."
        )
    elif returncode == 124 or "timed out" in detail.lower():
        bridge_url = os.environ.get(BRIDGE_URL_ENV, "").strip()
        if bridge_url:
            hint = (
                " Hint: the PowerPoint bridge is reachable, but the app that "
                "started the bridge may not be allowed to control Microsoft "
                "PowerPoint. If you launched the bridge from Terminal, enable "
                "System Settings > Privacy & Security > Automation > Terminal > "
                "Microsoft PowerPoint, then restart the bridge."
            )
        else:
            hint = (
                " Hint: osascript timed out while talking to PowerPoint. Check "
                "macOS Automation permission for the launching app and make sure "
                "PowerPoint is responsive."
            )
    elif "-10004" in detail or "not authorized" in detail.lower() or "权限" in detail:
        bridge_url = os.environ.get(BRIDGE_URL_ENV, "").strip()
        if bridge_url:
            hint = (
                " Hint: macOS Automation blocked AppleEvents even through the "
                "PowerPoint bridge. Start the bridge from Terminal/Codex and allow "
                "that app to control Microsoft PowerPoint."
            )
        else:
            hint = (
                " Hint: macOS Automation may be blocking AppleEvents. Check "
                "System Settings > Privacy & Security > Automation for the app "
                "that launched this MCP server, or use WorkBuddy bridge mode."
            )
    return f"osascript failed (exit {returncode}): {detail}{hint}"


def _bridge_token() -> str:
    token = os.environ.get(BRIDGE_TOKEN_ENV, "").strip()
    if token:
        return token
    token_file = os.environ.get(BRIDGE_TOKEN_FILE_ENV, "").strip()
    if token_file:
        try:
            with open(os.path.expanduser(token_file), encoding="utf-8") as fh:
                return fh.read().strip()
        except OSError:
            return ""
    return ""


def _run_osascript_via_bridge(script: str, timeout: int) -> str:
    bridge_url = os.environ.get(BRIDGE_URL_ENV, "").strip().rstrip("/")
    if not bridge_url:
        raise RuntimeError("bridge URL is empty")
    payload = json.dumps({"script": script, "timeout": timeout}).encode("utf-8")
    body = ""
    last_http_error: tuple[int, str] | None = None
    for attempt in range(1, 4):
        request = urllib.request.Request(
            f"{bridge_url}/run-osascript",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {_bridge_token()}",
                # Some Agent sandboxes proxy localhost HTTP. Closing each request
                # avoids stale proxy connections that can drop POST bodies.
                "Connection": "close",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout + 10) as response:
                body = response.read().decode("utf-8")
            break
        except urllib.error.HTTPError as e:
            error_body = e.read().decode("utf-8", errors="replace")
            last_http_error = (e.code, error_body)
            retryable_empty_body = (
                e.code == 400
                and ("empty JSON body" in error_body or '"body_preview": ""' in error_body)
                and attempt < 3
            )
            if retryable_empty_body:
                time.sleep(0.35 * attempt)
                continue
            raise RuntimeError(f"PowerPoint bridge HTTP {e.code}: {error_body}") from e
        except urllib.error.URLError as e:
            if attempt < 3:
                time.sleep(0.35 * attempt)
                continue
            raise RuntimeError(
                f"PowerPoint bridge is not reachable at {bridge_url}. Start the bridge outside the Agent sandbox."
            ) from e
    else:
        if last_http_error:
            code, error_body = last_http_error
            raise RuntimeError(f"PowerPoint bridge HTTP {code}: {error_body}")
        raise RuntimeError(f"PowerPoint bridge is not reachable at {bridge_url}.")
    try:
        result = json.loads(body)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"PowerPoint bridge returned invalid JSON: {body[:500]}") from e
    if not result.get("ok"):
        detail = (result.get("stderr") or result.get("stdout") or result.get("error") or "").strip()
        returncode = int(result.get("returncode", 1))
        raise RuntimeError(_osascript_error_message(returncode, detail))
    return str(result.get("stdout", "")).strip()


def _run_osascript(script: str, timeout: int = 60) -> str:
    """Run an AppleScript via osascript and return stdout."""
    logger.debug("Running AppleScript:\n%s", script)
    if os.environ.get(BRIDGE_URL_ENV, "").strip():
        return _run_osascript_via_bridge(script, timeout)
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"osascript timed out after {timeout}s") from e
    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(_osascript_error_message(result.returncode, detail))
    return result.stdout.strip()


def _sandbox_tmp_dir(prefix: str) -> str:
    os.makedirs(POWERPOINT_SANDBOX_TMP, exist_ok=True)
    return tempfile.mkdtemp(prefix=prefix, dir=POWERPOINT_SANDBOX_TMP)


def _load_label_font(size: int):
    """Try macOS system TTFs first, fall back to Pillow's bitmap default."""
    for ttf in (
        "/System/Library/Fonts/SFNSRounded.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ):
        try:
            return ImageFont.truetype(ttf, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def _compose_labeled_cell(
    im: "PILImage.Image",
    slide_no: int,
    label_height: int = 36,
    font_size: int = 22,
    border_color: tuple[int, int, int] = (210, 210, 210),
) -> "PILImage.Image":
    """Return a NEW image: `slide N` label on top, then the slide thumbnail with a thin
    light border. White background. Matches the deck-overview reference layout (clean
    sans-serif label above the slide, no overlay badges).
    """
    font = _load_label_font(font_size)
    label = f"slide {slide_no}"
    cell_w = im.width
    cell_h = label_height + im.height

    cell = PILImage.new("RGB", (cell_w, cell_h), (255, 255, 255))
    draw = ImageDraw.Draw(cell)
    # Label baseline ~6 px above the slide.
    try:
        bbox = draw.textbbox((0, 0), label, font=font)
        th = bbox[3] - bbox[1]
    except AttributeError:
        _, th = draw.textsize(label, font=font)
    label_y = max(0, (label_height - th) // 2 - 2)
    draw.text((2, label_y), label, fill=(40, 40, 40), font=font)

    # Paste thumbnail; outline it with a 1px light border so blank/white slides
    # have a visible edge against the white canvas.
    cell.paste(im.convert("RGB"), (0, label_height))
    draw.rectangle(
        (0, label_height, cell_w - 1, cell_h - 1),
        outline=border_color, width=1,
    )
    return cell


def _get_slide_counts_and_hidden() -> tuple[int, set[int], list[int]]:
    """Return (total_slides, hidden_set, visible_indices) for the active presentation.

    `print hidden slides` print-option is read-only with respect to `save as PDF` —
    PowerPoint always excludes hidden slides from the PDF export regardless of the
    flag. So we query the hidden set ourselves and map PDF pages ↔ PowerPoint slide
    indices in Python.
    """
    script = '''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set total to count of slides of p
    set hiddenList to {}
    repeat with i from 1 to total
        try
            if hidden of slide show transition of slide i of p then
                set end of hiddenList to i
            end if
        end try
    end repeat
    set AppleScript's text item delimiters to ","
    set h to hiddenList as text
    set AppleScript's text item delimiters to ""
    return (total as text) & "|" & h
end tell
'''
    out = _run_osascript(script)
    total_str, _, hidden_str = out.partition("|")
    try:
        total = int(total_str)
    except ValueError:
        raise RuntimeError(f"Couldn't parse slide count: {out!r}")
    hidden = set()
    for piece in (hidden_str or "").split(","):
        piece = piece.strip()
        if piece.isdigit():
            hidden.add(int(piece))
    visible = [i for i in range(1, total + 1) if i not in hidden]
    return total, hidden, visible


def _save_png_to_path(png_bytes: bytes, save_to_path: str | None) -> str | None:
    """Optionally persist a PNG to a caller-specified path. Returns the absolute path
    written (or None if save_to_path was empty)."""
    if not save_to_path:
        return None
    target = os.path.abspath(os.path.expanduser(save_to_path.strip()))
    parent = os.path.dirname(target)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(target, "wb") as fh:
        fh.write(png_bytes)
    logger.info("Saved PNG (%d bytes) to %s", len(png_bytes), target)
    return target


_ENUM_FRAGMENT_RE = re.compile(r"^[a-z0-9][a-z0-9 ]*[a-z0-9]$|^[a-z0-9]$")
_HEX_COLOR_RE = re.compile(r"^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")


def _safe_enum_fragment(value: str, *, label: str) -> str:
    """Return a conservative AppleScript enum suffix such as `blank` or `rounded rectangle`.

    AppleScript enum names are inserted as identifiers, not quoted strings, so keep this
    intentionally narrow. This also prevents model-supplied values from becoming script.
    """
    normalized = value.strip().lower().replace("_", " ").replace("-", " ")
    normalized = re.sub(r"\s+", " ", normalized)
    if not normalized or not _ENUM_FRAGMENT_RE.match(normalized):
        raise ValueError(f"Invalid {label}: {value!r}")
    return normalized


def _hex_to_rgb(color: str | None, *, label: str = "color") -> tuple[int, int, int] | None:
    """Parse #RRGGBB / #RGB into a PowerPoint AppleScript RGB triplet."""
    if color is None:
        return None
    raw = color.strip()
    if raw.lower() in {"", "none", "transparent"}:
        return None
    match = _HEX_COLOR_RE.match(raw)
    if not match:
        raise ValueError(f"Invalid {label}: {color!r}; expected #RRGGBB, #RGB, none, or transparent.")
    hex_part = match.group(1)
    if len(hex_part) == 3:
        hex_part = "".join(ch * 2 for ch in hex_part)
    return (
        int(hex_part[0:2], 16),
        int(hex_part[2:4], 16),
        int(hex_part[4:6], 16),
    )


def _as_color_literal(color: str | None, *, label: str = "color") -> str | None:
    rgb = _hex_to_rgb(color, label=label)
    if rgb is None:
        return None
    # PowerPoint for Mac's sdef declares `fore color`/`back color` as an
    # integer list. In practice it expects the standard AppleScript RGB range:
    # three 16-bit channels, not CSS-style 0..255 channels.
    rgb16 = tuple(part * 257 for part in rgb)
    return "{" + ", ".join(str(part) for part in rgb16) + "}"


def _required_color_literal(color: str | None, *, label: str = "color") -> str:
    literal = _as_color_literal(color, label=label)
    if literal is None:
        raise ValueError(f"{label} must be a visible color; use #RRGGBB or #RGB.")
    return literal


def _normalize_transparency(value: float) -> float:
    """Accept either 0..1 or 0..100 transparency and normalize to PowerPoint's 0..1."""
    raw = float(value)
    if raw < 0:
        raise ValueError("transparency cannot be negative")
    if raw > 1:
        raw = raw / 100.0
    return min(raw, 1.0)


def _maybe_live_delay(delay_seconds: float | None) -> None:
    if delay_seconds is None:
        return
    delay = float(delay_seconds)
    if delay <= 0:
        return
    # Keep a runaway JSON script from sleeping for minutes per element.
    time.sleep(min(delay, 10.0))


_AUTOSHAPE_TYPES: dict[str, str] = {
    "rectangle": "autoshape rectangle",
    "rect": "autoshape rectangle",
    "rounded rectangle": "autoshape rounded rectangle",
    "round rect": "autoshape rounded rectangle",
    "oval": "autoshape oval",
    "ellipse": "autoshape oval",
    "circle": "autoshape oval",
    "triangle": "autoshape isosceles triangle",
    "right triangle": "autoshape right triangle",
    "diamond": "autoshape diamond",
    "parallelogram": "autoshape parallelogram",
    "trapezoid": "autoshape trapezoid",
    "hexagon": "autoshape hexagon",
    "pentagon": "autoshape regular pentagon",
    "chevron": "autoshape chevron",
    "right arrow": "autoshape right arrow",
    "left arrow": "autoshape left arrow",
    "up arrow": "autoshape up arrow",
    "down arrow": "autoshape down arrow",
    "line": "autoshape rectangle",
}


def _autoshape_enum(shape_type: str) -> str:
    safe = _safe_enum_fragment(shape_type, label="shape_type")
    if safe not in _AUTOSHAPE_TYPES:
        known = ", ".join(sorted(_AUTOSHAPE_TYPES))
        raise ValueError(f"Unsupported shape_type {shape_type!r}. Known values: {known}")
    return _AUTOSHAPE_TYPES[safe]


_TEXT_ALIGNMENTS = {
    "left": "paragraph align left",
    "center": "paragraph align center",
    "centered": "paragraph align center",
    "right": "paragraph align right",
    "justify": "paragraph align justify",
    "distributed": "paragraph align distribute",
}


def _paragraph_alignment_enum(align: str) -> str:
    safe = _safe_enum_fragment(align, label="align")
    if safe not in _TEXT_ALIGNMENTS:
        known = ", ".join(sorted(_TEXT_ALIGNMENTS))
        raise ValueError(f"Unsupported align {align!r}. Known values: {known}")
    return _TEXT_ALIGNMENTS[safe]


_VERTICAL_ANCHORS = {
    "top": "anchor top",
    "middle": "anchor middle",
    "center": "anchor middle",
    "bottom": "anchor bottom",
}


def _vertical_anchor_enum(anchor: str) -> str:
    safe = _safe_enum_fragment(anchor, label="vertical_anchor")
    if safe not in _VERTICAL_ANCHORS:
        known = ", ".join(sorted(_VERTICAL_ANCHORS))
        raise ValueError(f"Unsupported vertical_anchor {anchor!r}. Known values: {known}")
    return _VERTICAL_ANCHORS[safe]


_Z_ORDER_COMMANDS = {
    "front": "bring shape to front",
    "bring to front": "bring shape to front",
    "forward": "bring shape forward",
    "bring forward": "bring shape forward",
    "back": "send shape to back",
    "send to back": "send shape to back",
    "backward": "send shape backward",
    "send backward": "send shape backward",
}


def _z_order_enum(action: str) -> str:
    safe = _safe_enum_fragment(action, label="z_order")
    if safe not in _Z_ORDER_COMMANDS:
        known = ", ".join(sorted(_Z_ORDER_COMMANDS))
        raise ValueError(f"Unsupported z_order {action!r}. Known values: {known}")
    return _Z_ORDER_COMMANDS[safe]


def _style_shape_script_lines(
    shape_ref: str,
    *,
    fill_color: str | None,
    fill_transparency: float,
    line_color: str | None,
    line_weight: float,
    line_transparency: float,
) -> list[str]:
    lines: list[str] = []
    fill_literal = _as_color_literal(fill_color, label="fill_color")
    if fill_literal is None:
        lines.extend([
            "try",
            f"    set visible of fill format of {shape_ref} to false",
            "end try",
        ])
    else:
        lines.extend([
            "try",
            f"    set visible of fill format of {shape_ref} to true",
            f"    solid fill format of {shape_ref}",
            f"    set fore color of fill format of {shape_ref} to {fill_literal}",
            f"    set transparency of fill format of {shape_ref} to {_normalize_transparency(fill_transparency)}",
            "end try",
        ])

    line_literal = _as_color_literal(line_color, label="line_color")
    if line_literal is None or float(line_weight) <= 0:
        lines.extend([
            "try",
            f"    set line weight of line format of {shape_ref} to 0",
            f"    set transparency of line format of {shape_ref} to 1.0",
            "end try",
        ])
    else:
        lines.extend([
            "try",
            f"    set fore color of line format of {shape_ref} to {line_literal}",
            f"    set line weight of line format of {shape_ref} to {float(line_weight)}",
            f"    set transparency of line format of {shape_ref} to {_normalize_transparency(line_transparency)}",
            "end try",
        ])
    return lines


def _set_shape_name_line(shape_ref: str, shape_name: str | None) -> str:
    if not shape_name:
        return ""
    safe_name = _escape_applescript_string(shape_name)
    return f'set name of {shape_ref} to "{safe_name}"'


def _generated_shape_name(prefix: str) -> str:
    return f"{prefix}_{time.time_ns()}"


def _pdftoppm_binary() -> str:
    for candidate in (
        "/opt/homebrew/bin/pdftoppm",
        "/usr/local/bin/pdftoppm",
    ):
        if os.path.exists(candidate):
            return candidate
    found = shutil.which("pdftoppm")
    if found:
        return found
    raise RuntimeError(
        "pdftoppm was not found. Install poppler with Homebrew or add a directory "
        "containing pdftoppm to the MCP server PATH."
    )


# --- Tools: the broken upstream handles, fixed ----------------------------

@mcp.tool()
def pptx_add_slide(layout: str = "blank", position: int | None = None) -> dict[str, Any]:
    """Add a new slide to the active presentation with one of PowerPoint's built-in layouts.

    Use this instead of upstream `add_slide` — that one is broken (#20473).

    Note: `layout` is a *built-in* PowerPoint enum, NOT a corporate template's custom
    layout. PowerPoint Mac AppleScript exposes ~30 standard slide layouts (blank, title,
    title only, text, chart, etc.). Setting a custom layout from a corporate template
    via AppleScript is not exposed — you'd need to use python-pptx + slide layout XML
    references for that.

    Common enum values: blank, title, title only, text, two column text, chart,
    text and chart, organization chart, table, vertical title and text,
    title and content, section header, two content, comparison.

    Args:
        layout: Built-in layout name without the `slide layout ` prefix
                (so "blank" → AppleScript `slide layout blank`).
        position: 1-based index to insert at. If omitted, the slide is appended at the end.

    Returns:
        dict with `slide_index` of the new slide and the layout that was applied.
    """
    safe_layout = _safe_enum_fragment(layout, label="layout")  # AppleScript enums are lowercase

    if position is None:
        insertion = "make new slide at end of p"
    else:
        insertion = f"make new slide at after slide {int(position) - 1} of p" if int(position) > 1 \
            else "make new slide at before slide 1 of p"

    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set newSlide to {insertion}
    set layout of newSlide to slide layout {safe_layout}
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {
        "slide_index": int(out) if out.isdigit() else out,
        "layout_applied": f"slide layout {safe_layout}",
    }


@mcp.tool()
def pptx_add_slide_from_template(source_slide_index: int, position: int | None = None) -> dict[str, Any]:
    """Clone an existing slide via PowerPoint's clipboard — full visual copy.

    The new slide inherits every style detail from the source: layout placeholders,
    fonts, theme colors, AND any freeform decorative shapes the author drew on top
    (rectangles, custom icons, hand-placed text boxes — all of it).

    Why this works: PowerPoint Mac's `duplicate` verb returns -50 for slides, but the
    pair `copy object slide N` + `paste object` (both inside `tell active presentation`)
    rides PowerPoint's clipboard machinery and creates a true clone appended at the end.
    Then `move` repositions it.

    Args:
        source_slide_index: 1-based index of the slide to clone.
        position: 1-based target index for the clone. If omitted, the clone stays at
            the end of the deck.

    Returns:
        dict with `slide_index` of the new slide.
    """
    move_clause = ""
    if position is not None:
        # `paste object` appends to the end, then we move into place.
        if int(position) <= 1:
            move_clause = "move newSlide to before slide 1 of activePres"
        else:
            move_clause = f"move newSlide to before slide {int(position)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    tell activePres
        copy object slide {int(source_slide_index)}
        paste object
    end tell
    set newSlide to slide (count of slides of activePres) of activePres
    {move_clause}
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


def _image_pixel_size(path: str) -> tuple[int, int] | None:
    """Return (width_px, height_px) for an image, or None if it can't be probed."""
    try:
        result = subprocess.run(
            ["sips", "-g", "pixelWidth", "-g", "pixelHeight", path],
            capture_output=True, text=True, check=False, timeout=5,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None
    if result.returncode != 0:
        return None
    w = h = None
    for line in result.stdout.splitlines():
        parts = line.strip().split(":")
        if len(parts) == 2:
            key, val = parts[0].strip(), parts[1].strip()
            if key == "pixelWidth" and val.isdigit():
                w = int(val)
            elif key == "pixelHeight" and val.isdigit():
                h = int(val)
    if w and h:
        return (w, h)
    return None


# Slide canvas defaults for the auto-fit policy. Most decks are 1920×1080 points,
# but we use a more conservative 800×600 cap so a single image doesn't take over.
_AUTOFIT_MAX_W = 800.0
_AUTOFIT_MAX_H = 600.0


@mcp.tool()
def pptx_insert_image(
    slide_index: int,
    image_path: str,
    left: float = 50,
    top: float = 50,
    width: float = 0,
    height: float = 0,
) -> dict[str, Any]:
    """Insert an image into a slide, sized from the image's intrinsic dimensions.

    Sizing policy (live-tested):
      - If both `width` and `height` are positive: use them as-is (caller knows best).
      - If only one is positive: compute the other from the image's aspect ratio.
      - If both are 0: read the image's pixel dimensions via `sips`, fit into 800×600
        points while preserving aspect ratio. Pixels map 1:1 to points first, then
        scale down only if the image is larger than the cap.

    Implementation note: PowerPoint Mac's AppleScript dictionary does NOT expose `add
    picture` — the `picture` class is read-only, and `make new picture` returns -50.
    The dictionary-correct path is two steps: (1) `make new shape` with a rectangle of
    the target geometry, (2) `user picture <shape> picture file "<path>"` to set its
    fill to the image. Visually identical to a "real" picture shape; in OOXML this is
    a rectangle with `<a:blipFill>` instead of `<p:pic>`.

    Note: on first call PowerPoint may show a one-time TCC permission dialog asking
    whether to allow access to the source image. Approve it; future calls run silently.

    Args:
        slide_index: 1-based index of the target slide.
        image_path: Absolute POSIX path to the image file.
        left, top: Position in points (PowerPoint's native unit, 1 inch = 72 points).
        width, height: Size in points. 0 means "auto from image dimensions". If only
            one of them is given, the other is computed from the image's aspect ratio.

    Returns:
        dict with `shape_name`, `width`, `height`, and `applied_policy` (one of
        "as-given", "aspect-from-width", "aspect-from-height", "autofit-from-image").
    """
    safe_path = _escape_applescript_string(image_path)

    px = _image_pixel_size(image_path)
    aspect = (px[0] / px[1]) if (px and px[1]) else None

    final_w = float(width)
    final_h = float(height)
    policy = "as-given"

    if final_w > 0 and final_h > 0:
        policy = "as-given"
    elif final_w > 0 and final_h <= 0 and aspect:
        final_h = final_w / aspect
        policy = "aspect-from-width"
    elif final_h > 0 and final_w <= 0 and aspect:
        final_w = final_h * aspect
        policy = "aspect-from-height"
    elif aspect:
        # Both zero — autofit. Start from intrinsic pixels as points; scale down
        # only if larger than the cap.
        w_pt = float(px[0])
        h_pt = float(px[1])
        scale = min(_AUTOFIT_MAX_W / w_pt, _AUTOFIT_MAX_H / h_pt, 1.0)
        final_w = w_pt * scale
        final_h = h_pt * scale
        policy = "autofit-from-image"
    else:
        # No pixel info available and caller gave nothing — fall back to a safe square.
        final_w = 400.0
        final_h = 300.0
        policy = "fallback-default"

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set newShape to make new shape at sl with properties {{left position:{float(left)}, top:{float(top)}, width:{float(final_w)}, height:{float(final_h)}, auto shape type:autoshape rectangle}}
    user picture newShape picture file "{safe_path}"
    return name of newShape
end tell
'''
    out = _run_osascript(script)
    return {
        "shape_name": out,
        "width": round(final_w, 2),
        "height": round(final_h, 2),
        "applied_policy": policy,
        "intrinsic_pixels": list(px) if px else None,
    }


@mcp.tool()
def pptx_focus_slide(slide_index: int) -> dict[str, Any]:
    """Bring PowerPoint to the front and show a specific slide in the active window.

    Use before live-building on a new slide so the human can watch the correct canvas.
    """
    script = f'''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
    go to slide (view of active window) number {int(slide_index)}
    return slide index of slide {int(slide_index)} of active presentation
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


@mcp.tool()
def pptx_live_pause(seconds: float = 0.8) -> dict[str, Any]:
    """Pause between live-build steps so the PowerPoint window visibly catches up."""
    delay = max(0.0, min(float(seconds), 10.0))
    time.sleep(delay)
    return {"paused_seconds": delay}


@mcp.tool()
def pptx_add_shape(
    slide_index: int,
    shape_type: str = "rectangle",
    left: float = 0,
    top: float = 0,
    width: float = 100,
    height: float = 100,
    fill_color: str | None = "#FFFFFF",
    fill_transparency: float = 0,
    line_color: str | None = "none",
    line_weight: float = 0,
    line_transparency: float = 0,
    shape_name: str | None = None,
    z_order: str | None = None,
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Add a geometric shape to a slide for live, from-scratch PPT construction.

    Coordinates and dimensions are PowerPoint points. `fill_color` / `line_color`
    accept `#RRGGBB`, `#RGB`, `none`, or `transparent`. `shape_type` is a conservative
    set of Mac-safe autoshape names: rectangle, rounded rectangle, oval, triangle,
    diamond, arrow variants, and a few simple diagram shapes.
    """
    shape_enum = _autoshape_enum(shape_type)
    resolved_shape_name = shape_name or _generated_shape_name("pptx_mcp_shape")
    style_lines = _style_shape_script_lines(
        "targetShape",
        fill_color=fill_color,
        fill_transparency=fill_transparency,
        line_color=line_color,
        line_weight=line_weight,
        line_transparency=line_transparency,
    )
    name_line = _set_shape_name_line("newShape", resolved_shape_name)
    z_line = f"z order targetShape z order position {_z_order_enum(z_order)}" if z_order else ""
    safe_resolved_name = _escape_applescript_string(resolved_shape_name)
    style_block = "\n    ".join(line for line in [*style_lines, z_line] if line)

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set newShape to make new shape at sl with properties {{left position:{float(left)}, top:{float(top)}, width:{float(width)}, height:{float(height)}, auto shape type:{shape_enum}}}
    {name_line}
    set targetShape to missing value
    repeat with i from 1 to (count of shapes of sl)
        set candidateShape to shape i of sl
        try
            if (name of candidateShape) is "{safe_resolved_name}" then
                set targetShape to candidateShape
                exit repeat
            end if
        end try
    end repeat
    if targetShape is missing value then
        error "Could not refetch shape named '{safe_resolved_name}' after creation"
    end if
    {style_block}
    return "ok"
end tell
'''
    _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {
        "shape_name": resolved_shape_name,
        "left": float(left),
        "top": float(top),
        "width": float(width),
        "height": float(height),
        "shape_type": shape_type,
    }


@mcp.tool()
def pptx_add_text_box(
    slide_index: int,
    text: str,
    left: float = 50,
    top: float = 50,
    width: float = 400,
    height: float = 80,
    font_size: float = 24,
    font_name: str | None = None,
    font_color: str = "#000000",
    bold: bool = False,
    italic: bool = False,
    align: str = "left",
    vertical_anchor: str = "top",
    fill_color: str | None = "none",
    fill_transparency: float = 0,
    line_color: str | None = "none",
    line_weight: float = 0,
    margin_left: float = 0,
    margin_right: float = 0,
    margin_top: float = 0,
    margin_bottom: float = 0,
    shape_name: str | None = None,
    z_order: str | None = None,
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Add a text-bearing shape that behaves like a PowerPoint text box.

    PowerPoint for Mac's AppleScript dictionary exposes native text boxes as a shape
    subtype. This creates a real text box, hides fill/line by default, and writes the
    text into its text frame with the requested font styling.
    """
    safe_text = _escape_applescript_string(text)
    safe_font = _escape_applescript_string(font_name) if font_name else None
    font_literal = _required_color_literal(font_color, label="font_color")
    align_enum = _paragraph_alignment_enum(align)
    anchor_enum = _vertical_anchor_enum(vertical_anchor)
    resolved_shape_name = shape_name or _generated_shape_name("pptx_mcp_text")
    style_lines = _style_shape_script_lines(
        "targetShape",
        fill_color=fill_color,
        fill_transparency=fill_transparency,
        line_color=line_color,
        line_weight=line_weight,
        line_transparency=0,
    )
    name_line = _set_shape_name_line("newShape", resolved_shape_name)
    font_name_line = f'set font name of font of tr to "{safe_font}"' if safe_font else ""
    z_line = f"z order targetShape z order position {_z_order_enum(z_order)}" if z_order else ""
    safe_resolved_name = _escape_applescript_string(resolved_shape_name)
    style_block = "\n    ".join(style_lines)
    font_lines = "\n        ".join(
        line for line in [
            f'set content of tr to "{safe_text}"',
            f"set font size of font of tr to {float(font_size)}",
            font_name_line,
            f"set font color of font of tr to {font_literal}",
            f"set bold of font of tr to {str(bool(bold)).lower()}",
            f"set italic of font of tr to {str(bool(italic)).lower()}",
            f"set alignment of paragraph format of tr to {align_enum}",
            f"set vertical anchor of text frame of targetShape to {anchor_enum}",
            f"set margin left of text frame of targetShape to {float(margin_left)}",
            f"set margin right of text frame of targetShape to {float(margin_right)}",
            f"set margin top of text frame of targetShape to {float(margin_top)}",
            f"set margin bottom of text frame of targetShape to {float(margin_bottom)}",
            "set word wrap of text frame of targetShape to true",
            z_line,
        ] if line
    )

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set newShape to make new text box at sl with properties {{left position:{float(left)}, top:{float(top)}, width:{float(width)}, height:{float(height)}}}
    {name_line}
    set targetShape to missing value
    repeat with i from 1 to (count of shapes of sl)
        set candidateShape to shape i of sl
        try
            if (name of candidateShape) is "{safe_resolved_name}" then
                set targetShape to candidateShape
                exit repeat
            end if
        end try
    end repeat
    if targetShape is missing value then
        error "Could not refetch text box named '{safe_resolved_name}' after creation"
    end if
    {style_block}
    set tr to text range of text frame of targetShape
        {font_lines}
    return "ok"
end tell
'''
    _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {
        "shape_name": resolved_shape_name,
        "text": text,
        "left": float(left),
        "top": float(top),
        "width": float(width),
        "height": float(height),
    }


@mcp.tool()
def pptx_set_shape_fill(
    slide_index: int,
    shape_name: str,
    fill_color: str | None = "#FFFFFF",
    transparency: float = 0,
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Set or clear a shape's fill color."""
    safe_name = _escape_applescript_string(shape_name)
    fill_literal = _as_color_literal(fill_color, label="fill_color")
    if fill_literal is None:
        fill_script = "set visible of fill format of shp to false"
        display_color = None
    else:
        fill_script = f'''
set visible of fill format of shp to true
solid fill format of shp
set fore color of fill format of shp to {fill_literal}
set transparency of fill format of shp to {_normalize_transparency(transparency)}
'''
        display_color = fill_color

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                {fill_script}
                return name of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    out = _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {"shape_name": out, "fill_color": display_color, "transparency": _normalize_transparency(transparency)}


@mcp.tool()
def pptx_set_shape_line(
    slide_index: int,
    shape_name: str,
    line_color: str | None = "#000000",
    weight: float = 1,
    transparency: float = 0,
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Set or visually clear a shape's border/line."""
    safe_name = _escape_applescript_string(shape_name)
    line_literal = _as_color_literal(line_color, label="line_color")
    if line_literal is None or float(weight) <= 0:
        line_script = '''
set line weight of line format of shp to 0
set transparency of line format of shp to 1.0
'''
        display_color = None
        display_weight = 0.0
    else:
        line_script = f'''
set fore color of line format of shp to {line_literal}
set line weight of line format of shp to {float(weight)}
set transparency of line format of shp to {_normalize_transparency(transparency)}
'''
        display_color = line_color
        display_weight = float(weight)

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                {line_script}
                return name of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    out = _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {
        "shape_name": out,
        "line_color": display_color,
        "weight": display_weight,
        "transparency": _normalize_transparency(transparency),
    }


@mcp.tool()
def pptx_set_text_style(
    slide_index: int,
    shape_name: str,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    align: str | None = None,
    vertical_anchor: str | None = None,
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Update text styling for an existing shape with a text frame."""
    safe_name = _escape_applescript_string(shape_name)
    lines: list[str] = []
    if font_size is not None:
        lines.append(f"set font size of font of tr to {float(font_size)}")
    if font_name is not None:
        lines.append(f'set font name of font of tr to "{_escape_applescript_string(font_name)}"')
    if font_color is not None:
        lines.append(f"set font color of font of tr to {_required_color_literal(font_color, label='font_color')}")
    if bold is not None:
        lines.append(f"set bold of font of tr to {str(bool(bold)).lower()}")
    if italic is not None:
        lines.append(f"set italic of font of tr to {str(bool(italic)).lower()}")
    if align is not None:
        lines.append(f"set alignment of paragraph format of tr to {_paragraph_alignment_enum(align)}")
    if vertical_anchor is not None:
        lines.append(f"set vertical anchor of text frame of shp to {_vertical_anchor_enum(vertical_anchor)}")
    if not lines:
        raise ValueError("No text style changes requested.")
    mutations = "\n                ".join(lines)

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                set tr to text range of text frame of shp
                {mutations}
                return name of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    out = _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {"shape_name": out}


@mcp.tool()
def pptx_set_z_order(
    slide_index: int,
    shape_name: str,
    action: str = "front",
    delay_seconds: float = 0.0,
) -> dict[str, Any]:
    """Move a shape forward/backward in the slide stacking order."""
    safe_name = _escape_applescript_string(shape_name)
    action_enum = _z_order_enum(action)
    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                z order shp z order position {action_enum}
                return name of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    out = _run_osascript(script)
    _maybe_live_delay(delay_seconds)
    return {"shape_name": out, "action": action}


@mcp.tool()
def pptx_run_live_sequence(
    steps: list[dict[str, Any]],
    default_delay_seconds: float = 0.6,
) -> dict[str, Any]:
    """Run a JSON sequence of live PPT actions with pauses between visible steps.

    Supported step types: `create_presentation`, `add_slide`, `focus_slide`, `shape`,
    `text`, `image`, `set_fill`, `set_line`, `set_text`, `set_text_style`,
    `set_geometry`, `z_order`, `pause`, `save`.

    The model can produce a declarative scene plan, then PowerPoint visibly mutates
    one element at a time.
    """
    results: list[dict[str, Any]] = []
    for idx, step in enumerate(steps, 1):
        kind = str(step.get("type", "")).strip().lower().replace("-", "_")
        delay = float(step.get("delay_seconds", default_delay_seconds))
        payload = dict(step)
        payload.pop("type", None)
        payload.pop("delay_seconds", None)

        try:
            if kind == "create_presentation":
                result = pptx_create_presentation(**payload)
            elif kind == "add_slide":
                result = pptx_add_slide(**payload)
            elif kind == "focus_slide":
                result = pptx_focus_slide(**payload)
            elif kind == "shape":
                result = pptx_add_shape(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "text":
                result = pptx_add_text_box(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "image":
                result = pptx_insert_image(**payload)
            elif kind == "set_fill":
                result = pptx_set_shape_fill(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "set_line":
                result = pptx_set_shape_line(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "set_text":
                result = pptx_set_text_in_shape_by_name(**payload)
            elif kind == "set_text_style":
                result = pptx_set_text_style(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "set_geometry":
                result = pptx_set_shape_geometry(**payload)
            elif kind == "z_order":
                result = pptx_set_z_order(delay_seconds=delay, **payload)
                delay = 0
            elif kind == "pause":
                result = pptx_live_pause(**payload)
                delay = 0
            elif kind == "save":
                result = pptx_save_presentation(**payload)
            else:
                raise ValueError(f"Unsupported live sequence step type: {kind!r}")
        except Exception as exc:
            label = payload.get("shape_name") or payload.get("text") or payload.get("layout") or ""
            label_text = f" ({label})" if label else ""
            raise RuntimeError(f"Live sequence step {idx} `{kind}`{label_text} failed: {exc}") from exc

        results.append({"step": idx, "type": kind, "result": result})
        _maybe_live_delay(delay)
    return {"steps_run": len(results), "results": results}


@mcp.tool()
def pptx_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Read all text from a slide's shapes.

    Use this instead of upstream `get_slide_content` — that one is broken (#20473).

    Returns:
        dict with `text` (newline-joined) and `shapes` (per-shape entries).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set acc to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (has text frame of shp) then
                set shpName to (name of shp) as text
                set shpText to ""
                try
                    set shpText to (content of text range of text frame of shp) as text
                end try
                set lineStr to shpName & "<<F>>" & shpText
                if acc is "" then
                    set acc to lineStr
                else
                    set acc to acc & "<<NL>>" & lineStr
                end if
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    shapes: list[dict[str, str]] = []
    text_lines: list[str] = []
    if out:
        for line in out.split("<<NL>>"):
            if "<<F>>" in line:
                name, _, txt = line.partition("<<F>>")
                shapes.append({"name": name, "text": txt})
                text_lines.append(txt)
            elif line:
                shapes.append({"name": "", "text": line})
                text_lines.append(line)
    return {"text": "\n".join(text_lines), "shapes": shapes}


# --- Tools: new — visual feedback, addressing, layout ops -----------------

@mcp.tool()
def pptx_get_slide_thumbnail(
    slide_index: int,
    dpi: int = 100,
    save_to_path: str | None = None,
) -> Image:
    """Render a single slide as a PNG and return it inline for the assistant to see.

    Implementation: PowerPoint exports the active presentation to PDF inside its
    sandboxed temp directory (`~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/pptx-live-mcp/`),
    then `pdftoppm` extracts the requested page as PNG. Writing outside the PowerPoint
    sandbox triggers TCC prompts and silently fails, so we always stay inside.

    Args:
        slide_index: 1-based index of the slide to render.
        dpi: Render resolution. 100 is a good default for inline previews.
        save_to_path: If provided, also persist the PNG to this filesystem path
            (any directory; parent dirs are created as needed; `~` is expanded).
            Useful for keeping a thumbnail artifact alongside commits or docs.

    Returns:
        Image (PNG) wrapped as an MCP ImageContent block — visible inline to the model.
    """
    pdftoppm_path = _pdftoppm_binary()

    # PowerPoint excludes hidden slides from PDF export regardless of the
    # `print hidden slides` flag. Map the requested PowerPoint slide_index to the
    # corresponding PDF page index (= position among visible slides).
    total, hidden, visible = _get_slide_counts_and_hidden()
    if int(slide_index) in hidden:
        raise RuntimeError(
            f"Slide {slide_index} is hidden — PowerPoint excludes hidden slides from PDF "
            f"export, so there's no PNG to render. Unhide the slide in PowerPoint "
            f"(slide show transition → hidden = false) or pick another."
        )
    if int(slide_index) < 1 or int(slide_index) > total:
        raise RuntimeError(f"slide_index {slide_index} out of range [1, {total}]")
    try:
        pdf_page = visible.index(int(slide_index)) + 1
    except ValueError:
        raise RuntimeError(f"slide_index {slide_index} not in visible list (unexpected)")

    tmp_dir = _sandbox_tmp_dir("thumb_")
    pdf_path = os.path.join(tmp_dir, "deck.pdf")

    safe_pdf = _escape_applescript_string(pdf_path)
    script = f'''
tell application "Microsoft PowerPoint"
    save active presentation in (POSIX file "{safe_pdf}") as save as PDF
end tell
'''
    _run_osascript(script, timeout=180)

    if not os.path.exists(pdf_path):
        raise RuntimeError(
            f"PDF export produced no file at {pdf_path}. Check that PowerPoint has an "
            f"active presentation open and is responsive."
        )

    prefix = os.path.join(tmp_dir, "slide")
    result = subprocess.run(
        [pdftoppm_path,
         "-f", str(pdf_page), "-l", str(pdf_page),
         "-png", "-r", str(int(dpi)),
         pdf_path, prefix],
        capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr.strip()}")

    # pdftoppm zero-pads the page number based on the total page count.
    candidates = (
        glob.glob(f"{prefix}-{pdf_page}.png")
        + glob.glob(f"{prefix}-{pdf_page:02d}.png")
        + glob.glob(f"{prefix}-{pdf_page:03d}.png")
    )
    if not candidates:
        raise RuntimeError(
            f"pdftoppm produced no PNG for PDF page {pdf_page} (slide {slide_index}). "
            f"tmp_dir: {os.listdir(tmp_dir)}"
        )

    # Open the raw PNG, wrap in a labeled cell (label above + light border).
    # Label shows the PowerPoint slide_index, not the PDF page index.
    raw = PILImage.open(candidates[0]).convert("RGB")
    label_h = max(28, int(dpi) // 3)
    font_size = max(18, int(dpi) // 5)
    composed = _compose_labeled_cell(raw, int(slide_index), label_height=label_h, font_size=font_size)
    buf = io.BytesIO()
    composed.save(buf, format="PNG", optimize=True)
    png_bytes = buf.getvalue()
    _save_png_to_path(png_bytes, save_to_path)
    return Image(data=png_bytes, format="png")


@mcp.tool()
def pptx_list_shapes(slide_index: int) -> dict[str, Any]:
    """Inventory every shape on a slide — name, geometry, current text, and placeholder
    info if the shape has one. Diagnostic step before `pptx_set_text_in_shape_by_name`.

    Why this exists (and not `list_placeholders`): in heavy corporate templates, slides
    are commonly built from freeform shapes rather than layout placeholders. PowerPoint
    AppleScript reports `count of placeholders of slide` = 0 for these. So we iterate
    `shapes` (which always exist) and surface placeholder info only when present.

    Args:
        slide_index: 1-based index of the slide.

    Returns:
        dict with `shapes`: list of {name, left, top, width, height, text,
        placeholder_idx (or null), placeholder_type (or null)}.
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set acc to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            set shpName to (name of shp) as text
            set L to ""
            try
                set L to (left position of shp) as text
            end try
            set T to ""
            try
                set T to (top of shp) as text
            end try
            set W to ""
            try
                set W to (width of shp) as text
            end try
            set H to ""
            try
                set H to (height of shp) as text
            end try
            set shpText to ""
            try
                if (has text frame of shp) then
                    set shpText to (content of text range of text frame of shp) as text
                end if
            end try
            set phIdx to ""
            set phType to ""
            try
                set phIdx to (|placeholder index| of placeholder format of shp) as text
                try
                    set phType to (|placeholder type| of placeholder format of shp) as text
                end try
            end try
            set rowStr to shpName & "<<F>>" & L & "<<F>>" & T & "<<F>>" & W & "<<F>>" & H & "<<F>>" & phIdx & "<<F>>" & phType & "<<F>>" & shpText
            if acc is "" then
                set acc to rowStr
            else
                set acc to acc & "<<NL>>" & rowStr
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    shapes: list[dict[str, Any]] = []
    if out:
        for line in out.split("<<NL>>"):
            parts = line.split("<<F>>", 7)
            if len(parts) != 8:
                continue
            name, L, T, W, H, phIdx, phType, text = parts
            entry: dict[str, Any] = {
                "name": name,
                "text": text,
                "placeholder_idx": int(phIdx) if phIdx.isdigit() else None,
                "placeholder_type": phType or None,
            }
            for key, raw in (("left", L), ("top", T), ("width", W), ("height", H)):
                try:
                    entry[key] = float(raw)
                except ValueError:
                    entry[key] = raw or None
            shapes.append(entry)
    return {"shapes": shapes}


@mcp.tool()
def pptx_set_text_in_shape_by_name(
    slide_index: int, shape_name: str, text: str
) -> dict[str, Any]:
    """Write text into a shape addressed by its `name` (the stable identifier in
    PowerPoint's AppleScript dictionary).

    Why by name and not by `placeholder_format.idx`: in this template, AppleScript
    reports zero placeholders on every slide — shapes are freeform. The `name` property
    ("Text 0", "Title 1", "Группа 14", etc.) is the only stable identifier available
    through the AppleScript bridge. Use `pptx_list_shapes` first to find the right
    name.

    Caveat: `set content of text range` replaces the text but preserves the formatting
    of the first run only (PowerPoint extends the first run's rPr over the new content).
    Multi-run formatting is collapsed.

    Args:
        slide_index: 1-based index of the target slide.
        shape_name: Exact `name` of the shape (case-sensitive).
        text: New text content.

    Returns:
        dict with `shape_name` of the updated shape.
    """
    safe_name = _escape_applescript_string(shape_name)
    safe_text = _escape_applescript_string(text)
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set matched to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                set content of text range of text frame of shp to "{safe_text}"
                set matched to (name of shp) as text
                exit repeat
            end if
        end try
    end repeat
    if matched is "" then
        error "No shape named '{safe_name}' on slide {int(slide_index)}"
    end if
    return matched
end tell
'''
    out = _run_osascript(script)
    return {"shape_name": out}


@mcp.tool()
def pptx_move_slide(from_index: int, to_index: int) -> dict[str, Any]:
    """Reorder slides natively via PowerPoint AppleScript.

    Args:
        from_index: 1-based current position of the slide to move.
        to_index: 1-based target position.

    Returns:
        dict with `new_index` (the slide's index after the move).
    """
    if int(to_index) <= int(from_index):
        loc = f"to before slide {int(to_index)} of activePres"
    else:
        loc = f"to after slide {int(to_index)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set srcSlide to slide {int(from_index)} of activePres
    move srcSlide {loc}
    return slide index of srcSlide
end tell
'''
    out = _run_osascript(script)
    return {"new_index": int(out) if out.isdigit() else out}


@mcp.tool()
def pptx_add_blank_slide_from_template(
    source_slide_index: int, position: int | None = None
) -> dict[str, Any]:
    """Append a blank new slide that inherits the source slide's custom layout, **without**
    any of the source's freeform shapes/decorations.

    Use when you want a clean layout shell (just the placeholder slots defined in the
    layout XML), not a full visual clone. Contrast with `pptx_add_slide_from_template`,
    which copies everything including the author's hand-placed shapes.

    Mechanism: `make new slide at end of p` (yields a slide with 0 shapes), then
    `set custom layout of newSlide to (custom layout of slide M of p)`. PowerPoint
    materializes the layout's placeholder shapes on the new slide.

    Args:
        source_slide_index: 1-based index of the slide whose custom layout to inherit.
        position: 1-based target index. If omitted, the new slide stays at the end.

    Returns:
        dict with `slide_index` and the resulting `shape_count` (number of placeholders
        materialized from the layout XML).
    """
    move_clause = ""
    if position is not None:
        if int(position) <= 1:
            move_clause = "move newSlide to before slide 1 of p"
        else:
            move_clause = f"move newSlide to before slide {int(position)} of p"

    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set newSlide to make new slide at end of p
    set sourceLayout to custom layout of slide {int(source_slide_index)} of p
    set custom layout of newSlide to sourceLayout
    {move_clause}
    return (slide index of newSlide as text) & "|" & (count of shapes of newSlide)
end tell
'''
    out = _run_osascript(script)
    parts = out.split("|")
    if len(parts) == 2:
        idx, count = parts
        return {
            "slide_index": int(idx) if idx.isdigit() else idx,
            "shape_count": int(count) if count.isdigit() else count,
        }
    return {"raw": out}


@mcp.tool()
def pptx_delete_shape_by_name(slide_index: int, shape_name: str) -> dict[str, Any]:
    """Delete the first shape on a slide whose `name` matches.

    Pairs with `pptx_set_slide_layout_from_template` — that tool adds the new layout's
    placeholders on top of existing shapes (additive), so callers usually need to delete
    stale shapes from the previous layout afterwards. Also useful for trimming unwanted
    placeholders from cloned slides.

    Args:
        slide_index: 1-based index of the slide.
        shape_name: Exact `name` of the shape to delete (case-sensitive).

    Returns:
        dict with `deleted` (True if a shape was deleted), `shapes_remaining`,
        and `shape_name` of the deleted shape.
    """
    safe_name = _escape_applescript_string(shape_name)
    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set deletedName to ""
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                set deletedName to (name of shp) as text
                delete shp
                exit repeat
            end if
        end try
    end repeat
    return deletedName & "|" & (count of shapes of sl)
end tell
'''
    out = _run_osascript(script)
    name, _, count = out.partition("|")
    return {
        "deleted": bool(name),
        "shape_name": name,
        "shapes_remaining": int(count) if count.isdigit() else count,
    }


@mcp.tool()
def pptx_set_shape_geometry(
    slide_index: int,
    shape_name: str,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Reposition or resize a shape on a slide.

    Any argument left as `None` keeps the current value untouched — useful when you only
    want to shift X or only resize width.

    Args:
        slide_index: 1-based index of the slide.
        shape_name: Exact `name` of the target shape.
        left, top: New position in points. None leaves the value unchanged.
        width, height: New size in points. None leaves the value unchanged.

    Returns:
        dict with the shape's new geometry (left, top, width, height).
    """
    safe_name = _escape_applescript_string(shape_name)
    lines = []
    if left is not None:
        lines.append(f"set left position of shp to {float(left)}")
    if top is not None:
        lines.append(f"set top of shp to {float(top)}")
    if width is not None:
        lines.append(f"set width of shp to {float(width)}")
    if height is not None:
        lines.append(f"set height of shp to {float(height)}")
    if not lines:
        raise ValueError("No geometry change requested: pass at least one of left, top, width, height.")
    mutations = "\n            ".join(lines)

    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set found to false
    repeat with i from 1 to (count of shapes of sl)
        set shp to shape i of sl
        try
            if (name of shp) is "{safe_name}" then
                {mutations}
                set found to true
                set L to (left position of shp) as text
                set T to (top of shp) as text
                set W to (width of shp) as text
                set H to (height of shp) as text
                return L & "|" & T & "|" & W & "|" & H
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    out = _run_osascript(script)
    parts = out.split("|")
    if len(parts) == 4:
        try:
            L, T, W, H = (float(p) for p in parts)
            return {"left": L, "top": T, "width": W, "height": H}
        except ValueError:
            pass
    return {"raw": out}


@mcp.tool()
def pptx_set_slide_layout_from_template(
    slide_index: int, source_slide_index: int
) -> dict[str, Any]:
    """Apply the *custom layout* of an existing slide to another slide.

    For corporate templates that ship dozens of custom slide layouts (each with its own
    placeholder geometry, fonts, decorative master shapes), this is the only way to
    address a specific layout via AppleScript — built-in enums (`slide layout blank`,
    etc.) only cover ~30 standard layouts, not custom ones.

    Mechanism: `set custom layout of slide N to (custom layout of slide M)`. The target
    slide gains M's layout's placeholder shapes; existing shapes are NOT removed (this
    is layered, not replacement). If you want a clean visual match, delete unwanted
    shapes first.

    Args:
        slide_index: 1-based index of the slide whose layout to change.
        source_slide_index: 1-based index of a slide whose custom layout to copy.

    Returns:
        dict with `slide_index` and `shapes_added` (heuristic: shape-count delta).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set targetSlide to slide {int(slide_index)} of p
    set beforeCount to count of shapes of targetSlide
    set sourceLayout to custom layout of slide {int(source_slide_index)} of p
    set custom layout of targetSlide to sourceLayout
    set afterCount to count of shapes of targetSlide
    return (slide index of targetSlide as text) & "|" & beforeCount & "|" & afterCount
end tell
'''
    out = _run_osascript(script)
    parts = out.split("|")
    if len(parts) == 3:
        idx, before, after = parts
        return {
            "slide_index": int(idx) if idx.isdigit() else idx,
            "shapes_before": int(before),
            "shapes_after": int(after),
            "shapes_added": int(after) - int(before),
        }
    return {"raw": out}


@mcp.tool()
def pptx_replace_text_in_shape_by_name(
    slide_index: int, shape_name: str, old: str, new: str
) -> dict[str, Any]:
    """Replace a substring in one shape's text. Caveat: collapses styled runs.

    PowerPoint for Mac's AppleScript dictionary does NOT implement a native find/replace
    on text ranges — every syntax variant returns -2741 syntax error. We read the
    current text, do `str.replace` in Python, and write it back via
    `set content of text range`. This collapses every styled run on the text frame into
    one run with the formatting of the first run.

    If preserving styled runs matters (bold/colored phrases mid-text), use python-pptx
    paragraph/run-level edits instead.

    Args:
        slide_index: 1-based index of the slide.
        shape_name: Exact `name` of the shape.
        old: Substring to search for (exact match, case-sensitive).
        new: Replacement string.

    Returns:
        dict with `found` (True if a replacement happened), `before` text, `after` text.
    """
    safe_name = _escape_applescript_string(shape_name)
    # Step 1: read current text.
    read_script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                return content of text range of text frame of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    before_text = _run_osascript(read_script)
    if old not in before_text:
        return {"found": False, "before": before_text, "after": before_text}

    after_text = before_text.replace(old, new)
    safe_after = _escape_applescript_string(after_text)
    write_script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                set content of text range of text frame of shp to "{safe_after}"
                exit repeat
            end if
        end try
    end repeat
end tell
'''
    _run_osascript(write_script)
    return {"found": True, "before": before_text, "after": after_text}


@mcp.tool()
def pptx_get_deck_overview(
    start_slide: int = 1,
    per_page: int = 12,
    columns: int = 3,
    dpi: int = 60,
    save_to_path: str | None = None,
) -> Image:
    """Render multiple slides as a single grid image — quick overview of the deck.

    Mechanism: export the active presentation to PDF (sandboxed temp dir), use
    `pdftoppm` to extract pages `start_slide..start_slide+per_page-1` as PNGs, then
    compose them in a `columns`-wide grid via Pillow with a slide-number label
    above each thumbnail.

    For a 62-slide deck at default per_page=12, columns=4: 6 pages of 4×3 grids.
    The caller paginates by re-invoking with `start_slide=13, 25, ...`.

    Args:
        start_slide: 1-based first slide to include.
        per_page: How many slides per overview call.
        columns: Grid width. Rows derived from `ceil(per_page / columns)`.
        dpi: Per-slide render resolution. 60 = small thumbs; 100 = readable text.
        save_to_path: If provided, also persist the composite PNG to this filesystem
            path (any directory; parent dirs are created as needed; `~` is expanded).

    Returns:
        Image (PNG) wrapped as ImageContent — visible inline.
    """
    # Query slide count + hidden set. PowerPoint excludes hidden slides from the PDF
    # export, so we always work in "visible-slide" space and label thumbnails with
    # the real PowerPoint slide_index (not the PDF page index).
    total, hidden, visible = _get_slide_counts_and_hidden()
    if total == 0:
        raise RuntimeError("No active presentation, or it has 0 slides.")

    start = max(1, int(start_slide))
    if start > total:
        raise RuntimeError(
            f"start_slide={start} exceeds total slides ({total}). "
            f"Use start_slide in [1, {total}]."
        )
    end = min(total, start + int(per_page) - 1)

    # Build the slide-number list for this page, skipping hidden slides.
    requested_slides = [n for n in range(start, end + 1) if n not in hidden]
    if not requested_slides:
        raise RuntimeError(
            f"All slides in range [{start}, {end}] are hidden — no PDF pages to render. "
            f"Hidden: {sorted(hidden)}"
        )

    # Map each visible PowerPoint slide_index to its PDF page index.
    visible_pos = {ppt_idx: pdf_idx + 1 for pdf_idx, ppt_idx in enumerate(visible)}
    pdf_pages_needed = [visible_pos[n] for n in requested_slides]
    pdf_min, pdf_max = min(pdf_pages_needed), max(pdf_pages_needed)

    pdftoppm_path = _pdftoppm_binary()

    tmp_dir = _sandbox_tmp_dir("overview_")
    pdf_path = os.path.join(tmp_dir, "deck.pdf")
    safe_pdf = _escape_applescript_string(pdf_path)
    _run_osascript(
        f'''
tell application "Microsoft PowerPoint"
    save active presentation in (POSIX file "{safe_pdf}") as save as PDF
end tell
''',
        timeout=240,
    )
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF export produced no file at {pdf_path}.")

    prefix = os.path.join(tmp_dir, "slide")
    r = subprocess.run(
        [pdftoppm_path,
         "-f", str(pdf_min), "-l", str(pdf_max),
         "-png", "-r", str(int(dpi)),
         pdf_path, prefix],
        capture_output=True, text=True, check=False,
    )
    if r.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {r.stderr.strip()}")

    # Resolve each PowerPoint slide_index → its PDF page PNG.
    thumbs: list[tuple[int, "PILImage.Image"]] = []
    for ppt_idx in requested_slides:
        pdf_page = visible_pos[ppt_idx]
        candidates = []
        for w in (1, 2, 3, 4):
            candidates.extend(glob.glob(f"{prefix}-{pdf_page:0{w}d}.png"))
        if not candidates:
            continue
        thumbs.append((ppt_idx, PILImage.open(candidates[0]).convert("RGB")))

    if not thumbs:
        raise RuntimeError("pdftoppm produced no PNGs for the requested range.")

    # Uniform thumbnail size = max width/height across the batch (slides may vary
    # slightly during PDF export).
    thumb_w = max(im.width for _, im in thumbs)
    thumb_h = max(im.height for _, im in thumbs)

    cols = max(1, int(columns))
    rows = math.ceil(len(thumbs) / cols)

    # Layout: clean grid on white background. Each cell = label above + thumbnail
    # with a thin light border. Spacing creates the visual separation.
    label_h = max(28, thumb_w // 22)
    label_font_size = max(16, thumb_w // 28)
    gap_x = 28
    gap_y = 28
    canvas_margin = 32
    header_h = 36

    cell_w = thumb_w
    cell_h = label_h + thumb_h
    canvas_w = canvas_margin * 2 + cols * cell_w + (cols - 1) * gap_x
    canvas_h = canvas_margin + header_h + rows * cell_h + (rows - 1) * gap_y + canvas_margin

    canvas = PILImage.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    draw = ImageDraw.Draw(canvas)
    font_header = _load_label_font(20)

    hidden_in_range = sorted(h for h in hidden if start <= h <= end)
    hidden_note = f"  ·  hidden skipped: {hidden_in_range}" if hidden_in_range else ""
    header_text = (
        f"Slides {start}–{end} of {total}"
        f"  ·  page {((start - 1) // per_page) + 1} of {math.ceil(total / per_page)}"
        f"{hidden_note}"
    )
    draw.text((canvas_margin, canvas_margin // 2), header_text, fill=(40, 40, 40), font=font_header)

    for i, (slide_no, im) in enumerate(thumbs):
        col = i % cols
        row = i // cols
        x = canvas_margin + col * (cell_w + gap_x)
        y = canvas_margin + header_h + row * (cell_h + gap_y)
        cell = _compose_labeled_cell(
            im, slide_no,
            label_height=label_h, font_size=label_font_size,
        )
        canvas.paste(cell, (x, y))

    buf = io.BytesIO()
    canvas.save(buf, format="PNG", optimize=True)
    png_bytes = buf.getvalue()
    _save_png_to_path(png_bytes, save_to_path)
    return Image(data=png_bytes, format="png")


@mcp.tool()
def pptx_delete_slide(slide_index: int) -> dict[str, Any]:
    """Delete a slide from the active presentation by 1-based index.

    Delete a slide from the active presentation.

    Args:
        slide_index: 1-based index of the slide to delete.

    Returns:
        dict with `deleted_index` and `slides_remaining`.
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    delete slide {int(slide_index)} of p
    return (count of slides of p) as text
end tell
'''
    out = _run_osascript(script)
    try:
        return {"deleted_index": int(slide_index), "slides_remaining": int(out)}
    except ValueError:
        return {"deleted_index": int(slide_index), "raw": out}


@mcp.tool()
def pptx_copy_slide_from_pptx(
    source_pptx_path: str,
    source_slide_index: int,
    target_pptx_path: str | None = None,
    target_position: int | None = None,
    close_source_after: bool = True,
) -> dict[str, Any]:
    """Copy a slide from a DIFFERENT pptx into the target deck — full visual clone.

    Cross-presentation variant of `pptx_add_slide_from_template`. The in-deck clone
    does `tell active presentation / copy / paste / end tell` — bound to one deck. Here
    we address both decks by *filename* (so neither needs to be the active one) via
    `tell presentation <name>`. Mechanism stays the same: `copy object slide N` from
    source, `paste object` into target.

    Both files are opened in PowerPoint if not already open (and source can optionally
    be closed afterwards).

    Caveats:
      * Theme inheritance — PowerPoint usually applies the destination deck's theme to
        the pasted slide. AppleScript doesn't expose the "Keep Source Formatting" toggle.
      * Source file opens in the foreground (PowerPoint window flashes). If it's
        already open, that instance is reused — no re-open.

    Args:
        source_pptx_path: Absolute POSIX path to the source pptx.
        source_slide_index: 1-based index of the slide in source to copy.
        target_pptx_path: Absolute POSIX path to the target pptx. If None, falls back
            to whatever is the active presentation (must NOT be the same file as source).
        target_position: 1-based position where the pasted slide should land in target.
            If None, the slide stays at the end of the target (default paste behavior).
        close_source_after: If True (default), close source after copying. Set False
            if you plan to copy more slides from the same source — saves a reopen.

    Returns:
        dict with `new_slide_index`, `target_name`, `source_name`.
    """
    if not os.path.exists(source_pptx_path):
        raise RuntimeError(f"source pptx not found: {source_pptx_path}")

    source_filename = os.path.basename(source_pptx_path)
    safe_source_path = _escape_applescript_string(source_pptx_path)
    safe_source_name = _escape_applescript_string(source_filename)

    target_filename = None
    safe_target_path = None
    safe_target_name = None
    if target_pptx_path:
        if not os.path.exists(target_pptx_path):
            raise RuntimeError(f"target pptx not found: {target_pptx_path}")
        target_filename = os.path.basename(target_pptx_path)
        safe_target_path = _escape_applescript_string(target_pptx_path)
        safe_target_name = _escape_applescript_string(target_filename)

    if target_position is None:
        move_clause = ""
    else:
        move_clause = f'''
    set newIdx to count of slides of presentation targetName
    if {int(target_position)} ≤ newIdx then
        move slide newIdx of presentation targetName to before slide {int(target_position)} of presentation targetName
    end if
'''

    if target_pptx_path:
        # Open target if needed; address by name.
        target_resolve = f'''
    set targetName to "{safe_target_name}"
    set targetOpen to false
    repeat with i from 1 to count of presentations
        if (name of presentation i) is targetName then
            set targetOpen to true
            exit repeat
        end if
    end repeat
    if not targetOpen then
        open POSIX file "{safe_target_path}"
        delay 1
    end if
'''
    else:
        # Use the currently-active presentation as target. Captured *before* opening
        # the source (the open shifts active to source).
        target_resolve = '''
    set targetName to (name of active presentation) as text
'''

    close_clause = ""
    if close_source_after:
        close_clause = 'close presentation "' + safe_source_name + '" saving no'

    script = f'''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
{target_resolve}
    if targetName is "{safe_source_name}" then
        error "target and source are the same file — use pptx_add_slide_from_template for in-deck clones"
    end if
    -- Open source if not already open.
    set sourceOpen to false
    repeat with i from 1 to count of presentations
        if (name of presentation i) is "{safe_source_name}" then
            set sourceOpen to true
            exit repeat
        end if
    end repeat
    if not sourceOpen then
        open POSIX file "{safe_source_path}"
        delay 1
    end if
    tell presentation "{safe_source_name}"
        copy object slide {int(source_slide_index)}
    end tell
    tell presentation targetName
        paste object
    end tell
{move_clause}
    set finalIdx to (count of slides of presentation targetName)
    if {0 if target_position is None else int(target_position)} > 0 then
        set finalIdx to {0 if target_position is None else int(target_position)}
    end if
    {close_clause}
    return (finalIdx as text) & "|" & targetName
end tell
'''
    out = _run_osascript(script, timeout=180)
    parts = out.split("|", 1)
    if len(parts) == 2:
        idx, target_name = parts
        return {
            "new_slide_index": int(idx) if idx.isdigit() else idx,
            "target_name": target_name,
            "source_name": source_filename,
        }
    return {"raw": out}


# --- Presentation lifecycle (self-sufficiency: create/open/save/close/export_pdf) ---

@mcp.tool()
def pptx_create_presentation(
    save_to_path: str | None = None,
) -> dict[str, Any]:
    """Create a new empty presentation. If `save_to_path` is given, save it there
    immediately; otherwise it stays unsaved in memory (default PowerPoint name like
    "Presentation N").

    Returns:
        dict with `name` of the new presentation and `slide_count` (typically 1 —
        PowerPoint adds a blank title slide by default).
    """
    if save_to_path:
        target = os.path.abspath(os.path.expanduser(save_to_path.strip()))
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)
        safe_path = _escape_applescript_string(target)
        script = f'''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
    set p to make new presentation
    save p in (POSIX file "{safe_path}")
    return (name of p as text) & "|" & (count of slides of p)
end tell
'''
    else:
        script = '''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
    set p to make new presentation
    return (name of p as text) & "|" & (count of slides of p)
end tell
'''
    out = _run_osascript(script, timeout=60)
    name, _, count = out.partition("|")
    return {"name": name, "slide_count": int(count) if count.isdigit() else count}


@mcp.tool()
def pptx_open_presentation(pptx_path: str) -> dict[str, Any]:
    """Open a pptx file in PowerPoint. If already open, brings it to front.

    Args:
        pptx_path: Absolute POSIX path to the pptx file.

    Returns:
        dict with `name`, `slide_count`, and `was_already_open`.
    """
    if not os.path.exists(pptx_path):
        raise RuntimeError(f"pptx not found: {pptx_path}")
    safe_path = _escape_applescript_string(pptx_path)
    filename = os.path.basename(pptx_path)
    safe_name = _escape_applescript_string(filename)
    script = f'''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
    set wasOpen to false
    repeat with i from 1 to count of presentations
        if (name of presentation i) is "{safe_name}" then
            set wasOpen to true
            exit repeat
        end if
    end repeat
    if not wasOpen then
        open POSIX file "{safe_path}"
        delay 1
    end if
    set p to presentation "{safe_name}"
    return (name of p as text) & "|" & (count of slides of p) & "|" & (wasOpen as text)
end tell
'''
    out = _run_osascript(script, timeout=60)
    parts = out.split("|")
    if len(parts) != 3:
        return {"raw": out}
    name, count, was_open = parts
    return {
        "name": name,
        "slide_count": int(count) if count.isdigit() else count,
        "was_already_open": was_open.strip().lower() == "true",
    }


@mcp.tool()
def pptx_save_presentation(
    save_as_path: str | None = None,
) -> dict[str, Any]:
    """Save the active presentation. If `save_as_path` is given, save-as to that path
    (changes the file PowerPoint considers the current document). Otherwise saves to
    the current path (errors if the deck has never been saved).

    Args:
        save_as_path: Optional absolute POSIX path. If provided, the active
            presentation is saved to that path; `~` is expanded; parent dirs created.

    Returns:
        dict with `name` and `path` of the saved file.
    """
    if save_as_path:
        target = os.path.abspath(os.path.expanduser(save_as_path.strip()))
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)
        if os.path.commonpath([os.path.abspath(POWERPOINT_SANDBOX_TMP), target]) == os.path.abspath(POWERPOINT_SANDBOX_TMP):
            powerpoint_target = target
            copy_after_save = False
        else:
            tmp_dir = _sandbox_tmp_dir("save-")
            filename = os.path.basename(target) or "presentation.pptx"
            powerpoint_target = os.path.join(tmp_dir, filename)
            copy_after_save = True
        safe_path = _escape_applescript_string(powerpoint_target)
        script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    save p in (POSIX file "{safe_path}")
    return (name of p as text) & "|" & (full name of p as text)
end tell
'''
    else:
        script = '''
tell application "Microsoft PowerPoint"
    set p to active presentation
    save p
    return (name of p as text) & "|" & (full name of p as text)
end tell
'''
    out = _run_osascript(script, timeout=120)
    name, _, full = out.partition("|")
    if save_as_path and copy_after_save:
        if not os.path.exists(powerpoint_target):
            raise RuntimeError(f"PowerPoint reported save success but file is missing: {powerpoint_target}")
        shutil.copy2(powerpoint_target, target)
        return {"name": name, "path": target, "powerpoint_internal_path": full}
    return {"name": name, "path": full}


@mcp.tool()
def pptx_close_presentation(
    save_changes: bool = False,
    presentation_name: str | None = None,
) -> dict[str, Any]:
    """Close a presentation.

    Args:
        save_changes: If True, save before closing. If False (default), discard
            unsaved changes.
        presentation_name: Filename to close (e.g. "deck.pptx"). If None, closes the
            currently-active presentation.

    Returns:
        dict with `closed_name` and `remaining_presentations`.
    """
    saving_clause = "saving yes" if save_changes else "saving no"
    if presentation_name:
        safe_name = _escape_applescript_string(presentation_name)
        target = f'presentation "{safe_name}"'
    else:
        target = "active presentation"
    script = f'''
tell application "Microsoft PowerPoint"
    set closedName to (name of {target}) as text
    close {target} {saving_clause}
    return closedName & "|" & (count of presentations)
end tell
'''
    out = _run_osascript(script, timeout=60)
    name, _, count = out.partition("|")
    return {
        "closed_name": name,
        "remaining_presentations": int(count) if count.isdigit() else count,
    }


@mcp.tool()
def pptx_export_pdf(
    pdf_path: str,
    presentation_name: str | None = None,
) -> dict[str, Any]:
    """Export a presentation to PDF.

    Args:
        pdf_path: Absolute POSIX path where the PDF should be written.
            `~` is expanded; parent dirs are created.
            **Note:** PowerPoint is sandboxed — writing to paths outside
            `~/Library/Containers/com.microsoft.Powerpoint/Data/` may trigger TCC
            prompts on first use. Approve once and subsequent calls run silently.
        presentation_name: Filename to export. If None, exports the currently-active
            presentation.

    Returns:
        dict with `path` of the written PDF and `bytes_written`.
    """
    target = os.path.abspath(os.path.expanduser(pdf_path.strip()))
    parent = os.path.dirname(target)
    if parent:
        os.makedirs(parent, exist_ok=True)
    safe_path = _escape_applescript_string(target)

    if presentation_name:
        safe_name = _escape_applescript_string(presentation_name)
        pres_ref = f'presentation "{safe_name}"'
    else:
        pres_ref = "active presentation"

    script = f'''
tell application "Microsoft PowerPoint"
    save {pres_ref} in (POSIX file "{safe_path}") as save as PDF
end tell
'''
    _run_osascript(script, timeout=240)
    if not os.path.exists(target):
        raise RuntimeError(
            f"PDF export reported success but no file appeared at {target}. "
            f"Check sandbox permissions and try a path inside the PowerPoint container."
        )
    return {"path": target, "bytes_written": os.path.getsize(target)}


# --- Entry point ----------------------------------------------------------

def main() -> None:
    """Run the MCP server over stdio."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    mcp.run()


if __name__ == "__main__":
    main()
