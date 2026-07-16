"""SVG -> picture conversion for embedding in PPTX.

python-pptx cannot embed SVG (add_picture relies on PIL, which can't parse it),
so icons must be converted first. Preference order:

  1. **EMF** via Inkscape (`inkscape in.svg --export-filename=out.emf`) — a true
     vector object that scales crisply and stays editable/ungroupable in PowerPoint.
  2. **EMF** via LibreOffice (`soffice --headless --convert-to emf`) — fallback.
  3. **PNG** via macOS QuickLook (`qlmanage -t`) — zero-install raster fallback that
     works out of the box on macOS with no Homebrew/native libs.
  4. **PNG** via LibreOffice Draw — cross-platform fallback when vector output is
     intentionally disabled for maximum PowerPoint preview compatibility.
  5. **High-DPI PNG** via cairosvg — raster fallback where libcairo is present.
     (NOTE: cairosvg has no EMF output, despite a popular blog claiming
     `cairosvg.svg2emf` — that function does not exist.)

Results are cached on disk by content hash + options, so each icon converts once.

Brand-icon licensing guardrail: only ``recolorable`` icons (CC0 simple-icons) are
ever recolored; brand icons (AWS/Azure/GCP) are converted verbatim, uniform-scale
only — never distorted or recolored.
"""

from __future__ import annotations

import hashlib
import re
import shutil
import subprocess
from pathlib import Path

_CACHE_DIR = Path(__file__).resolve().parent.parent.parent / ".archgen_cache" / "icons"

# Rasterization DPI for the PNG fallback — high enough to stay crisp when scaled
# into a ~1-inch slot on a 13.3-inch slide.
PNG_DPI = 300
RASTER_CACHE_VERSION = 5


class ConversionError(RuntimeError):
    pass


def to_picture(svg_path: Path, recolor_hex: str | None = None, prefer_vector: bool = True) -> Path:
    """Convert ``svg_path`` to a PPTX-embeddable picture; return the cached file path.

    ``recolor_hex`` (e.g. ``"#0078D4"``) is applied ONLY for recolorable icons;
    callers must not pass it for brand icons.
    """
    svg_bytes = _normalize_size(svg_path.read_bytes())
    if recolor_hex:
        svg_bytes = _apply_fill(svg_bytes, recolor_hex)

    errors: list[str] = []
    for engine in _candidate_engines(prefer_vector):
        key = hashlib.sha256(
            svg_bytes + f"|{engine}|{recolor_hex}|{PNG_DPI}|raster-v{RASTER_CACHE_VERSION}".encode()
        ).hexdigest()[:16]
        ext = "emf" if engine in ("inkscape", "soffice") else "png"
        out = _CACHE_DIR / f"{svg_path.stem}.{key}.{ext}"
        if out.is_file():
            return out
        out.parent.mkdir(parents=True, exist_ok=True)

        # Write the (possibly recolored) SVG to a temp file for CLI converters.
        tmp_svg = _CACHE_DIR / f"_src_{key}.svg"
        tmp_svg.write_bytes(svg_bytes)
        try:
            if engine == "inkscape":
                _inkscape(tmp_svg, out)
            elif engine == "soffice":
                _soffice(tmp_svg, out)
            elif engine == "soffice_png":
                _soffice_png(tmp_svg, out)
            elif engine == "qlmanage":
                _qlmanage(tmp_svg, out)
            else:
                _cairosvg(svg_bytes, out)
            if out.is_file():
                return out
            errors.append(f"{engine}: produced no output")
        except Exception as exc:  # noqa: BLE001 - a broken optional engine must fall through
            out.unlink(missing_ok=True)
            detail = getattr(exc, "stderr", None) or str(exc)
            errors.append(f"{engine}: {str(detail).strip()}")
        finally:
            tmp_svg.unlink(missing_ok=True)

    detail = "; ".join(errors)
    raise ConversionError(f"could not convert {svg_path}; {detail}")


def available_engine(prefer_vector: bool = True) -> str:
    """Report which conversion engine would be used (for diagnostics/CLI `doctor`)."""
    return _choose_engine(prefer_vector)


def _choose_engine(prefer_vector: bool) -> str:
    return _candidate_engines(prefer_vector)[0]


def _candidate_engines(prefer_vector: bool) -> list[str]:
    """Installed converters in preference order.

    Capability detection is optimistic: a binary can be present yet reject a
    particular SVG or fail in a headless environment. ``to_picture`` therefore
    tries every candidate rather than treating the first installed engine as final.
    """
    engines: list[str] = []
    if prefer_vector:
        if shutil.which("inkscape"):
            engines.append("inkscape")
        if shutil.which("soffice") or shutil.which("libreoffice"):
            engines.append("soffice")
    # LibreOffice's Draw PNG export respects the normalized 512px SVG dimensions
    # and works in headless/sandboxed environments where QuickLook may be blocked.
    if shutil.which("soffice") or shutil.which("libreoffice"):
        engines.append("soffice_png")
    # Remaining raster fallbacks: macOS QuickLook, then cairosvg (needs libcairo).
    if shutil.which("qlmanage"):
        engines.append("qlmanage")
    engines.append("cairosvg")
    return engines


def _inkscape(svg: Path, out: Path) -> None:
    # Inkscape 1.x CLI. (0.x used `-M`; we target 1.x, the current release.)
    subprocess.run(
        ["inkscape", str(svg), "--export-type=emf", f"--export-filename={out}"],
        check=True, capture_output=True, text=True,
    )


def _soffice(svg: Path, out: Path) -> None:
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    subprocess.run(
        [binary, "--headless", "--convert-to", "emf", "--outdir", str(out.parent), str(svg)],
        check=True, capture_output=True, text=True,
    )
    produced = out.parent / (svg.stem + ".emf")
    if produced != out and produced.is_file():
        produced.replace(out)


def _soffice_png(svg: Path, out: Path) -> None:
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    subprocess.run(
        [binary, "--headless", "--convert-to", "png", "--outdir", str(out.parent), str(svg)],
        check=True, capture_output=True, text=True,
    )
    produced = out.parent / (svg.stem + ".png")
    if produced != out and produced.is_file():
        produced.replace(out)


def _qlmanage(svg: Path, out: Path) -> None:
    # macOS QuickLook thumbnailer renders SVG to PNG with no extra installs.
    # It writes "<name>.png" into the output dir; we then move it into place.
    subprocess.run(
        ["qlmanage", "-t", "-s", "512", "-o", str(out.parent), str(svg)],
        check=True, capture_output=True, text=True,
    )
    produced = out.parent / (svg.name + ".png")
    if produced.is_file():
        produced.replace(out)
        _trim_quicklook_canvas(out)


def _trim_quicklook_canvas(path: Path) -> None:
    """Turn QuickLook's opaque white thumbnail into a transparent icon.

    QuickLook renders SVGs onto a solid white canvas, so icons arrive sitting on
    a white box. We (1) flood-fill the white background to transparent starting
    from the borders — so white *enclosed inside* a logo is preserved — then
    (2) crop to the visible artwork and recenter it, so PowerPoint receives a
    clean icon rather than a white card.
    """
    from PIL import Image, ImageChops, ImageDraw

    with Image.open(path) as source:
        rgba = source.convert("RGBA")
    w, h = rgba.size

    # (1) Flood the border-connected white to a sentinel color, then zero its alpha.
    rgb = rgba.convert("RGB")
    marker = (255, 0, 255)
    seeds = [
        (0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1),
        (w // 2, 0), (w // 2, h - 1), (0, h // 2), (w - 1, h // 2),
    ]
    flooded = False
    for sx, sy in seeds:
        r, g, b = rgb.getpixel((sx, sy))
        if r >= 245 and g >= 245 and b >= 245:
            ImageDraw.floodfill(rgb, (sx, sy), marker, thresh=32)
            flooded = True
    if flooded:
        keep = ImageChops.difference(rgb, Image.new("RGB", (w, h), marker)).convert("L")
        keep = keep.point(lambda v: 0 if v == 0 else 255)  # 0 where background
        r, g, b, a = rgba.split()
        rgba = Image.merge("RGBA", (r, g, b, ImageChops.multiply(a, keep)))

    # (2) Trim transparent margins and refit the artwork, centered.
    bbox = rgba.split()[3].getbbox()
    if bbox is None or bbox == (0, 0, w, h):
        rgba.save(path)
        return
    artwork = rgba.crop(bbox)
    scale = min(w / artwork.width, h / artwork.height)
    artwork = artwork.resize(
        (max(1, round(artwork.width * scale)), max(1, round(artwork.height * scale))),
        Image.Resampling.LANCZOS,
    )
    canvas = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    canvas.alpha_composite(artwork, ((w - artwork.width) // 2, (h - artwork.height) // 2))
    canvas.save(path)


def _cairosvg(svg_bytes: bytes, out: Path) -> None:
    import cairosvg  # imported lazily: needs native libcairo

    cairosvg.svg2png(bytestring=svg_bytes, write_to=str(out), dpi=PNG_DPI)


def svg_aspect(svg_path: Path) -> float:
    """Return width/height aspect ratio of an SVG (from viewBox or width/height).

    Used to fit icons into their slot *without distortion* (a licensing requirement
    for brand icons). Defaults to 1.0 (square) if it can't be determined.
    """
    head = svg_path.read_text(encoding="utf-8", errors="replace")[:600]
    m = re.search(r'viewBox\s*=\s*"[\s\d.\-]*?\s([\d.]+)\s+([\d.]+)\s*"', head)
    if m:
        w, h = float(m.group(1)), float(m.group(2))
        if w > 0 and h > 0:
            return w / h
    mw = re.search(r'\bwidth\s*=\s*"([\d.]+)', head)
    mh = re.search(r'\bheight\s*=\s*"([\d.]+)', head)
    if mw and mh and float(mh.group(1)) > 0:
        return float(mw.group(1)) / float(mh.group(1))
    return 1.0


def _normalize_size(svg_bytes: bytes) -> bytes:
    """Give viewBox-based SVGs an explicit square raster target.

    QuickLook does not consistently expand a dimensionless SVG to its thumbnail
    canvas. In particular, the official AWS 48px pack can become a tiny glyph in a
    large white preview. Replacing the intrinsic dimensions with a 512px target
    preserves the viewBox/aspect ratio while making the artwork fill the raster.
    """
    text = svg_bytes.decode("utf-8", errors="replace")
    open_tag_end = text.find(">")
    if open_tag_end == -1:
        return svg_bytes
    head, rest = text[:open_tag_end], text[open_tag_end:]
    if "viewBox" not in head:
        return svg_bytes  # no viewBox -> intrinsic size is all we have; leave it
    head = re.sub(r'\s(width|height)\s*=\s*"[^"]*"', "", head)
    head += ' width="512px" height="512px"'
    return (head + rest).encode("utf-8")


def _apply_fill(svg_bytes: bytes, hex_color: str) -> bytes:
    """Set a single fill color on a monochrome SVG (simple-icons style).

    simple-icons ship as a single ``<path>`` with no fill; we inject one on the
    root ``<svg>`` so the shape picks up the brand color.
    """
    text = svg_bytes.decode("utf-8", errors="replace")
    if "fill=" not in text.split(">", 1)[0]:
        text = re.sub(r"<svg\b", f'<svg fill="{hex_color}"', text, count=1)
    return text.encode("utf-8")
