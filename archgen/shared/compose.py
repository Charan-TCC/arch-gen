"""compose — shared native-shape drawing kit (python-pptx only).

SHARED FILE. An identical copy lives in both repos:
    PPT Gen  pptgen/shared/compose.py
    Arch Gen archgen/shared/compose.py
A parity test in each repo fails if the two drift. Edit both together (or copy one
over the other). Contract: docs/COMPOSE_API.md.

Merge-readiness rules — keep these true so the repos can be merged later by
deleting one copy and fixing imports:
  * import nothing from pptgen/ or archgen/ — python-pptx and stdlib only;
  * colours, fonts and shape language come in through `Style`;
  * icons come in through `IconSource` (anything with `.png(key, size_in)`);
  * every coordinate and size is in INCHES.

Invariants (Arch Gen CLAUDE.md §7, PPT Gen RULES_CARD):
  * an icon and its label are ONE group shape (`icon_card`);
  * text boxes never auto-grow (auto_size NONE, word_wrap on);
  * no shadows or effects on anything drawn here.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Iterable, Protocol, Sequence

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

__all__ = [
    "Style", "IconSource", "NoIcons",
    "kill_effects", "set_background", "rrect", "text", "zone", "band",
    "connector", "callout", "badge", "icon_card",
    "text_height", "snap", "row", "column", "port", "link",
    "cover", "COVER_LAYOUTS", "COVER_VISUALS",
]

SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5


# --------------------------------------------------------------------------- style

@dataclass
class Style:
    """Everything visual a builder needs. Build it from a theme/preset dict with
    `Style.from_tokens`, or construct it directly in a bespoke script."""

    background: str = "FFFFFF"
    surface: str = "F5F7FA"
    surface_alt: str = "E9EDF2"
    border: str = "B4BDCA"
    ink: str = "0F1B2D"          # primary text
    muted: str = "4A5668"        # secondary text
    inverse: str = "FFFFFF"      # text on accent fills
    accent: str = "0B3D91"
    accent2: str = "0F766E"
    font_heading: str = "Aptos"
    font_body: str = "Aptos"
    corner_radius: float = 0.04  # 0 = sharp
    line_pt: float = 0.75
    card: str = "outline"        # outline | filled | block | rule-top
    cover_layout: str = "split-right-visual"
    dark: bool = False
    size_cover: float = 40       # type scale (pt)
    size_section: float = 32
    size_title: float = 26
    size_subtitle: float = 18
    size_body: float = 14
    size_caption: float = 11
    size_footer: float = 9
    extra: dict = field(default_factory=dict)

    @staticmethod
    def _h(v: str) -> str:
        return v.strip().lstrip("#").upper()

    @classmethod
    def from_tokens(cls, d: dict) -> "Style":
        """From a theme.json / styles/*.json dict (the shared preset format)."""
        c = {k: cls._h(v) for k, v in d["color"].items()}
        f = d.get("font", {})
        s = d.get("style", {})
        ts = d.get("type_scale", {})
        sizes = {f"size_{k}": float(ts[k]) for k in
                 ("cover", "section", "title", "subtitle", "body", "caption", "footer") if k in ts}
        return cls(**sizes,
            background=c["background"], surface=c["surface"],
            surface_alt=c["surface-alt"], border=c["border"],
            ink=c["text-primary"], muted=c["text-secondary"],
            inverse=c["text-inverse"], accent=c["accent"],
            accent2=c["accent-secondary"],
            font_heading=f.get("heading", f.get("primary", "Aptos")),
            font_body=f.get("body", f.get("primary", "Aptos")),
            corner_radius=float(s.get("corner_radius", 0.04)),
            line_pt=float(s.get("line_pt", 0.75)),
            card=s.get("card", "outline"),
            cover_layout=s.get("cover", "split-right-visual"),
            dark=d.get("mode") == "dark",
        )


class IconSource(Protocol):
    def png(self, key: str, size_in: float, tone: str = "auto") -> str | None:
        """Path to a PNG for `key`, or None if unavailable.
        tone: "auto" (normal) or "inverse" (icon sits on an accent fill)."""


class NoIcons:
    """IconSource that resolves nothing — visuals fall back to native shapes."""

    def png(self, key: str, size_in: float, tone: str = "auto") -> str | None:
        return None


def _rgb(h: str) -> RGBColor:
    return RGBColor.from_string(Style._h(h))


# --------------------------------------------------------------------------- primitives

def kill_effects(shp) -> None:
    """Strip theme shadow/effects (pictures and groups inherit them too)."""
    try:
        shp.shadow.inherit = False
    except Exception:
        pass
    el = shp._element
    st = el.find(qn("p:style"))
    if st is not None:
        el.remove(st)
    sppr = el.find(qn("p:spPr"))
    if sppr is not None and sppr.find(qn("a:effectLst")) is None:
        sppr.append(sppr.makeelement(qn("a:effectLst"), {}))


def set_background(slide, hex_color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(hex_color)


def rrect(slide, x, y, w, h, *, fill: str | None, line: str | None,
          line_pt: float = 0.75, radius: float = 0.04, shape=None):
    """Rounded (or sharp, radius 0) rectangle. fill/line None = transparent."""
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius > 0 else MSO_SHAPE.RECTANGLE)
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_pt)
    kill_effects(shp)
    return shp


Run = tuple  # (text, size_pt, bold, hex_color)


def text_height(paras, w: float, size: float = 14, line: float = 1.2) -> float:
    """Estimated height (in) text needs at width w — use it to size boxes to their
    text instead of guessing (oversized boxes break vertical centring)."""
    if isinstance(paras, str):
        paras = [[(paras, size, False, "")]]
    total = 0.0
    for para in paras:
        txt = "".join(r[0] for r in para)
        sz = max((r[1] for r in para), default=size)
        per_line = max(1, int(max(w - 0.05, 0.1) * 72 / (sz * 0.52)))
        lines = max(1, -(-len(txt) // per_line))
        total += lines * sz * line / 72
    return round(total + 0.04, 3)


def snap(v: float, step: float = 1 / 32) -> float:
    """Snap a coordinate to a 1/32in grid so peers land on identical values."""
    return round(round(v / step) * step, 4)


def row(n: int, x: float, w: float, item_w: float) -> list[float]:
    """Left edges for n items of width item_w spread evenly across [x, x+w]."""
    if n <= 1:
        return [snap(x + (w - item_w) / 2)]
    gap = (w - n * item_w) / (n - 1)
    return [snap(x + i * (item_w + gap)) for i in range(n)]


def column(n: int, y: float, h: float, item_h: float) -> list[float]:
    """Top edges for n items of height item_h spread evenly down [y, y+h]."""
    return row(n, y, h, item_h)


_SIDES = {"top": 0, "left": 1, "bottom": 2, "right": 3}


def port(shape, side: str, offset: float = 0.0) -> tuple[float, float]:
    """(x, y) in inches of the middle of a shape's side, shifted along the side by
    `offset` inches (for parallel arrows into one shape)."""
    l, t = shape.left / 914400, shape.top / 914400
    w, h = shape.width / 914400, shape.height / 914400
    return {"top": (l + w / 2 + offset, t), "bottom": (l + w / 2 + offset, t + h),
            "left": (l, t + h / 2 + offset), "right": (l + w, t + h / 2 + offset)}[side]


def link(slide, a, b, *, a_side: str = "right", b_side: str = "left",
         lane: float = 0.0, color: str = "5B7183", pt: float = 1.5,
         dashed: bool = False, elbow: bool | None = None, tail: bool = False,
         bend: float = 0.5):
    """Arrow from shape a to shape b, attached exactly to their edges.
    Elbows leaving top/bottom route vertical-first; `bend` (0..1 from a) places
    the middle run — give parallel elbows different bends.

    lane == 0: the connector is GLUED to both shapes (moves with them in
    PowerPoint). lane != 0: endpoints are offset along both sides by `lane` inches
    so parallel arrows between the same columns never sit on top of each other.
    """
    x1, y1 = port(a, a_side, lane)
    x2, y2 = port(b, b_side, lane)
    if elbow is None:
        elbow = abs(x1 - x2) > 0.01 and abs(y1 - y2) > 0.01
    cn = connector(slide, x1, y1, x2, y2, color=color, pt=pt, dashed=dashed,
                   elbow=elbow, tail=tail)
    if lane == 0:
        try:
            cn.begin_connect(a, _SIDES[a_side])
            cn.end_connect(b, _SIDES[b_side])
        except Exception:
            pass
    if elbow:
        route_elbow(cn, Inches(x1), Inches(y1), Inches(x2), Inches(y2),
                    vertical=a_side in ("top", "bottom"), bend=bend)
    return cn


def text(slide, x, y, w, h, paras: Sequence[Sequence[Run]] | str, *,
         font: str = "Aptos", size: float = 14, color: str = "1F2937",
         bold: bool = False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         wrap: bool = True, line_spacing: float | None = None):
    """Text box that never auto-grows. `paras` is either a plain string or a list
    of paragraphs, each a list of (text, size_pt, bold, hex) runs.
    h=None sizes the box to its text (text_height) — prefer it."""
    if h is None:
        h = text_height(paras, w, size)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    if isinstance(paras, str):
        paras = [[(paras, size, bold, color)]]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for (txt, sz, b, col) in para:
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(sz)
            r.font.bold = b
            r.font.name = font
            r.font.color.rgb = _rgb(col)
    kill_effects(box)
    return box


def zone(slide, st: Style, x, y, w, h, title: str = "", *, fill: str | None = None,
         line: str | None = None, title_size: float = 11):
    """A labelled container. Returns the container shape (title is a sibling)."""
    shp = rrect(slide, x, y, w, h, fill=fill or st.surface, line=line or st.border,
                line_pt=st.line_pt, radius=st.corner_radius)
    if title:
        text(slide, x + 0.12, y + 0.08, w - 0.24, 0.3, title, font=st.font_heading,
             size=title_size, bold=True, color=st.muted)
    return shp


def band(slide, st: Style, x, y, w, h, title: str = "", *, fill: str | None = None,
         title_color: str | None = None, title_size: float = 11):
    """A full-width strip (e.g. a platform/security layer under several zones)."""
    shp = rrect(slide, x, y, w, h, fill=fill or st.surface_alt, line=None,
                radius=min(st.corner_radius, 0.02))
    if title:
        text(slide, x + 0.15, y, w - 0.3, h, title, font=st.font_heading,
             size=title_size, bold=True, color=title_color or st.ink,
             anchor=MSO_ANCHOR.MIDDLE)
    return shp


def connector(slide, x1, y1, x2, y2, *, color: str = "5B7183", pt: float = 1.5,
              dashed: bool = False, elbow: bool = False, head: bool = True,
              tail: bool = False):
    """Native connector. head = arrow at (x2,y2); tail = arrow at (x1,y1)."""
    kind = (MSO_CONNECTOR.ELBOW if (elbow and x1 != x2 and y1 != y2)
            else MSO_CONNECTOR.STRAIGHT)
    cn = slide.shapes.add_connector(kind, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = _rgb(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if tail:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if head:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    kill_effects(cn)
    return cn


def route_elbow(cn, x1: int, y1: int, x2: int, y2: int, *, vertical: bool,
                bend: float = 0.5) -> None:
    """Re-shape an elbow connector (EMU coords) so it routes from (x1,y1) to
    (x2,y2) and bends at `bend` (0..1, measured from the start).

    vertical=False: horizontal-vertical-horizontal (leave/enter left|right sides).
    vertical=True:  vertical-horizontal-vertical (leave/enter top|bottom sides) —
    done with a 90-degree-rotated bentConnector3, PowerPoint's own encoding, so
    the arrow never runs along a card edge. Staggering `bend` per arrow keeps
    parallel arrows' middle runs on separate lines."""
    dx, dy = x2 - x1, y2 - y1
    if dx == 0 or dy == 0:
        return
    xfrm = cn._element.spPr.get_or_add_xfrm()
    cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
    if vertical:
        w, h = abs(dy), abs(dx)
        xfrm.set("rot", "5400000")
        flip_h, flip_v = dy < 0, dx > 0
    else:
        w, h = abs(dx), abs(dy)
        xfrm.attrib.pop("rot", None)
        flip_h, flip_v = dx < 0, dy < 0
    for attr, on in (("flipH", flip_h), ("flipV", flip_v)):
        if on:
            xfrm.set(attr, "1")
        else:
            xfrm.attrib.pop(attr, None)
    off, ext = xfrm.get_or_add_off(), xfrm.get_or_add_ext()
    off.x, off.y = int(round(cx - w / 2)), int(round(cy - h / 2))
    ext.cx, ext.cy = int(w), int(h)
    geom = cn._element.spPr.find(qn("a:prstGeom"))
    if geom is not None:
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = geom.makeelement(qn("a:avLst"), {})
            geom.append(av)
        for gd in list(av):
            av.remove(gd)
        av.append(av.makeelement(qn("a:gd"), {"name": "adj1",
                                              "fmla": f"val {int(round(bend * 100000))}"}))


def callout(slide, st: Style, cx, cy, number: int | str, *, d: float = 0.32,
            fill: str | None = None):
    """Numbered circular marker centred on (cx, cy)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2),
                                 Inches(d), Inches(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill or st.accent)
    shp.line.fill.background()
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(number)
    r.font.size = Pt(max(9, round(d * 30)))
    r.font.bold = True
    r.font.name = st.font_body
    r.font.color.rgb = _rgb(st.inverse)
    kill_effects(shp)
    return shp


def badge(slide, st: Style, x, y, label: str, *, w: float | None = None,
          h: float = 0.28, fill: str | None = None, color: str | None = None,
          size: float = 9):
    """Small pill label (e.g. "CI/CD Pipeline")."""
    w = w or max(0.6, 0.075 * len(label) + 0.3)
    shp = rrect(slide, x, y, w, h, fill=fill or st.accent, line=None, radius=0.5)
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.name = st.font_body
    r.font.color.rgb = _rgb(color or st.inverse)
    return shp


def _picture(slide, path: str, x, y, size, *, keep_aspect: bool = False):
    """Square picture, or (keep_aspect) height=size with natural width — for logos."""
    if keep_aspect:
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), height=Inches(size))
    else:
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(size), Inches(size))
    kill_effects(pic)
    return pic


def icon_card(slide, st: Style, icons: IconSource, key: str | None, label: str,
              x, y, w, h, *, sublabel: str = "", icon_in: float = 0.42,
              card: str | None = None, label_size: float = 10):
    """Icon + label (+ optional sublabel) inside a card, as ONE group shape.

    card: outline | filled | block | none (default: st.card). The label box is
    card-bounded so it never spills into the container below.
    """
    card = card or st.card
    grp = slide.shapes.add_group_shape()
    sh = grp.shapes
    if card != "none":
        fill = {"filled": st.surface, "block": st.surface_alt}.get(card, st.background)
        line = None if card in ("filled", "block") else st.border
        bx = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if st.corner_radius > 0 else MSO_SHAPE.RECTANGLE,
                          Inches(x), Inches(y), Inches(w), Inches(h))
        if st.corner_radius > 0:
            try:
                bx.adjustments[0] = st.corner_radius
            except Exception:
                pass
        bx.fill.solid()
        bx.fill.fore_color.rgb = _rgb(fill)
        if line:
            bx.line.color.rgb = _rgb(line)
            bx.line.width = Pt(st.line_pt)
        else:
            bx.line.fill.background()
        kill_effects(bx)
    png = icons.png(key, icon_in) if key else None
    top = y + 0.1
    if png:
        pic = sh.add_picture(str(png), Inches(x + (w - icon_in) / 2), Inches(top),
                             Inches(icon_in), Inches(icon_in))
        kill_effects(pic)
        top += icon_in + 0.06
    lab_h = max(0.2, y + h - top - 0.06)
    tb = sh.add_textbox(Inches(x + 0.06), Inches(top), Inches(w - 0.12), Inches(lab_h))
    kill_effects(tb)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(label_size)
    r.font.bold = True
    r.font.name = st.font_body
    r.font.color.rgb = _rgb(st.ink)
    if sublabel:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = sublabel
        r2.font.size = Pt(max(8, label_size - 2))
        r2.font.name = st.font_body
        r2.font.color.rgb = _rgb(st.muted)
    kill_effects(grp)
    return grp


# --------------------------------------------------------------------------- cover visuals
# Each draws ONE simple infographic in the box (x, y, w, h) using topic icons.
# All native shapes + icon pictures, grouped so the visual moves as a unit.

def _disc(shapes, st, cx, cy, d, fill, line=None):
    s = shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2),
                         Inches(d), Inches(d))
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(fill)
    if line:
        s.line.color.rgb = _rgb(line)
        s.line.width = Pt(st.line_pt)
    else:
        s.line.fill.background()
    kill_effects(s)
    return s


def _ico(shapes, icons, key, cx, cy, size, tone="auto"):
    png = icons.png(key, size, tone) if key else None
    if png:
        p = shapes.add_picture(str(png), Inches(cx - size / 2), Inches(cy - size / 2),
                               Inches(size), Inches(size))
        kill_effects(p)
        return p
    return None


def _line(shapes, st, x1, y1, x2, y2, color, pt=1.25, dashed=False):
    cn = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                              Inches(x2), Inches(y2))
    cn.line.color.rgb = _rgb(color)
    cn.line.width = Pt(pt)
    if dashed:
        ln = cn.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    kill_effects(cn)
    return cn


def visual_hub_spoke(slide, st, icons, keys, x, y, w, h):
    """Centre icon with satellites — platforms, ecosystems, integrations."""
    g = slide.shapes.add_group_shape()
    sh = g.shapes
    cx, cy = x + w / 2, y + h / 2
    r = min(w, h) * 0.36
    sat = list(keys[1:7]) or [None] * 5
    n = len(sat)
    hub_d, sat_d = min(w, h) * 0.30, min(w, h) * 0.17
    _disc(sh, st, cx, cy, r * 2, st.background, line=st.border)  # orbit ring
    for i, k in enumerate(sat):
        a = -math.pi / 2 + 2 * math.pi * i / n
        sx, sy = cx + r * math.cos(a), cy + r * math.sin(a)
        _line(sh, st, cx, cy, sx, sy, st.border, dashed=True)
    _disc(sh, st, cx, cy, hub_d, st.accent)
    _ico(sh, icons, keys[0] if keys else None, cx, cy, hub_d * 0.52, "inverse")
    for i, k in enumerate(sat):
        a = -math.pi / 2 + 2 * math.pi * i / n
        sx, sy = cx + r * math.cos(a), cy + r * math.sin(a)
        _disc(sh, st, sx, sy, sat_d, st.surface, line=st.border)
        _ico(sh, icons, k, sx, sy, sat_d * 0.56)
    kill_effects(g)
    return g


def visual_ring(slide, st, icons, keys, x, y, w, h):
    """Icons on a ring with an accent arc — cycles, lifecycles, continuous flows."""
    g = slide.shapes.add_group_shape()
    sh = g.shapes
    cx, cy = x + w / 2, y + h / 2
    d = min(w, h) * 0.78
    ring = sh.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    ring.fill.background()
    ring.line.color.rgb = _rgb(st.accent)
    ring.line.width = Pt(3)
    kill_effects(ring)
    _disc(sh, st, cx, cy, d * 0.42, st.surface_alt)
    _ico(sh, icons, keys[0] if keys else None, cx, cy, d * 0.2)
    ks = list(keys[1:7]) or [None] * 5
    n = len(ks)
    nd = min(w, h) * 0.18
    for i, k in enumerate(ks):
        a = -math.pi / 2 + 2 * math.pi * i / n
        px, py = cx + d / 2 * math.cos(a), cy + d / 2 * math.sin(a)
        _disc(sh, st, px, py, nd, st.background, line=st.accent)
        _ico(sh, icons, k, px, py, nd * 0.55)
    kill_effects(g)
    return g


def visual_icon_cluster(slide, st, icons, keys, x, y, w, h):
    """Staggered honeycomb of icon tiles — a capability set at a glance."""
    g = slide.shapes.add_group_shape()
    sh = g.shapes
    ks = list(keys[:7]) or [None] * 7
    cols = 3
    rows = (len(ks) + cols - 1) // cols
    t = min(w / ((cols + 0.5) * 1.12), h / (rows * 1.12))
    gap = t * 0.12
    for i, k in enumerate(ks):
        row, col = divmod(i, cols)
        ox = (t + gap) / 2 if row % 2 else 0
        tx = x + (w - (cols * (t + gap) + (t + gap) / 2)) / 2 + col * (t + gap) + ox
        ty = y + (h - rows * (t + gap)) / 2 + row * (t + gap)
        hexa = sh.add_shape(MSO_SHAPE.HEXAGON, Inches(tx), Inches(ty), Inches(t), Inches(t * 0.9))
        hexa.fill.solid()
        hexa.fill.fore_color.rgb = _rgb(st.accent if i == 0 else (st.surface if i % 2 else st.surface_alt))
        hexa.line.fill.background()
        kill_effects(hexa)
        _ico(sh, icons, k, tx + t / 2, ty + t * 0.45, t * 0.46,
             "inverse" if i == 0 else "auto")
    kill_effects(g)
    return g


def visual_stack(slide, st, icons, keys, x, y, w, h):
    """Layered platform stack — architectures, tiers, maturity layers."""
    g = slide.shapes.add_group_shape()
    sh = g.shapes
    ks = list(keys[:4]) or [None] * 4
    n = len(ks)
    lh = h * 0.8 / n
    for i, k in enumerate(ks):
        inset = (n - 1 - i) * w * 0.06
        ly = y + h * 0.1 + i * lh
        layer = sh.add_shape(MSO_SHAPE.PARALLELOGRAM if st.corner_radius == 0 else MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(x + inset), Inches(ly), Inches(w - 2 * inset), Inches(lh * 0.82))
        layer.fill.solid()
        layer.fill.fore_color.rgb = _rgb(st.accent if i == n - 1 else (st.surface_alt if i % 2 else st.surface))
        layer.line.fill.background()
        kill_effects(layer)
        _ico(sh, icons, k, x + w / 2, ly + lh * 0.41, lh * 0.5,
             "inverse" if i == n - 1 else "auto")
    kill_effects(g)
    return g


def visual_flow_strip(slide, st, icons, keys, x, y, w, h):
    """Left-to-right icon chain with arrows — pipelines, journeys, migrations."""
    g = slide.shapes.add_group_shape()
    sh = g.shapes
    ks = list(keys[:5]) or [None] * 4
    n = len(ks)
    d = min(h * 0.45, w / (n * 1.6))
    step = (w - d) / max(n - 1, 1)
    cy = y + h / 2
    for i in range(n - 1):
        cn = _line(sh, st, x + d / 2 + i * step + d / 2, cy, x + d / 2 + (i + 1) * step - d / 2, cy,
                   st.accent, pt=2)
        ln = cn.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    for i, k in enumerate(ks):
        cx = x + d / 2 + i * step
        _disc(sh, st, cx, cy, d, st.accent if i == n - 1 else st.surface,
              line=None if i == n - 1 else st.border)
        _ico(sh, icons, k, cx, cy, d * 0.52, "inverse" if i == n - 1 else "auto")
    kill_effects(g)
    return g


COVER_VISUALS = {
    "hub-spoke": visual_hub_spoke,
    "ring": visual_ring,
    "icon-cluster": visual_icon_cluster,
    "stack": visual_stack,
    "flow-strip": visual_flow_strip,
}

COVER_LAYOUTS = ("split-right-visual", "left-text-right-visual", "centered-visual",
                 "color-block-left")


# --------------------------------------------------------------------------- cover

def cover(slide, st: Style, icons: IconSource, *, heading: str, subheading: str,
          logos: Iterable[str] = (), visual: str = "hub-spoke",
          visual_keys: Sequence[str] = (), footer: str = "", layout: str | None = None,
          motif_png: str | None = None, slide_w: float = SLIDE_W_IN,
          slide_h: float = SLIDE_H_IN):
    """The ONLY cover a deck needs: heading, one-line subheading, logos, one visual.

    heading      one line (<= ~60 chars)
    subheading   one line (<= ~90 chars)
    logos        icon keys for company / client / platform logos (max 3 used)
    visual       one of COVER_VISUALS, drawn from `visual_keys` (topic icons)
    motif_png    brand mode: a brand motif image used INSTEAD of `visual`
    footer       optional single small line (date · presenter · confidentiality)
    layout       one of COVER_LAYOUTS (default: st.cover_layout)
    """
    if visual not in COVER_VISUALS and motif_png is None:
        raise ValueError(f"unknown cover visual {visual!r}; one of {sorted(COVER_VISUALS)}")
    layout = layout or st.cover_layout
    if layout not in COVER_LAYOUTS:
        raise ValueError(f"unknown cover layout {layout!r}; one of {COVER_LAYOUTS}")
    W, H = slide_w, slide_h
    m = 0.9   # 64pt — the PPT Gen grid margin
    set_background(slide, st.background)
    head_col, sub_col = st.ink, st.muted

    if layout == "centered-visual":
        vbox = (W / 2 - 1.9, 0.55, 3.8, 3.4)
        tbox = (m, 4.15, W - 2 * m)
        align = PP_ALIGN.CENTER
    else:
        vbox = (W * 0.56, 0.9, W * 0.44 - m, H - 2.0)
        tbox = (m, 2.2, W * 0.5 - m)
        align = PP_ALIGN.LEFT
        if layout == "color-block-left":
            rrect(slide, 0, 0, W * 0.5, H, fill=st.accent, line=None, radius=0)
            head_col, sub_col = st.inverse, st.inverse
        elif layout == "left-text-right-visual":
            rrect(slide, m, tbox[1] - 0.3, 0.9, 0.06, fill=st.accent, line=None, radius=0)

    # visual
    if motif_png:
        s = min(vbox[2], vbox[3])
        _picture(slide, motif_png, vbox[0] + (vbox[2] - s) / 2, vbox[1] + (vbox[3] - s) / 2, s)
    else:
        COVER_VISUALS[visual](slide, st, icons, list(visual_keys), *vbox)

    # heading + subheading
    tx, ty, tw = tbox
    # Largest on-scale size that keeps the heading to two lines.
    hs = st.size_title
    for cand in (st.size_cover, st.size_section, st.size_title):
        per_line = tw * 72 / (cand * 0.62)
        if math.ceil(len(heading) / per_line) <= 2:
            hs = cand
            break
    per_line = tw * 72 / (hs * 0.55)
    hh = min(2, math.ceil(len(heading) / per_line)) * hs * 1.2 / 72 + 0.05
    text(slide, tx, ty, tw, hh, heading, font=st.font_heading, size=hs, bold=True,
         color=head_col, align=align, anchor=MSO_ANCHOR.BOTTOM, line_spacing=0.95)
    text(slide, tx, ty + hh + 0.15, tw, st.size_subtitle * 1.25 / 72 + 0.05, subheading, font=st.font_body,
         size=st.size_subtitle, color=sub_col, align=align, wrap=False)

    # logos: top-left row (or centred on a centred layout)
    lg = [k for k in logos][:3]
    size, gap = 0.45, 0.3
    pngs = [p for p in (icons.png(k, size) for k in lg) if p]
    ly = H - 1.05 if layout == "centered-visual" else 0.6
    pics = []
    lx = m
    for png in pngs:
        pic = _picture(slide, png, lx, ly, size, keep_aspect=True)
        pics.append(pic)
        lx += pic.width / 914400 + gap
    if layout == "centered-visual" and pics:   # centre the row
        row_w = lx - gap - m
        shift = Inches((W - row_w) / 2 - m)
        for pic in pics:
            pic.left = pic.left + shift

    if footer:
        fy = H - 0.45
        fw = (W * 0.5 - 2 * m) if layout == "color-block-left" else W - 2 * m
        text(slide, m, fy, fw, st.size_footer * 1.25 / 72 + 0.05, footer, font=st.font_body, size=st.size_footer,
             color=sub_col,
             align=PP_ALIGN.CENTER if layout == "centered-visual" else PP_ALIGN.LEFT,
             wrap=False)
    return slide
