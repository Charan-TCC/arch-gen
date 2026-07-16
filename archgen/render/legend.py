"""Auto-generated legends — derived from what's actually in the spec, so they
never go stale.

* Color legend: one row per provider present among the nodes.
* Connection legend: one row per line treatment present in a flow diagram.
* Abbreviation legend: one row per abbreviation declared and used.

Both are native shapes (swatch rects + text boxes) pinned to fixed corners.
python-pptx's chart-legend API can't express custom entries, so we hand-build them.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

from archgen.registry.resolver import IconResolver
from archgen.render.theme import Theme
from archgen.spec.lint import _ABBREV_RE
from archgen.spec.models import Diagram


def add_legends(slide, diagram: Diagram, theme: Theme, resolver: IconResolver | None = None) -> None:
    resolver = resolver or IconResolver()
    lcfg = theme["legend"]

    providers = _providers_present(diagram, resolver)
    connections = _flow_connections_present(diagram)
    used_abbrevs = _abbrevs_used(diagram)

    row_h = Emu(Pt(lcfg["font_pt"] + lcfg["row_gap_pt"]))
    pad = Emu(Inches(lcfg["pad_in"]))
    swatch = Emu(Inches(lcfg["swatch_in"]))
    title_h = Emu(Pt(lcfg["title_font_pt"] + 4))

    bottom = int(theme.slide_h) - Emu(Inches(0.12))
    left0 = int(theme.margin)

    # --- color legend (bottom-left) --------------------------------------
    if providers:
        width = Emu(Inches(2.2))
        height = title_h + row_h * len(providers) + pad
        top = bottom - height
        panel = _panel(slide, left0, top, width, height, theme)
        _title(slide, "Providers", left0 + pad, top + Emu(Pt(2)), width - 2 * pad, theme)
        y = top + title_h
        for prov in providers:
            pal = theme.provider(prov)
            sw = slide.shapes.add_shape(1, left0 + pad, y + Emu(Pt(1)), swatch, swatch)  # 1 = rectangle
            sw.fill.solid()
            sw.fill.fore_color.rgb = Theme.rgb(pal["hex"])
            sw.line.fill.background()
            sw.shadow.inherit = False
            _text(slide, pal["display"], left0 + pad + swatch + Emu(Pt(5)), y, width - swatch - 3 * pad, theme, lcfg["font_pt"])
            y += row_h

    # --- connection legend (bottom-center) -------------------------------
    if connections:
        width = Emu(Inches(3.35))
        height = title_h + row_h * len(connections) + pad
        left = (int(theme.slide_w) - width) // 2
        top = bottom - height
        _panel(slide, left, top, width, height, theme)
        _title(slide, "Connections", left + pad, top + Emu(Pt(2)), width - 2 * pad, theme)
        y = top + title_h
        for style, kind, display in connections:
            cfg = theme["connections"].get(kind, theme["edge"])
            x1 = left + pad
            x2 = x1 + Emu(Inches(0.72))
            cy = y + row_h // 2
            conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, cy, x2, cy)
            conn.line.color.rgb = Theme.rgb(cfg["color"])
            conn.line.width = Pt(theme["edge"]["line_pt"])
            ln = conn.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
            if style in ("dashed", "dotted"):
                ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash" if style == "dashed" else "sysDot"}))
            _text(slide, display, x2 + Emu(Pt(7)), y, width - (x2 - left) - 2 * pad, theme, lcfg["font_pt"])
            y += row_h

    # --- abbreviation legend (bottom-right) ------------------------------
    if used_abbrevs:
        width = Emu(Inches(3.0))
        height = title_h + row_h * len(used_abbrevs) + pad
        left = int(theme.slide_w) - int(theme.margin) - width
        top = bottom - height
        _panel(slide, left, top, width, height, theme)
        _title(slide, "Abbreviations", left + pad, top + Emu(Pt(2)), width - 2 * pad, theme)
        y = top + title_h
        for abbr in used_abbrevs:
            expansion = diagram.abbreviations[abbr]
            _text(slide, f"{abbr} — {expansion}", left + pad, y, width - 2 * pad, theme, lcfg["font_pt"])
            y += row_h


def _connection_legend_kinds(diagram: Diagram) -> list[str]:
    order = ["https", "diagnostics", "dns", "private_connection", "virtual_network_link", "traffic"]
    present = {e.kind.value for e in diagram.edges}
    kinds = [kind for kind in order if kind in present]
    if not kinds or kinds == ["traffic"]:
        return []
    return kinds


def connection_legend_height(diagram: Diagram) -> int:
    """EMU height of the network connection legend (0 if none renders).

    The renderer reserves exactly this much space at the top of a network slide
    so content is scaled to sit *below* the legend instead of under it.
    """
    kinds = _connection_legend_kinds(diagram)
    if not kinds:
        return 0
    return int(Emu(Pt(18)) + Emu(Pt(17)) * len(kinds) + Emu(Pt(9)) + Emu(Pt(34.5)))


def add_connection_legend(slide, diagram: Diagram, theme: Theme) -> None:
    """Derived traffic legend for semantic network diagrams."""
    kinds = _connection_legend_kinds(diagram)
    if not kinds:
        return

    left = Emu(Pt(21))
    top = Emu(Pt(12))
    width = Emu(Inches(3.46))
    row_h = Emu(Pt(17))
    pad = Emu(Pt(9))
    title_h = Emu(Pt(18))
    height = title_h + row_h * len(kinds) + pad
    _panel(slide, left, top, width, height, theme)
    _title(slide, "Connections", left + pad, top + Emu(Pt(2)), width - 2 * pad, theme)

    y = top + title_h + Emu(Pt(2))
    for kind in kinds:
        cfg = theme["connections"][kind]
        x1 = left + pad
        x2 = x1 + Emu(Inches(0.72))
        cy = y + row_h // 2
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, cy, x2, cy)
        conn.line.color.rgb = Theme.rgb(cfg["color"])
        conn.line.width = Pt(theme["network"]["connection_line_pt"])
        ln = conn.line._get_or_add_ln()
        tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"})
        ln.append(tail)
        _text(
            slide,
            cfg["display"],
            x2 + Emu(Pt(7)),
            y,
            width - (x2 - left) - pad,
            theme,
            theme["legend"]["font_pt"],
        )
        y += row_h

def _providers_present(diagram: Diagram, resolver: IconResolver) -> list[str]:
    order = ["azure", "aws", "gcp", "tech", "generic"]
    present = set()
    for n in diagram.nodes:
        if resolver.exists(n.icon_key):
            present.add(resolver.get(n.icon_key).provider)
        else:
            present.add(n.icon_key.split(".", 1)[0])
    return [p for p in order if p in present]


def _flow_connections_present(diagram: Diagram) -> list[tuple[str, str, str]]:
    """Return distinct flow-line semantics only when a legend adds value."""
    pairs = {(edge.style.value, edge.kind.value) for edge in diagram.edges}
    if len(pairs) < 2:
        return []
    order = [("solid", "traffic"), ("dashed", "https")]
    displays = {
        ("solid", "traffic"): "Request / data flow",
        ("dashed", "https"): "Authentication / RBAC",
    }
    ordered = [pair for pair in order if pair in pairs]
    ordered.extend(sorted(pairs - set(ordered)))
    return [
        (style, kind, displays.get((style, kind), f"{style.title()} {kind.replace('_', ' ')}"))
        for style, kind in ordered
    ]


def _abbrevs_used(diagram: Diagram) -> list[str]:
    used: set[str] = set()
    for text in [diagram.title, *(n.label for n in diagram.nodes), *(g.label for g in diagram.groups)]:
        used |= set(_ABBREV_RE.findall(text))
    return sorted(a for a in used if a in diagram.abbreviations)


def _panel(slide, left, top, width, height, theme: Theme):
    lcfg = theme["legend"]
    panel = slide.shapes.add_shape(1, left, top, width, height)  # rectangle
    panel.fill.solid()
    panel.fill.fore_color.rgb = Theme.rgb(lcfg["box_fill"])
    panel.line.color.rgb = Theme.rgb(lcfg["box_line"])
    panel.line.width = Pt(0.75)
    panel.shadow.inherit = False
    style = panel._element.find(qn("p:style"))
    if style is not None:
        panel._element.remove(style)
    sp_pr = panel._element.spPr
    for tag in ("a:effectLst", "a:effectDag"):
        existing = sp_pr.find(qn(tag))
        if existing is not None:
            sp_pr.remove(existing)
    sp_pr.append(OxmlElement("a:effectLst"))
    return panel


def _title(slide, text, left, top, width, theme: Theme):
    lcfg = theme["legend"]
    box = slide.shapes.add_textbox(left, top, width, Emu(Pt(lcfg["title_font_pt"] + 3)))
    tf = box.text_frame
    tf.margin_top = tf.margin_bottom = 0
    tf.margin_left = tf.margin_right = 0
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = text
    r = p.runs[0]
    r.font.bold = True
    r.font.size = Pt(lcfg["title_font_pt"])
    r.font.name = theme.font
    r.font.color.rgb = Theme.rgb(theme["group"]["title_color"])


def _text(slide, text, left, top, width, theme: Theme, size_pt: float):
    box = slide.shapes.add_textbox(left, top, width, Emu(Pt(size_pt + 4)))
    tf = box.text_frame
    tf.margin_top = tf.margin_bottom = 0
    tf.margin_left = tf.margin_right = 0
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.text = text
    r = p.runs[0]
    r.font.size = Pt(size_pt)
    r.font.name = theme.font
    r.font.color.rgb = Theme.rgb(theme["node"]["label_color"])
