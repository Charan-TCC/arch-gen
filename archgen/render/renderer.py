"""Deterministic renderer: (Diagram + Layout) -> editable .pptx.

Everything here is aesthetics. Icons are embedded as picture objects (EMF vector
when a vector converter is available, else PNG), uniform-scaled and never
distorted. Every other element — container tints, node boxes, labels, arrows,
legends — is a native PowerPoint shape, so the whole slide is editable.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Pt

from archgen.layout.engine import compute_layout
from archgen.layout.geometry import Rect, Transform
from archgen.registry import emf
from archgen.registry.resolver import IconResolver
from archgen.render.legend import (
    add_connection_legend,
    add_legends,
    connection_legend_height,
    legends_reserve_height,
)
from archgen.render.theme import Theme, load_theme
from archgen.spec.models import Diagram, DiagramKind, GroupKind, NodeRole, Provider


@dataclass
class _EdgeBinding:
    """Deferred connector-to-node binding: connectors are drawn before nodes
    (so routes sit behind cards), so we record what to bind and apply it once
    the node shapes exist. See Renderer._draw_edges / _bind_edges."""

    conn: object
    src: str
    dst: str
    src_idx: int
    dst_idx: int
    coords: tuple[tuple[int, int], tuple[int, int]]


class Renderer:
    def __init__(self, theme: Theme | None = None, resolver: IconResolver | None = None):
        self.theme = theme or load_theme()
        self.resolver = resolver or IconResolver()
        self._layered = False
        # Post-layout quality findings for the last render (flow only); lets
        # callers iterate on a text report instead of rasterizing every attempt.
        self.diagnostics: list = []

    def render(
        self,
        diagram: Diagram,
        out_path: str | Path,
        engine: str = "auto",
        auto_direction: bool = False,
        expand_slide: bool | None = None,
        quality: str = "draft",
    ) -> Path:
        prs = Presentation()
        network = diagram.kind == DiagramKind.network
        # Layered mode packs groups into parallel bands with renderer-owned card
        # sizing; it is a flow-only, graphviz-only path.
        self._layered = (not network) and diagram.layered
        if network:
            prs.slide_width = Inches_val(self.theme["network"]["slide_width_in"])
            prs.slide_height = Inches_val(self.theme["network"]["slide_height_in"])
        else:
            prs.slide_width = int(self.theme.slide_w)
            prs.slide_height = int(self.theme.slide_h)

        # Slide bounds are needed to clamp free-floating labels on-slide.
        self._slide_w = int(prs.slide_width)
        self._slide_h = int(prs.slide_height)

        content_height = int(prs.slide_height)
        top_reserve = 0
        legend_reserve = 0
        if not network:
            # Reserve a true footer band for the derived legends. The reserve is
            # computed from the legends that will actually render (the
            # abbreviation panel grows a row per entry) — a fixed band lets tall
            # legends swallow the bottom-right of the diagram.
            legend_reserve = max(
                Inches_val(self.theme["slide"]["bottom_legend_reserve_in"]),
                legends_reserve_height(diagram, self.theme, self.resolver),
            )
            content_height -= legend_reserve
            # Layered boards use tighter margins (and a title band) so the packed
            # tiers get the vertical room the airy default would eat.
            margin_emu = int(self.theme.margin) + (Emu(Pt(6)) if self._layered else Emu(Pt(28)))
            if self._layered:
                top_reserve = Emu(Pt(40))
        else:
            # Reserve exactly the connection-legend strip at the top so the
            # legend never overlaps the first container.
            top_reserve = connection_legend_height(diagram)
            margin_emu = Emu(Pt(self.theme["network"]["layout_margin_pt"]))

        # For layered diagrams the renderer picks the card size that fits the
        # densest tier into the available area, then hands it to the layout.
        # Card size depends on direction (columns vs rows), so it is recomputed
        # per orientation candidate below.
        def card_size_for(diag):
            if not self._layered:
                return None, None
            from archgen.layout.graphviz_layout import layered_card_size

            avail_w_in = (int(prs.slide_width) - 2 * margin_emu) / 914400
            avail_h_in = (content_height - 2 * margin_emu - top_reserve) / 914400
            return layered_card_size(diag, avail_w_in, avail_h_in)

        def make_tf(la):
            return Transform(
                la.graph_bb,
                slide_w_emu=int(prs.slide_width),
                slide_h_emu=content_height,
                margin_emu=margin_emu,
                top_reserve_emu=top_reserve,
            )

        node_w, node_h = card_size_for(diagram)
        layout, engine_used = compute_layout(diagram, engine=engine, node_w=node_w, node_h=node_h)
        tf = make_tf(layout)

        # Deterministic direction choice: lay out both orientations (dot runs in
        # milliseconds) and keep the one with the better fit — a decision the
        # tool can make so the agent never burns an iteration discovering it.
        direction_flipped = False
        auto_direction_tried = False
        if auto_direction and not network and engine_used == "graphviz":
            auto_direction_tried = True
            flipped_dir = "LR" if diagram.direction == "TB" else "TB"
            alt_diagram = diagram.model_copy(update={"direction": flipped_dir})
            alt_w, alt_h = card_size_for(alt_diagram)
            alt_layout, _ = compute_layout(alt_diagram, engine=engine, node_w=alt_w, node_h=alt_h)
            alt_tf = make_tf(alt_layout)
            if alt_tf.scale > tf.scale * 1.15:
                diagram, layout, tf = alt_diagram, alt_layout, alt_tf
                direction_flipped = True

        # Readability floor: content that would shrink below MIN_READABLE_SCALE
        # gets a BIGGER SLIDE instead of smaller text — labels have exactly one
        # readable size, slides can be any size. Whether the slide MAY grow is
        # the user's `page` choice: "free" grows, "16:9" locks the deck slide
        # and the layout report demands reshaping instead. Growth is capped so
        # a pathological spec still fails loudly rather than yielding a
        # wall-sized slide. Layered mode self-sizes its cards and is exempt.
        if expand_slide is None:
            expand_slide = diagram.page == "free"
        slide_expanded_from = None
        if expand_slide and not network and not self._layered:
            from archgen.layout.diagnostics import MIN_READABLE_SCALE
            from archgen.layout.geometry import EMU_PER_POINT

            if tf.scale < MIN_READABLE_SCALE:
                slide_expanded_from = tf.scale
                target = MIN_READABLE_SCALE * 1.02  # epsilon above the floor
                x0, y0, x1, y1 = layout.graph_bb
                need_w = int((x1 - x0) * EMU_PER_POINT * target) + 2 * margin_emu
                need_h = (
                    int((y1 - y0) * EMU_PER_POINT * target)
                    + 2 * margin_emu
                    + top_reserve
                    + legend_reserve
                )
                cap_w, cap_h = 3 * int(self.theme.slide_w), 3 * int(self.theme.slide_h)
                prs.slide_width = max(int(self.theme.slide_w), min(need_w, cap_w))
                prs.slide_height = max(int(self.theme.slide_h), min(need_h, cap_h))
                self._slide_w = int(prs.slide_width)
                self._slide_h = int(prs.slide_height)
                content_height = int(prs.slide_height) - legend_reserve
                tf = make_tf(layout)

        self.diagnostics = []
        if not network:
            from archgen.layout.diagnostics import diagnose

            self.diagnostics = diagnose(
                diagram,
                layout,
                scale=tf.scale,
                avail_w_emu=int(prs.slide_width) - 2 * margin_emu,
                avail_h_emu=content_height - 2 * margin_emu - top_reserve,
                label_cap_in=self.theme["node"]["label_max_width_in"],
                direction_flipped=direction_flipped,
                auto_direction_tried=auto_direction_tried,
                quality=quality,
            )
            if slide_expanded_from is not None:
                from archgen.layout.diagnostics import Finding

                self.diagnostics.append(
                    Finding(
                        "info",
                        f"page: \"free\" — slide expanded to {self._slide_w / 914400:.1f}\" × "
                        f"{self._slide_h / 914400:.1f}\" to keep text readable (content "
                        f"would have shrunk to {slide_expanded_from:.0%} on the 16:9 slide).",
                    )
                )

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        if not network:
            self._add_title(slide, diagram.title)
        provider_by_group = self._group_providers(diagram)

        # Network connectors are drawn before nodes so routes sit behind icons/cards.
        self._draw_containers(slide, diagram, layout, tf, provider_by_group)
        if network:
            node_rects = self._node_rects(layout, tf)
            self._draw_network_edges(slide, diagram, node_rects)
            self._draw_network_nodes(slide, diagram, node_rects)
            add_connection_legend(slide, diagram, self.theme)
        else:
            # Flow routes sit behind the cards, matching the network renderer and
            # keeping arrow shafts from visually cutting through node artwork.
            node_rects = self._node_rects(layout, tf)
            edge_binds = self._draw_edges(slide, diagram, node_rects)
            node_shapes = self._draw_nodes(slide, diagram, node_rects, scale=tf.scale)
            # Bind connectors to their node shapes now that both exist, so the
            # arrows stay attached when a node is moved in PowerPoint.
            self._bind_edges(edge_binds, node_shapes)
            add_legends(slide, diagram, self.theme, slide_w=self._slide_w, slide_h=self._slide_h)

        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out_path))
        return Path(out_path)

    # -- title -------------------------------------------------------------
    def _add_title(self, slide, title: str) -> None:
        title_size = self.theme["slide"]["title_font_pt"]
        box = slide.shapes.add_textbox(
            self.theme.margin,
            Emu(Pt(10)),
            self.theme.slide_w - 2 * self.theme.margin,
            Emu(Pt(title_size + 12)),
        )
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        run = p.runs[0]
        run.font.size = Pt(title_size)
        run.font.bold = True
        run.font.name = self.theme.font
        run.font.color.rgb = Theme.rgb(self.theme["node"]["label_color"])

    # -- containers --------------------------------------------------------
    def _draw_containers(self, slide, diagram, layout, tf, provider_by_group) -> None:
        # Outer (shallower) groups first so nested ones stack on top.
        depth = {g.id: self._depth(g.id, diagram) for g in diagram.groups}
        clusters = sorted(layout.clusters, key=lambda c: depth.get(c.id, 0))
        group_by_id = {g.id: g for g in diagram.groups}
        gcfg = self.theme["group"]
        for c in clusters:
            if not c.label:
                continue
            rect = tf.bbox_rect(*c.bb)
            group = group_by_id.get(c.id)
            # Draw the spec's label, not dot's (build_dot pre-wraps the label
            # with \n escapes so dot reserves title headroom).
            label = group.label if group is not None else c.label
            if diagram.kind == DiagramKind.network and group is not None:
                self._draw_network_container(slide, group, label, rect)
                continue
            prov = provider_by_group.get(c.id, diagram.provider_default.value)
            pal = self.theme.provider(prov)
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rect.left, rect.top, rect.width, rect.height)
            shape.fill.solid()
            
            fill_hex = group.fill_color.lstrip("#") if group and group.fill_color else pal["tint"]
            shape.fill.fore_color.rgb = Theme.rgb(fill_hex)
            
            line_hex = group.border_color.lstrip("#") if group and group.border_color else pal["line"]
            shape.line.color.rgb = Theme.rgb(line_hex)
            shape.line.width = Pt(gcfg["line_pt"])
            
            if group and group.style:
                dash_map = {"dashed": "dash", "dotted": "sysDot", "solid": "solid"}
                dash_val = dash_map.get(group.style)
                if dash_val and dash_val != "solid":
                    self._set_line_dash(shape, dash_val)
                    
            self._remove_effects(shape)
            self._set_corner_radius(shape, 0.08)
            # Container title (top-left, inside). The box hugs its wrapped text:
            # estimate lines at this width and size the height to match, instead
            # of a fixed one-line box the text visibly overflows when it wraps.
            # The font scales down with the layout (dot reserved headroom in
            # graph space, so it shrank with the scale — a fixed-size title
            # would overflow that band straight into the first node).
            title_pt = max(9.0, gcfg["title_font_pt"] * min(tf.scale, 1.0))
            avail_w = rect.width - Emu(Pt(12))
            n_lines = _wrapped_line_count(label, avail_w, title_pt, bold=True)
            title_h = Emu(Pt(int(n_lines * title_pt * 1.35) + 4))
            title = slide.shapes.add_textbox(rect.left + Emu(Pt(6)), rect.top + Emu(Pt(3)), avail_w, title_h)
            ttf = title.text_frame
            ttf.word_wrap = True
            ttf.auto_size = MSO_AUTO_SIZE.NONE
            ttf.margin_top = ttf.margin_bottom = 0
            para = ttf.paragraphs[0]
            para.text = label
            r = para.runs[0]
            r.font.size = Pt(title_pt)
            r.font.bold = True
            r.font.name = self.theme.font
            title_hex = group.border_color.lstrip("#") if group and group.border_color else pal["line"]
            r.font.color.rgb = Theme.rgb(title_hex)

    def _draw_network_container(self, slide, group, label: str, rect: Rect) -> None:
        ncfg = self.theme["network"]
        geometry = MSO_SHAPE.ROUNDED_RECTANGLE if group.kind == GroupKind.network else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(geometry, rect.left, rect.top, rect.width, rect.height)
        self._remove_effects(shape)
        if group.kind in (GroupKind.cloud, GroupKind.network):
            shape.fill.background()
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = Theme.rgb(ncfg["panel_fill"])
        line_color = ncfg["cloud_line"] if group.kind == GroupKind.cloud else (
            ncfg["vnet_line"] if group.kind == GroupKind.network else ncfg["panel_line"]
        )
        shape.line.color.rgb = Theme.rgb(line_color)
        shape.line.width = Pt(1.5 if group.kind in (GroupKind.cloud, GroupKind.network) else 1.0)
        if group.kind == GroupKind.network:
            self._set_line_dash(shape, "dash")
            self._set_corner_radius(shape, 0.04)

        title_h = Emu(Pt(27 if group.kind in (GroupKind.subnet, GroupKind.cluster) else 16))
        title = slide.shapes.add_textbox(
            rect.left + Emu(Pt(5)), rect.top + Emu(Pt(2)), rect.width - Emu(Pt(10)), title_h
        )
        title.text_frame.word_wrap = group.kind in (GroupKind.subnet, GroupKind.cluster)
        title.text_frame.margin_top = title.text_frame.margin_bottom = 0
        title.text_frame.margin_left = title.text_frame.margin_right = 0
        p = title.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if group.kind in (GroupKind.subnet, GroupKind.cluster) else PP_ALIGN.LEFT
        p.text = label
        run = p.runs[0]
        run.font.size = Pt(9.0 if group.kind in (GroupKind.subnet, GroupKind.cluster) else ncfg["title_font_pt"])
        run.font.bold = group.kind in (GroupKind.cloud, GroupKind.network)
        run.font.name = self.theme.font
        run.font.color.rgb = Theme.rgb(ncfg["vnet_line"] if group.kind == GroupKind.network else ncfg["panel_line"])

    # -- nodes -------------------------------------------------------------
    def _draw_nodes(self, slide, diagram, rects: dict[str, Rect], scale: float = 1.0) -> dict[str, object]:
        ncfg = self.theme["node"]
        icon_by_node = {n.id: n for n in diagram.nodes}
        shapes: dict[str, object] = {}
        for nid, rect in rects.items():
            node = icon_by_node[nid]
            grp = slide.shapes.add_group_shape()
            self._place_icon(grp, node, rect)
            self._place_label(grp, node.label, rect, scale=scale)
            shapes[nid] = grp
        return shapes

    def _node_rects(self, layout, tf) -> dict[str, Rect]:
        return {
            nid: tf.node_rect(nl.cx, nl.cy, nl.w, nl.h)
            for nid, nl in layout.nodes.items()
        }

    def _draw_network_nodes(self, slide, diagram: Diagram, rects: dict[str, Rect]) -> None:
        nodes = {n.id: n for n in diagram.nodes}
        cfg = self.theme["network"]
        for nid, rect in rects.items():
            node = nodes[nid]
            role = node.role
            if role == NodeRole.workload:
                card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rect.left, rect.top, rect.width, rect.height)
                card.fill.solid()
                card.fill.fore_color.rgb = Theme.rgb(cfg["workload_fill"])
                card.line.color.rgb = Theme.rgb(cfg["workload_line"])
                card.line.width = Pt(0.75)
                self._remove_effects(card)
                icon_w = int(rect.width * 0.25)
                self._place_icon_fitted(
                    slide, node,
                    rect.left + Emu(Pt(5)), rect.top + Emu(Pt(5)),
                    icon_w, rect.height - Emu(Pt(10)),
                )
                self._network_text(
                    slide, node.label,
                    rect.left + icon_w, rect.top,
                    rect.width - icon_w, rect.height,
                    color=cfg["workload_text"], bold=True, alignment=PP_ALIGN.CENTER,
                )
                continue

            if role == NodeRole.dns:
                icon_w = min(int(rect.height), Emu(Pt(30)))
                self._place_icon_fitted(slide, node, rect.left, rect.top, icon_w, rect.height)
                self._network_text(
                    slide, node.label,
                    rect.left + icon_w + Emu(Pt(4)), rect.top,
                    rect.width - icon_w - Emu(Pt(4)), rect.height,
                    alignment=PP_ALIGN.LEFT,
                )
                continue

            if role == NodeRole.network_anchor:
                self._place_icon_fitted(slide, node, rect.left, rect.top, rect.width, rect.height)
                continue

            icon_h = int(rect.height * 0.56)
            icon_w = min(int(rect.width * 0.58), icon_h)
            self._place_icon_fitted(
                slide, node,
                rect.left + (rect.width - icon_w) // 2, rect.top,
                icon_w, icon_h,
            )
            # Give the label the icon-freed band plus a little of the inter-node
            # gap below, so multi-word labels ("Azure Front Door") don't clip.
            # Keep the horizontal widening small so labels don't bleed into the
            # neighbouring node in tightly-packed sections.
            self._network_text(
                slide, node.label,
                rect.left - Emu(Pt(3)), rect.top + icon_h,
                rect.width + Emu(Pt(6)), rect.height - icon_h + Emu(Pt(12)),
                alignment=PP_ALIGN.CENTER,
            )

    def _recolor_hex(self, entry) -> str | None:
        """Tint color for a recolorable monochrome logo: its official brand color
        when known, else the provider accent. Brand icons are never recolored."""
        if not entry.recolorable:
            return None
        brand = (entry.brand_hex or "").lstrip("#")
        pal = self.theme.provider(entry.provider)
        return brand or pal["hex"]

    def _place_icon_fitted(self, slide, node, left: int, top: int, width: int, height: int) -> None:
        entry = self.resolver.get(node.icon_key)
        recolor = self._recolor_hex(entry)
        pic_path = emf.to_picture(
            entry.path,
            recolor_hex=(f"#{recolor}" if recolor else None),
            prefer_vector=False,
        )
        aspect = emf.svg_aspect(entry.path)
        if width <= 0 or height <= 0:
            return
        if width / height >= aspect:
            h = height
            w = int(h * aspect)
        else:
            w = width
            h = int(w / aspect)
        slide.shapes.add_picture(
            str(pic_path), left + (width - w) // 2, top + (height - h) // 2, width=w, height=h
        )

    def _network_text(
        self, slide, text: str, left: int, top: int, width: int, height: int,
        *, color: str | None = None, bold: bool = False, alignment=PP_ALIGN.CENTER,
    ) -> None:
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(Pt(1))
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = alignment
        p.text = text
        r = p.runs[0]
        r.font.size = Pt(self.theme["network"]["label_font_pt"])
        r.font.bold = bold
        r.font.name = self.theme.font
        r.font.color.rgb = Theme.rgb(color or self.theme["node"]["label_color"])

    def _icon_slot(self, rect: Rect) -> int:
        """Icon-square size for a card. The configured inch value is the target,
        but it's capped at ~55% of the card height so a card shrunk by the
        uniform layout scale keeps room below the icon for the label — otherwise
        the label's top lands past the card bottom and hangs into whatever
        container sits underneath (see :meth:`_place_label`).

        Layered cards are renderer-sized (often smaller), so the icon is a pure
        fraction of card height — it scales down with the card instead of clamping
        to a fixed inch target that would crowd the label."""
        if self._layered:
            return int(rect.height * 0.46)
        target = Emu(Inches_val(self.theme["node"]["icon_slot_in"]))
        return min(target, int(rect.height * 0.55))

    def _place_icon(self, container, node, rect: Rect) -> None:
        entry = self.resolver.get(node.icon_key)
        recolor = self._recolor_hex(entry)
        # PNG pictures render consistently in PowerPoint, LibreOffice, and slide
        # preview pipelines; SVG source artwork is still preserved in the registry.
        pic_path = emf.to_picture(
            entry.path,
            recolor_hex=(f"#{recolor}" if recolor else None),
            prefer_vector=False,
        )
        aspect = emf.svg_aspect(entry.path)

        slot = self._icon_slot(rect)
        # Fit within the slot preserving aspect ratio (never distort brand icons).
        if aspect >= 1:
            w = slot
            h = int(slot / aspect)
        else:
            h = slot
            w = int(slot * aspect)
        icon_left = rect.left + (rect.width - w) // 2
        icon_top = rect.top + (int(rect.height * 0.07) if self._layered else Emu(Pt(8)))
        container.shapes.add_picture(str(pic_path), icon_left, icon_top, width=w, height=h)

    def _place_label(self, container, text: str, rect: Rect, scale: float = 1.0) -> None:
        ncfg = self.theme["node"]
        if self._layered:
            # Renderer-sized cards: label hugs the card width and its font scales
            # with card height (clamped to a readable floor) so smaller cards keep
            # legible, unbroken labels instead of clamping to fixed inches/points.
            pad = int(rect.width * 0.04)
            max_w = rect.width - 2 * pad
            left = rect.left + pad
            top = rect.top + self._icon_slot(rect) + int(rect.height * 0.05)
            height = max(rect.top + rect.height - top - int(rect.height * 0.04), Emu(Pt(10)))
            font_pt = _clamp_pt(rect.height / 914400 * 14.0, 7.5, ncfg["label_font_pt"])
        else:
            max_w = min(rect.width, Emu(Inches_val(ncfg["label_max_width_in"])))
            left = rect.left + (rect.width - max_w) // 2
            # Label sits below the icon slot. Using the rect-relative slot (not the
            # raw inch value) keeps `top` inside the card even when the layout scale
            # has shrunk it — so the box below never spills into the next container.
            top = rect.top + self._icon_slot(rect) + Emu(Pt(6))
            # Fixed label region occupying the lower part of the card. word_wrap keeps
            # long (linter-capped) labels wrapping to <=2 lines *inside* the card; we
            # deliberately avoid SHAPE_TO_FIT_TEXT, which would grow the box past the
            # card border and collide with the container.
            height = max(rect.top + rect.height - top - Emu(Pt(3)), Emu(Pt(12)))
            # The card was laid out in graph space and shrank with the uniform
            # scale; a fixed-size label overflows it and breaks mid-word. Scale
            # the font down too (readability floor 8pt — the layout report
            # already flags anything this cramped).
            font_pt = _clamp_pt(ncfg["label_font_pt"] * min(scale, 1.0), 8.0, ncfg["label_font_pt"])
        # HUG the text (§7): the box is only as wide as its content needs, up to
        # max_w (beyond which it wraps). A full-width box "works" visually but is
        # what the user grabs in PowerPoint — an oversized box that misrepresents
        # the label and collides with neighbours when moved.
        needed_w = _text_width_emu(text, font_pt) + Emu(Pt(6))
        box_w = min(max_w, max(needed_w, Emu(Pt(24))))
        left = left + (max_w - box_w) // 2
        # Height hugs the wrapped line count too, still clamped inside the card.
        n_lines = _wrapped_line_count(text, box_w, font_pt)
        height = min(height, max(Emu(Pt(int(n_lines * font_pt * 1.3) + 3)), Emu(Pt(12))))
        box = container.shapes.add_textbox(left, top, box_w, height)
        tframe = box.text_frame
        tframe.word_wrap = True
        tframe.auto_size = MSO_AUTO_SIZE.NONE
        tframe.vertical_anchor = MSO_ANCHOR.TOP
        tframe.margin_top = 0
        tframe.margin_bottom = 0
        tframe.margin_left = Emu(Pt(1))
        tframe.margin_right = Emu(Pt(1))
        p = tframe.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = text
        r = p.runs[0]
        r.font.size = Pt(font_pt)
        r.font.name = self.theme.font
        r.font.color.rgb = Theme.rgb(ncfg["label_color"])

    # -- edges -------------------------------------------------------------
    def _draw_edges(self, slide, diagram, node_rects: dict[str, Rect]) -> list["_EdgeBinding"]:
        """Draw connectors and return the data needed to bind each one to its
        node shapes (deferred to _bind_edges, since the node shapes are drawn
        after the edges so the routes sit behind the cards)."""
        ecfg = self.theme["edge"]
        placed_labels: list[Rect] = []
        binds: list[_EdgeBinding] = []
        for e in diagram.edges:
            a = node_rects.get(e.src)
            b = node_rects.get(e.dst)
            if a is None or b is None:
                continue
            # Native PowerPoint elbow (right-angle) connector — editable, and it
            # auto-routes cleanly instead of cutting diagonally across the slide.
            # Anchors are chosen purely from geometry, so this works for any graph.
            src_side, dst_side = _anchor_sides(a, b)
            (x1, y1), (x2, y2) = _side_anchor(a, src_side), _side_anchor(b, dst_side)
            connector_type = MSO_CONNECTOR.ELBOW if (x1 != x2 and y1 != y2) else MSO_CONNECTOR.STRAIGHT
            conn = slide.shapes.add_connector(connector_type, x1, y1, x2, y2)
            semantic = self.theme["connections"].get(e.kind.value, {})
            conn.line.color.rgb = Theme.rgb(semantic.get("color", ecfg["color"]))
            conn.line.width = Pt(ecfg["line_pt"])
            self._dash_and_arrow(conn, e.style.value)
            self._remove_effects(conn)
            binds.append(
                _EdgeBinding(
                    conn=conn,
                    src=e.src,
                    dst=e.dst,
                    src_idx=_SIDE_TO_CXN_IDX[src_side],
                    dst_idx=_SIDE_TO_CXN_IDX[dst_side],
                    coords=((x1, y1), (x2, y2)),
                )
            )
            if e.label:
                cx, cy = self._edge_label_spot(
                    e.label, (x1, y1), (x2, y2), node_rects, placed_labels
                )
                placed_labels.append(self._edge_label_rect(e.label, cx, cy))
                self._edge_label(slide, e.label, cx, cy)
        return binds

    def _bind_edges(self, binds: list["_EdgeBinding"], node_shapes: dict[str, object]) -> None:
        """Attach each connector's ends to its source/destination node shapes so
        moving a node in PowerPoint drags its arrows along (DESIGN_GUIDELINES §6).

        begin_connect/end_connect snap the endpoint to the node group's *content*
        bounding box (icon+label, narrower than the layout card), which would
        shift every rendered arrow inward. We restore the layout-computed anchors
        afterwards: the static render (qlmanage PNG) keeps its routing, while
        PowerPoint re-derives the endpoints from the stCxn/endCxn binding on open."""
        for bind in binds:
            src_shape = node_shapes.get(bind.src)
            dst_shape = node_shapes.get(bind.dst)
            if src_shape is None or dst_shape is None:
                continue
            try:
                bind.conn.begin_connect(src_shape, bind.src_idx)
                bind.conn.end_connect(dst_shape, bind.dst_idx)
            except Exception:  # noqa: BLE001 - binding is best-effort, geometry already drawn
                continue
            (x1, y1), (x2, y2) = bind.coords
            bind.conn.begin_x, bind.conn.begin_y = x1, y1
            bind.conn.end_x, bind.conn.end_y = x2, y2

    def _edge_label_rect(self, text: str, cx: int, cy: int) -> Rect:
        font_pt = self.theme["edge"]["label_font_pt"]
        w = _text_width_emu(text, font_pt) + Emu(Pt(6))
        h = Emu(Pt(font_pt + 5))
        return Rect(left=cx - w // 2, top=cy - h // 2, width=w, height=h)

    def _edge_label_spot(
        self,
        text: str,
        p1: tuple[int, int],
        p2: tuple[int, int],
        node_rects: dict[str, Rect],
        placed_labels: list[Rect],
    ) -> tuple[int, int]:
        """Pick a label spot ON the connector's route that collides with neither
        node cards nor already-placed edge labels. PowerPoint's elbow runs
        half-way along the dominant axis, so candidates walk that Z-path; the
        endpoint midpoint is only the fallback."""
        (x1, y1), (x2, y2) = p1, p2
        mx, my = (x1 + x2) // 2, (y1 + y2) // 2
        # A connector that runs straight through a node stack has NO clear
        # on-path spot — the lateral nudges sit the label beside the line,
        # which beats sitting it on a node.
        nudge = Inches_val(0.42)
        candidates = [
            (mx, my),                       # middle of the elbow's cross segment
            ((x1 + mx) // 2, y1),           # horizontal-dominant: leg leaving the source
            ((mx + x2) // 2, y2),           # horizontal-dominant: leg entering the target
            (x1, (y1 + my) // 2),           # vertical-dominant: leg leaving the source
            (x2, (my + y2) // 2),           # vertical-dominant: leg entering the target
            (mx + nudge, my), (mx - nudge, my),
            (x1 + nudge, (y1 + my) // 2), (x1 - nudge, (y1 + my) // 2),
        ]
        def is_free(rect: Rect) -> bool:
            # Inflate the probe: the width estimate is for the theme font, and
            # preview rasterizers substitute wider ones — a near-miss on paper
            # reads as a collision on screen.
            probe = Rect(
                left=rect.left - rect.width // 8,
                top=rect.top - rect.height // 8,
                width=rect.width + rect.width // 4,
                height=rect.height + rect.height // 4,
            )
            return not any(_rects_intersect(probe, r) for r in node_rects.values()) and not any(
                _rects_intersect(probe, r) for r in placed_labels
            )

        for cx, cy in candidates:
            if is_free(self._edge_label_rect(text, cx, cy)):
                return cx, cy
        # Nothing on or near the path is clear (typical for a connector running
        # straight through a node stack): shove the label sideways just past
        # whatever card the midpoint hits.
        mid_rect = self._edge_label_rect(text, mx, my)
        hit = next((r for r in node_rects.values() if _rects_intersect(mid_rect, r)), None)
        if hit is not None:
            gap = Emu(Pt(4))
            for cx in (
                hit.left + hit.width + mid_rect.width // 2 + gap,
                hit.left - mid_rect.width // 2 - gap,
            ):
                if is_free(self._edge_label_rect(text, cx, my)):
                    return cx, my
        return mx, my

    def _draw_network_edges(self, slide, diagram, node_rects: dict[str, Rect]) -> None:
        nodes = {node.id: node for node in diagram.nodes}
        kind_indexes: dict[str, int] = {}
        for edge in diagram.edges:
            a = node_rects.get(edge.src)
            b = node_rects.get(edge.dst)
            if a is None or b is None:
                continue
            index = kind_indexes.get(edge.kind.value, 0)
            kind_indexes[edge.kind.value] = index + 1
            points = self._network_route(
                edge.kind.value,
                nodes[edge.src].role,
                nodes[edge.dst].role,
                a,
                b,
                index,
                Inches_val(self.theme["network"]["slide_width_in"]),
            )
            cfg = self.theme["connections"][edge.kind.value]
            conn = self._add_network_polyline(slide, points, cfg["color"], edge.style.value)
            if edge.label:
                mid = points[len(points) // 2]
                self._edge_label(slide, edge.label, *mid)

    def _network_route(
        self, kind: str, src_role: NodeRole, dst_role: NodeRole,
        a: Rect, b: Rect, index: int, slide_width: int,
    ) -> list[tuple[int, int]]:
        """Return a clean Manhattan route using semantic connection lanes."""
        gap = Emu(Pt(10))
        if kind in ("https", "traffic"):
            if src_role == NodeRole.external:
                start, end = _side_anchor(a, "right"), _side_anchor(b, "left")
                lane_x = min(end[0] - gap, start[0] + Emu(Pt(18)))
                return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])
            dx, dy = b.cx - a.cx, b.cy - a.cy
            if abs(dy) > max(a.height, b.height):
                if dy > 0:
                    start, end = _side_anchor(a, "bottom"), _side_anchor(b, "top")
                else:
                    start, end = _side_anchor(a, "top"), _side_anchor(b, "bottom")
                lane_y = (start[1] + end[1]) // 2 + Emu(Pt((index % 3 - 1) * 4))
                return _dedupe_points([start, (start[0], lane_y), (end[0], lane_y), end])
            if dx >= 0:
                start, end = _side_anchor(a, "right"), _side_anchor(b, "left")
            else:
                start, end = _side_anchor(a, "left"), _side_anchor(b, "right")
            if abs(start[1] - end[1]) <= Emu(Pt(8)):
                return [start, end]
            lane_x = (start[0] + end[0]) // 2
            return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])

        if kind == "virtual_network_link":
            start, end = _side_anchor(a, "top"), _side_anchor(b, "left")
            lane_x = Emu(Pt(276 + index * 34))
            return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])

        if kind == "dns":
            if dst_role == NodeRole.dns:
                start = _side_anchor(a, "top")
                end = _side_anchor(b, "bottom" if b.cy < a.cy else "left")
                lane_y = end[1] + Emu(Pt(10 + (index % 3) * 6))
                return _dedupe_points([start, (start[0], lane_y), (end[0], lane_y), end])
            start, end = _side_anchor(a, "right"), _side_anchor(b, "left")
            lane_x = (start[0] + end[0]) // 2
            return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])

        if kind == "private_connection":
            start = _side_anchor(a, "bottom")
            end = _side_anchor(b, "top")
            if src_role == NodeRole.private_endpoint:
                return [start, end]
            lane_y = end[1] - Emu(Pt(13 + (index % 3) * 6))
            return _dedupe_points([start, (start[0], lane_y), (end[0], lane_y), end])

        if kind == "diagnostics":
            start = _side_anchor(a, "right")
            end = _side_anchor(b, "right")
            lane_x = slide_width - Emu(Pt(37 + index * 8))
            return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])

        start, end = _border_points(a, b)
        lane_x = (start[0] + end[0]) // 2
        return _dedupe_points([start, (lane_x, start[1]), (lane_x, end[1]), end])

    def _add_network_polyline(self, slide, points, color: str, style: str):
        if len(points) == 2:
            shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *points[0], *points[1])
        else:
            builder = slide.shapes.build_freeform(points[0][0], points[0][1], scale=1.0)
            builder.add_line_segments(points[1:], close=False)
            shape = builder.convert_to_shape()
            shape.fill.background()
        shape.line.color.rgb = Theme.rgb(color)
        shape.line.width = Pt(self.theme["network"]["connection_line_pt"])
        self._dash_and_arrow(shape, style)
        self._remove_effects(shape)
        return shape

    def _edge_label(self, slide, text: str, cx: int, cy: int) -> None:
        ecfg = self.theme["edge"]
        font_pt = ecfg["label_font_pt"]
        # Size the box to hold the text on ONE line. python-pptx text boxes carry
        # ~0.1" internal margins by default; left unzeroed they eat the usable
        # width and force short labels (e.g. "15") to wrap into stacked digits.
        pad = Emu(Pt(3))
        text_w = Emu(Pt(len(text) * font_pt * 0.62))
        w = text_w + 2 * pad
        h = Emu(Pt(font_pt + 5))
        left = cx - w // 2
        top = cy - h // 2
        # Keep the label fully on-slide.
        left = max(0, min(left, self._slide_w - w))
        top = max(0, min(top, self._slide_h - h))
        box = slide.shapes.add_textbox(left, top, w, h)
        tframe = box.text_frame
        tframe.word_wrap = False
        tframe.auto_size = MSO_AUTO_SIZE.NONE
        tframe.margin_left = tframe.margin_right = pad
        tframe.margin_top = tframe.margin_bottom = 0
        p = tframe.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = text
        r = p.runs[0]
        r.font.size = Pt(font_pt)
        r.font.name = self.theme.font
        r.font.color.rgb = Theme.rgb(ecfg["label_color"])
        # Small white backing so the label reads over lines.
        box.fill.solid()
        box.fill.fore_color.rgb = Theme.rgb("FFFFFF")
        box.line.fill.background()
        self._remove_effects(box)

    # -- xml helpers -------------------------------------------------------
    def _dash_and_arrow(self, conn, style: str) -> None:
        ln = conn.line._get_or_add_ln()
        # Arrowhead at the destination end.
        tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
        ln.append(tail)
        if style in ("dashed", "dotted"):
            d = ln.makeelement(qn("a:prstDash"), {"val": "dash" if style == "dashed" else "sysDot"})
            ln.append(d)

    def _set_corner_radius(self, shape, frac: float) -> None:
        try:
            shape.adjustments[0] = frac
        except (IndexError, KeyError):
            pass

    def _set_line_dash(self, shape, dash: str) -> None:
        ln = shape.line._get_or_add_ln()
        existing = ln.find(qn("a:prstDash"))
        if existing is not None:
            ln.remove(existing)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))

    def _remove_effects(self, shape) -> None:
        """Force a flat treatment even when the Office theme defines shadows."""
        style = shape._element.find(qn("p:style"))
        if style is not None:
            shape._element.remove(style)
        sp_pr = shape._element.spPr
        for tag in ("a:effectLst", "a:effectDag"):
            existing = sp_pr.find(qn(tag))
            if existing is not None:
                sp_pr.remove(existing)
        sp_pr.append(OxmlElement("a:effectLst"))

    def _group_providers(self, diagram: Diagram) -> dict[str, str]:
        prov_of_node = {n.id: self.resolver.get(n.icon_key).provider if self.resolver.exists(n.icon_key) else "generic" for n in diagram.nodes}
        members: dict[str, list[str]] = {}
        for n in diagram.nodes:
            if n.group:
                members.setdefault(n.group, []).append(n.id)
        out: dict[str, str] = {}
        for g in diagram.groups:
            if g.provider is not None:
                out[g.id] = g.provider.value
                continue
            provs = [prov_of_node[nid] for nid in members.get(g.id, [])]
            provs = [p for p in provs if p not in ("generic", "tech")] or provs
            out[g.id] = max(set(provs), key=provs.count) if provs else diagram.provider_default.value
        return out

    def _depth(self, gid: str, diagram: Diagram) -> int:
        parent = {g.id: g.parent for g in diagram.groups}
        d, cur = 0, parent.get(gid)
        while cur is not None:
            d += 1
            cur = parent.get(cur)
        return d


def Inches_val(inches: float) -> int:
    from pptx.util import Inches

    return int(Inches(inches))


def _text_width_emu(text: str, font_pt: float, bold: bool = False) -> int:
    """Generous single-line width estimate for the default theme font. Errs
    wide on purpose: an over-wide box still hugs far better than a full-width
    one, while an under-estimate causes surprise wrapping."""
    per_char_pt = font_pt * (0.68 if bold else 0.62)
    return int(Emu(Pt(len(text) * per_char_pt)))


def _wrapped_line_count(text: str, width_emu: int, font_pt: float, bold: bool = False) -> int:
    """How many lines `text` needs when word-wrapped into `width_emu`."""
    if width_emu <= 0:
        return 1
    lines, cur = 1, ""
    for word in text.split():
        cand = f"{cur} {word}".strip()
        if cur and _text_width_emu(cand, font_pt, bold) > width_emu:
            lines += 1
            cur = word
        else:
            cur = cand
    return lines


def _rects_intersect(a: Rect, b: Rect) -> bool:
    return not (
        a.left + a.width <= b.left
        or b.left + b.width <= a.left
        or a.top + a.height <= b.top
        or b.top + b.height <= a.top
    )


def _clamp_pt(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _border_points(a: Rect, b: Rect) -> tuple[tuple[int, int], tuple[int, int]]:
    """Endpoints on the borders of a and b along the line joining their centers."""
    return _edge_point(a, b.cx, b.cy), _edge_point(b, a.cx, a.cy)


def _edge_point(rect: Rect, tx: int, ty: int) -> tuple[int, int]:
    """Intersection of the segment (rect center -> target) with rect's border."""
    cx, cy = rect.cx, rect.cy
    dx, dy = tx - cx, ty - cy
    if dx == 0 and dy == 0:
        return cx, cy
    hw, hh = rect.width / 2, rect.height / 2
    scale = min(hw / abs(dx) if dx else float("inf"), hh / abs(dy) if dy else float("inf"))
    return int(cx + dx * scale), int(cy + dy * scale)


# Connection-site index for each rectangle side. python-pptx numbers a shape's
# connection points counter-clockwise from the top-centre, so top=0, left=1,
# bottom=2, right=3 — used to bind connector ends to node shapes (see
# Renderer._bind_edges).
_SIDE_TO_CXN_IDX = {"top": 0, "left": 1, "bottom": 2, "right": 3}


def _anchor_sides(a: Rect, b: Rect) -> tuple[str, str]:
    """Which side of the source and destination an elbow connector leaves/enters,
    chosen by dominant axis so the right-angle route faces the two nodes."""
    dx, dy = b.cx - a.cx, b.cy - a.cy
    # Prefer vertical (top/bottom) anchors when there is significant vertical separation,
    # preventing horizontal lines from cutting across vertical columns.
    if abs(dx) >= 1.8 * abs(dy):  # horizontally dominant
        return ("right", "left") if dx >= 0 else ("left", "right")
    # vertically dominant -> leave/enter top|bottom sides
    return ("bottom", "top") if dy >= 0 else ("top", "bottom")


def _elbow_anchors(a: Rect, b: Rect) -> tuple[tuple[int, int], tuple[int, int]]:
    """Exit/entry points for an elbow connector, chosen by dominant axis so the
    right-angle route leaves and enters the facing sides of each node."""
    sa, sb = _anchor_sides(a, b)
    return _side_anchor(a, sa), _side_anchor(b, sb)


def _side_anchor(rect: Rect, side: str) -> tuple[int, int]:
    if side == "left":
        return rect.left, rect.cy
    if side == "right":
        return rect.left + rect.width, rect.cy
    if side == "top":
        return rect.cx, rect.top
    if side == "bottom":
        return rect.cx, rect.top + rect.height
    raise ValueError(f"unknown rectangle side: {side}")


def _dedupe_points(points: list[tuple[int, int]]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for point in points:
        if not out or point != out[-1]:
            out.append(point)
    return out


def render_spec(
    diagram: Diagram,
    out_path: str | Path,
    engine: str = "auto",
    auto_direction: bool = False,
    expand_slide: bool | None = None,
    quality: str = "draft",
) -> Path:
    # expand_slide defaults to None so the renderer derives it from the spec's
    # `page` field (16:9 locks the slide, free lets it grow). Hardcoding True
    # here silently turned formal 16:9 decks into tall posters via the
    # `generate` path — see Renderer.render.
    return Renderer().render(
        diagram,
        out_path,
        engine=engine,
        auto_direction=auto_direction,
        expand_slide=expand_slide,
        quality=quality,
    )
