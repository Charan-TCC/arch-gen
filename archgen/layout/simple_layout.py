"""Pure-Python **group-aware** layered layout — the zero-dependency fallback when
Graphviz `dot` is not installed.

Produces the same :class:`~archgen.layout.graphviz_layout.Layout` structure and
coordinate convention (points, y-up, node ``pos`` = center, ``bb`` = x0,y0,x1,y1),
so geometry + renderer are engine-agnostic.

Key property: each group is laid out as a self-contained **block** (recursively),
then blocks are arranged along the flow like ordinary items. Because a block is a
solid rectangle placed on the grid with spacing, **container rectangles never
overlap** — the main thing a naive per-node layering gets wrong. Graphviz still
minimizes edge crossings better; install it (`brew install graphviz`) for the
tightest results.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from archgen.layout.graphviz_layout import (
    NODE_H_IN,
    NODE_W_IN,
    ClusterLayout,
    EdgeLayout,
    Layout,
    NodeLayout,
)
from archgen.spec.models import Diagram

_PT = 72.0
NODE_W = NODE_W_IN * _PT
NODE_H = NODE_H_IN * _PT
RANKSEP = 0.7 * _PT      # gap between layers (main axis)
NODESEP = 0.45 * _PT     # gap between items in a layer (cross axis)
PAD = 0.22 * _PT         # padding inside a container
TITLE_BAND = 24.0        # vertical space reserved at the top of a container for its title


@dataclass
class _Block:
    """Result of laying out one scope (the whole diagram, or one group)."""

    w: float
    h: float
    node_pos: dict[str, tuple[float, float]]  # node_id -> (x, y) top-left-origin, y-down, absolute within block
    boxes: list[tuple[str, float, float, float, float]] = field(default_factory=list)  # (gid, x0,y0,x1,y1)


def compute_layout(diagram: Diagram) -> Layout:
    horizontal = diagram.direction == "LR"

    members: dict[str | None, list[str]] = {}
    for n in diagram.nodes:
        members.setdefault(n.group, []).append(n.id)
    children: dict[str | None, list[str]] = {}
    for g in diagram.groups:
        children.setdefault(g.parent, []).append(g.id)
    label_of = {g.id: g.label for g in diagram.groups}

    # All node ids reachable under a group (for edge projection).
    def all_nodes(gid: str) -> list[str]:
        out = list(members.get(gid, []))
        for c in children.get(gid, []):
            out += all_nodes(c)
        return out

    succ_all: dict[str, list[str]] = {n.id: [] for n in diagram.nodes}
    for e in diagram.edges:
        if e.src in succ_all:
            succ_all[e.src].append(e.dst)

    def layout_scope(scope_gid: str | None) -> _Block:
        node_items = list(members.get(scope_gid, []))
        group_items = list(children.get(scope_gid, []))
        sub_blocks: dict[str, _Block] = {g: layout_scope(g) for g in group_items}

        # item i -> (kind, key, main_size, cross_size)
        items: list[tuple[str, str]] = [("node", nid) for nid in node_items] + [("group", g) for g in group_items]
        if not items:
            return _Block(NODE_W, NODE_H, {})

        def item_size(kind: str, key: str) -> tuple[float, float]:
            if kind == "node":
                w, h = NODE_W, NODE_H
            else:
                w, h = sub_blocks[key].w, sub_blocks[key].h
            return (w, h) if horizontal else (h, w)  # returns (main, cross)

        # Map every node to the index of the item that contains it (in this scope).
        item_of: dict[str, int] = {}
        for idx, (kind, key) in enumerate(items):
            if kind == "node":
                item_of[key] = idx
            else:
                for nid in all_nodes(key):
                    item_of[nid] = idx

        # Item-level adjacency (project node edges up to this scope's items).
        isucc: dict[int, set[int]] = {i: set() for i in range(len(items))}
        for src, dsts in succ_all.items():
            iu = item_of.get(src)
            if iu is None:
                continue
            for dst in dsts:
                iv = item_of.get(dst)
                if iv is not None and iu != iv:
                    isucc[iu].add(iv)

        layer = _longest_path(list(range(len(items))), isucc)
        by_layer: dict[int, list[int]] = {}
        for i in range(len(items)):
            by_layer.setdefault(layer[i], []).append(i)

        # Main-axis position per layer (advance by the layer's max main size).
        layer_main_start: dict[int, float] = {}
        cur = 0.0
        for lvl in sorted(by_layer):
            layer_main_start[lvl] = cur
            main_extent = max(item_size(*items[i])[0] for i in by_layer[lvl])
            cur += main_extent + RANKSEP
        # Cross-axis: total block cross size = widest layer stack.
        layer_cross_extent = {
            lvl: sum(item_size(*items[i])[1] for i in ids) + NODESEP * (len(ids) - 1)
            for lvl, ids in by_layer.items()
        }
        block_cross = max(layer_cross_extent.values())

        node_pos: dict[str, tuple[float, float]] = {}
        boxes: list[tuple[str, float, float, float, float]] = []

        for lvl, ids in by_layer.items():
            main_extent = max(item_size(*items[i])[0] for i in ids)
            main_start = layer_main_start[lvl]
            cross_cursor = (block_cross - layer_cross_extent[lvl]) / 2  # center the stack
            for i in ids:
                kind, key = items[i]
                m_size, c_size = item_size(kind, key)
                m_center = main_start + main_extent / 2
                c_center = cross_cursor + c_size / 2
                cross_cursor += c_size + NODESEP
                # Map (main,cross) -> (x,y) top-left origin, y-down.
                if horizontal:
                    x_center, y_center = m_center, c_center
                    w, h = (NODE_W, NODE_H) if kind == "node" else (sub_blocks[key].w, sub_blocks[key].h)
                else:
                    x_center, y_center = c_center, m_center
                    w, h = (NODE_W, NODE_H) if kind == "node" else (sub_blocks[key].w, sub_blocks[key].h)

                if kind == "node":
                    node_pos[key] = (x_center, y_center)
                else:
                    sub = sub_blocks[key]
                    ox, oy = x_center - w / 2, y_center - h / 2
                    for nid, (nx, ny) in sub.node_pos.items():
                        node_pos[nid] = (ox + nx, oy + ny)
                    boxes.append((key, ox, oy, ox + w, oy + h))
                    for (cg, bx0, by0, bx1, by1) in sub.boxes:
                        boxes.append((cg, ox + bx0, oy + by0, ox + bx1, oy + by1))

        # Block extent from placed items.
        w_total = cur - RANKSEP if horizontal else block_cross
        h_total = block_cross if horizontal else cur - RANKSEP

        # If this scope *is* a group, wrap with padding + title band and shift content.
        if scope_gid is not None:
            shift_x, shift_y = PAD, PAD + TITLE_BAND
            node_pos = {n: (x + shift_x, y + shift_y) for n, (x, y) in node_pos.items()}
            boxes = [(g, x0 + shift_x, y0 + shift_y, x1 + shift_x, y1 + shift_y) for (g, x0, y0, x1, y1) in boxes]
            w_total += 2 * PAD
            h_total += 2 * PAD + TITLE_BAND

        return _Block(w=w_total, h=h_total, node_pos=node_pos, boxes=boxes)

    root = layout_scope(None)

    # Flip y-down -> y-up (convention) and build the Layout.
    H = root.h
    nodes: dict[str, NodeLayout] = {}
    for nid, (x, y) in root.node_pos.items():
        nodes[nid] = NodeLayout(id=nid, cx=x, cy=H - y, w=NODE_W, h=NODE_H)
    clusters = [
        ClusterLayout(id=gid, label=label_of.get(gid, gid), bb=(x0, H - y1, x1, H - y0))
        for (gid, x0, y0, x1, y1) in root.boxes
    ]
    edges = [EdgeLayout(src=e.src, dst=e.dst, points=[]) for e in diagram.edges]
    graph_bb = (0.0, 0.0, root.w, root.h)
    return Layout(graph_bb=graph_bb, nodes=nodes, clusters=clusters, edges=edges)


def _longest_path(items: list[int], succ: dict[int, set[int]]) -> dict[int, int]:
    layer = {i: 0 for i in items}
    for _ in range(len(items) + 1):
        changed = False
        for i in items:
            for s in succ[i]:
                if layer[s] < layer[i] + 1:
                    layer[s] = layer[i] + 1
                    changed = True
        if not changed:
            break
    return layer
