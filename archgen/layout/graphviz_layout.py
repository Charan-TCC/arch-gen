"""Graphviz layout adapter.

Builds a DOT document from the spec (groups -> `subgraph cluster_*`), runs
`dot -Tjson`, and parses back node positions/sizes, cluster bounding boxes, and
edge spline points — all in Graphviz point units for :mod:`archgen.layout.geometry`.

We shell out to `dot` (rather than use pygraphviz) to avoid a C-extension build;
`-Tjson` gives a clean machine-readable layout.
"""

from __future__ import annotations

import json
import shutil
import subprocess
from dataclasses import dataclass, field

from archgen.spec.models import Diagram

# Node box size in inches, fed to dot as fixed size. The renderer re-derives
# pixel sizes from the returned layout, so this only sets relative proportions.
# Height leaves room for the icon (upper) + a two-line label (lower) inside one card.
NODE_W_IN = 1.5
NODE_H_IN = 1.35

# Default spacing (points/inches) for ordinary flow diagrams.
# Tight enough that the diagram reads as one composition (the old 1.2/2.0
# left acres of empty canvas and forced poster-sized slides); ortho routing
# still has room to run channels between cards.
NODESEP_IN = 0.55
RANKSEP_IN = 0.95
CLUSTER_MARGIN_PT = 24

# Layered mode uses much tighter spacing so tiers pack into clean bands instead
# of the airy default that makes tall columns overflow the slide.
LAYERED_NODESEP_IN = 0.3
LAYERED_RANKSEP_IN = 1.2
LAYERED_CLUSTER_MARGIN_PT = 12

# Card-size envelope the renderer picks within, for layered auto-sizing.
LAYERED_CARD_W_MIN, LAYERED_CARD_W_MAX = 1.35, 1.85
LAYERED_CARD_H_MIN, LAYERED_CARD_H_MAX = 0.62, 0.92

# Cluster title metrics fed to dot so it reserves REAL headroom for the title
# the renderer will draw (13pt bold, wrapping in narrow clusters). Kept in sync
# with theme group.title_font_pt; the char width is a generous bold estimate so
# dot reserves at least as many lines as PowerPoint will render — over-reserving
# is safe (airier container), under-reserving overlaps the first node.
CLUSTER_TITLE_PT = 13
CLUSTER_TITLE_CHAR_IN = 0.105
# The dot-side reserve scales down with the layout while the drawn title font
# clamps at a readable minimum — reserve at a larger fontsize so the headroom
# still clears the drawn title at the lowest readable scale (~0.5).
CLUSTER_TITLE_RESERVE_PT = 20


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def layered_card_size(diagram: Diagram, avail_w_in: float, avail_h_in: float) -> tuple[float, float]:
    """Pick a uniform card (w, h) in inches so the DENSEST tier fits the slide.

    The renderer owns sizing (the spec only declares ``layered``): we count nodes
    per tier and size cards so the busiest tier fits the available area at a
    legibility floor. Sparse diagrams land at the max (normal-looking) size.
    """
    by_group: dict[str | None, int] = {}
    for n in diagram.nodes:
        by_group[n.group] = by_group.get(n.group, 0) + 1
    tiers = [gid for gid in by_group if gid is not None] or list(by_group)
    n_tiers = max(len(tiers), 1)
    n_max = max(by_group.values(), default=1)

    ns, rs = LAYERED_NODESEP_IN, LAYERED_RANKSEP_IN
    cm = LAYERED_CLUSTER_MARGIN_PT / 72.0
    title = (CLUSTER_TITLE_PT * 1.6) / 72.0  # per-cluster title headroom band

    if diagram.direction == "LR":
        # Tiers are COLUMNS side by side (ranksep apart); nodes stack vertically
        # inside a tier (nodesep apart). Vertical budget is one tier's internals.
        v_over = (n_max - 1) * ns + 2 * cm + title
        h_over = (n_tiers - 1) * rs + 2 * cm * n_tiers + 0.4
        card_h = _clamp((avail_h_in - v_over) / n_max, LAYERED_CARD_H_MIN, LAYERED_CARD_H_MAX)
        card_w = _clamp((avail_w_in - h_over) / n_tiers, LAYERED_CARD_W_MIN, LAYERED_CARD_W_MAX)
    else:
        # TB: tiers are ROWS stacked (ranksep apart, each with its own margins
        # and title band); nodes lie side by side inside a row (nodesep apart).
        v_over = n_tiers * (2 * cm + title) + (n_tiers - 1) * rs
        h_over = (n_max - 1) * ns + 2 * cm + 0.4
        card_h = _clamp((avail_h_in - v_over) / n_tiers, LAYERED_CARD_H_MIN, LAYERED_CARD_H_MAX)
        card_w = _clamp((avail_w_in - h_over) / n_max, LAYERED_CARD_W_MIN, LAYERED_CARD_W_MAX)
    return card_w, card_h


class LayoutError(RuntimeError):
    pass


@dataclass
class NodeLayout:
    id: str
    cx: float  # center x, points
    cy: float  # center y, points
    w: float  # width, points
    h: float  # height, points


@dataclass
class ClusterLayout:
    id: str
    label: str
    bb: tuple[float, float, float, float]  # x0,y0,x1,y1 in points


@dataclass
class EdgeLayout:
    src: str
    dst: str
    points: list[tuple[float, float]] = field(default_factory=list)  # route polyline, points (start -> arrow tip)
    label_pos: tuple[float, float] | None = None  # graphviz-placed label centre, points


@dataclass
class Layout:
    graph_bb: tuple[float, float, float, float]
    nodes: dict[str, NodeLayout]
    clusters: list[ClusterLayout]
    edges: list[EdgeLayout]


def compute_layout(
    diagram: Diagram,
    dot_bin: str | None = None,
    *,
    node_w: float | None = None,
    node_h: float | None = None,
) -> Layout:
    dot = dot_bin or shutil.which("dot")
    if not dot:
        raise LayoutError(
            "Graphviz 'dot' not found on PATH. Install it (macOS: `brew install graphviz`)."
        )
    dot_src = build_dot(diagram, node_w=node_w, node_h=node_h)
    try:
        proc = subprocess.run(
            [dot, "-Tjson"], input=dot_src, capture_output=True, text=True, check=True
        )
    except subprocess.CalledProcessError as exc:  # pragma: no cover - surfaced to caller
        raise LayoutError(f"dot failed: {exc.stderr.strip()}") from exc
    return _parse_json(json.loads(proc.stdout))


def build_dot(diagram: Diagram, *, node_w: float | None = None, node_h: float | None = None) -> str:
    """Render the spec to a DOT document. Groups become `cluster_*` subgraphs.

    ``node_w``/``node_h`` override the default card proportions (the renderer sizes
    cards for layered mode). When ``diagram.layered`` is set, tiers are pinned to a
    single rank each (clean bands) and spacing/margins tighten.
    """
    layered = diagram.layered
    nw = node_w if node_w is not None else NODE_W_IN
    nh = node_h if node_h is not None else NODE_H_IN
    nodesep = LAYERED_NODESEP_IN if layered else NODESEP_IN
    ranksep = LAYERED_RANKSEP_IN if layered else RANKSEP_IN
    cluster_margin = LAYERED_CLUSTER_MARGIN_PT if layered else CLUSTER_MARGIN_PT

    lines: list[str] = ["digraph arch {"]
    lines.append(f'  rankdir={diagram.direction};')
    lines.append("  compound=true;")  # allow edges to/from clusters
    lines.append("  newrank=true;")
    # Orthogonal routing: graphviz computes right-angle routes that go AROUND
    # nodes; the renderer draws these instead of inventing its own elbows.
    # Ortho can't carry edge labels, so labels become xlabels, which dot
    # places next to the route without overlapping nodes.
    lines.append("  splines=ortho;")
    lines.append("  forcelabels=true;")
    lines.append('  node [shape=box, fixedsize=true, '
                 f'width={nw}, height={nh}];')
    lines.append(f"  nodesep={nodesep};")
    lines.append(f"  ranksep={ranksep};")

    # Index nodes by group so we can emit them inside cluster subgraphs.
    by_group: dict[str | None, list[str]] = {}
    for n in diagram.nodes:
        by_group.setdefault(n.group, []).append(n.id)

    # Build nesting: children per parent group.
    children: dict[str | None, list[str]] = {}
    for g in diagram.groups:
        children.setdefault(g.parent, []).append(g.id)
    group_by_id = {g.id: g for g in diagram.groups}

    def emit_group(gid: str, indent: str) -> list[str]:
        g = group_by_id[gid]
        out = [f'{indent}subgraph {_safe("cluster_" + gid)} {{']
        if g.label:
            # Pre-wrap the title at the narrowest width this cluster can have
            # (its widest member card + margins) so dot reserves one line of
            # headroom PER wrapped line — a wide title in a one-column tier
            # otherwise overlaps the first node when the renderer wraps it.
            min_w_in = nw + 2 * (cluster_margin / 72.0)
            wrapped = _wrap_title(g.label, min_w_in)
            out.append(f'{indent}  label="{_esc(wrapped)}";')
            out.append(f"{indent}  fontsize={CLUSTER_TITLE_RESERVE_PT};")
            out.append(f'{indent}  style="rounded";')
            out.append(f'{indent}  margin={cluster_margin};')
        else:
            out.append(f'{indent}  style="invis";')
            out.append(f'{indent}  margin={cluster_margin // 2};')
        for nid in by_group.get(gid, []):
            out.append(f'{indent}  {_safe(nid)};')
        for child in children.get(gid, []):
            out += emit_group(child, indent + "  ")
        out.append(f"{indent}}}")
        return out

    # Ungrouped nodes at top level.
    for nid in by_group.get(None, []):
        lines.append(f"  {_safe(nid)};")
    # Top-level groups (parent is None) recurse into their children.
    for gid in children.get(None, []):
        lines += emit_group(gid, "  ")

    style_map = {"solid": "solid", "dashed": "dashed", "dotted": "dotted"}
    for e in diagram.edges:
        attrs = [f'style={style_map[e.style.value]}']
        if e.label:
            attrs.append(f'xlabel="{_esc(e.label)}"')
        lines.append(f"  {_safe(e.src)} -> {_safe(e.dst)} [{', '.join(attrs)}];")

    # Layered mode: pin every tier (group) to one rank so it renders as a clean
    # parallel band instead of staggering by edge depth. newrank=true (set above)
    # lets these top-level rank=same constraints reach nodes inside clusters.
    if layered:
        for g in diagram.groups:
            members = by_group.get(g.id, [])
            if len(members) > 1:
                pinned = "; ".join(_safe(nid) for nid in members)
                lines.append(f"  {{rank=same; {pinned};}}")

    lines.append("}")
    return "\n".join(lines)


def _parse_json(doc: dict) -> Layout:
    graph_bb = _parse_bb(doc.get("bb", "0,0,100,100"))
    id_by_gvid: dict[str, str] = {}
    nodes: dict[str, NodeLayout] = {}
    # First pass: real nodes have a "pos"; clusters have a "bb".
    clusters: list[ClusterLayout] = []
    for obj in doc.get("objects", []):
        name = str(obj.get("name", ""))
        gvid = str(obj.get("_gvid", ""))
        if name.startswith("cluster_") and "bb" in obj:
            clusters.append(
                ClusterLayout(
                    id=name[len("cluster_"):],
                    label=obj.get("label", ""),
                    bb=_parse_bb(obj["bb"]),
                )
            )
        elif "pos" in obj:
            cx, cy = _parse_pos(obj["pos"])
            w = float(obj.get("width", NODE_W_IN)) * 72.0
            h = float(obj.get("height", NODE_H_IN)) * 72.0
            nodes[name] = NodeLayout(id=name, cx=cx, cy=cy, w=w, h=h)
            id_by_gvid[gvid] = name

    edges: list[EdgeLayout] = []
    for e in doc.get("edges", []):
        tail = id_by_gvid.get(str(e.get("tail")))
        head = id_by_gvid.get(str(e.get("head")))
        if tail is None or head is None:
            continue
        xlp = e.get("xlp") or e.get("lp")
        edges.append(EdgeLayout(src=tail, dst=head, points=_parse_spline(e.get("pos", "")),
                                label_pos=_parse_pos(xlp) if xlp else None))

    return Layout(graph_bb=graph_bb, nodes=nodes, clusters=clusters, edges=edges)


def _parse_bb(s: str) -> tuple[float, float, float, float]:
    parts = [float(x) for x in s.split(",")]
    return (parts[0], parts[1], parts[2], parts[3])


def _parse_pos(s: str) -> tuple[float, float]:
    x, y = s.split(",")[:2]
    return float(x), float(y)


def _parse_spline(s: str) -> list[tuple[float, float]]:
    """Graphviz edge ``pos`` -> polyline from tail to arrow tip.

    Format: ``[s,x,y ][e,x,y ]p1 p2 ...``. The ``e`` point is the arrow TIP and
    belongs at the END (it is listed first). For ortho routes the control points
    are collinear runs, so deduping + dropping collinear midpoints yields the
    exact right-angle path.
    """
    start = end = None
    pts: list[tuple[float, float]] = []
    for token in s.split():
        try:
            if token.startswith("e,"):
                x, y = token[2:].split(",")[:2]
                end = (float(x), float(y))
            elif token.startswith("s,"):
                x, y = token[2:].split(",")[:2]
                start = (float(x), float(y))
            elif "," in token:
                x, y = token.split(",")[:2]
                pts.append((float(x), float(y)))
        except ValueError:
            continue
    if start:
        pts.insert(0, start)
    if end:
        pts.append(end)
    return simplify_polyline(pts)


def simplify_polyline(pts: list[tuple[float, float]], eps: float = 0.5) -> list[tuple[float, float]]:
    out: list[tuple[float, float]] = []
    for p in pts:
        if out and abs(p[0] - out[-1][0]) < eps and abs(p[1] - out[-1][1]) < eps:
            continue
        out.append(p)
    i = 1
    while i < len(out) - 1:
        (ax, ay), (bx, by), (cx, cy) = out[i - 1], out[i], out[i + 1]
        if (abs(ax - bx) < eps and abs(bx - cx) < eps) or (abs(ay - by) < eps and abs(by - cy) < eps):
            out.pop(i)
        else:
            i += 1
    return out


def _safe(name: str) -> str:
    return '"' + name.replace('"', '\\"') + '"'


def _esc(text: str) -> str:
    return text.replace('"', '\\"')


def _wrap_title(label: str, width_in: float) -> str:
    """Greedy word-wrap a cluster title to the given width, joined with DOT's
    ``\\n`` escape so dot reserves a headroom line per wrapped line."""
    per_line = max(int(width_in / CLUSTER_TITLE_CHAR_IN), 6)
    lines: list[str] = []
    cur = ""
    for word in label.split():
        cand = f"{cur} {word}".strip()
        if cur and len(cand) > per_line:
            lines.append(cur)
            cur = word
        else:
            cur = cand
    if cur:
        lines.append(cur)
    return "\\n".join(lines)
