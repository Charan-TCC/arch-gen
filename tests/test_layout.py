"""Layout tests: the pure-Python fallback keeps group containers non-overlapping."""

from __future__ import annotations

import shutil

import pytest

from archgen.layout.engine import compute_layout as compute_flow_layout
from archgen.layout.graphviz_layout import build_dot, layered_card_size
from archgen.layout.simple_layout import compute_layout
from archgen.layout.network_layout import CANVAS_H, compute_layout as compute_network_layout
from archgen.spec.models import Diagram, Edge, Group, Node


def _multi_group() -> Diagram:
    return Diagram(
        title="t",
        direction="LR",
        groups=[Group(id="g1", label="One"), Group(id="g2", label="Two"), Group(id="g3", label="Three")],
        nodes=[
            Node(id="a", icon_key="generic.client", label="A", group="g1"),
            Node(id="b", icon_key="generic.client", label="B", group="g1"),
            Node(id="c", icon_key="generic.client", label="C", group="g2"),
            Node(id="d", icon_key="generic.client", label="D", group="g2"),
            Node(id="e", icon_key="generic.client", label="E", group="g3"),
        ],
        edges=[Edge(src="a", dst="c"), Edge(src="b", dst="c"), Edge(src="c", dst="e"), Edge(src="d", dst="e")],
        abbreviations={},
    )


def _overlaps(a, b) -> bool:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    # Positive-area intersection (touching edges is allowed).
    return min(ax1, bx1) - max(ax0, bx0) > 1e-6 and min(ay1, by1) - max(ay0, by0) > 1e-6


def test_group_containers_do_not_overlap():
    layout = compute_layout(_multi_group())
    boxes = [c.bb for c in layout.clusters]
    assert len(boxes) == 3
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            assert not _overlaps(boxes[i], boxes[j]), f"containers {i} and {j} overlap"


def test_all_nodes_positioned():
    d = _multi_group()
    layout = compute_layout(d)
    assert set(layout.nodes) == {n.id for n in d.nodes}


def test_nested_group_inside_parent():
    d = Diagram(
        title="t",
        groups=[Group(id="outer", label="Outer"), Group(id="inner", label="Inner", parent="outer")],
        nodes=[
            Node(id="x", icon_key="generic.client", label="X", group="outer"),
            Node(id="y", icon_key="generic.client", label="Y", group="inner"),
        ],
        edges=[Edge(src="x", dst="y")],
        abbreviations={},
    )
    layout = compute_layout(d)
    boxes = {c.id: c.bb for c in layout.clusters}
    ox0, oy0, ox1, oy1 = boxes["outer"]
    ix0, iy0, ix1, iy1 = boxes["inner"]
    # Inner box is contained within the outer box.
    assert ox0 <= ix0 and oy0 <= iy0 and ox1 >= ix1 and oy1 >= iy1


def test_network_layout_uses_semantic_zones():
    d = Diagram.model_validate(
        {
            "title": "Private network",
            "kind": "network",
            "groups": [
                {"id": "cloud", "label": "Cloud", "kind": "cloud", "role": "boundary"},
                {"id": "vnet", "label": "VNet", "parent": "cloud", "kind": "network"},
                {"id": "ingress", "label": "Ingress", "parent": "vnet", "kind": "subnet", "role": "ingress"},
            ],
            "nodes": [
                {"id": "client", "icon_key": "generic.client", "label": "Client", "role": "external"},
                {"id": "dns", "icon_key": "generic.database", "label": "DNS", "group": "cloud", "role": "dns"},
                {"id": "gateway", "icon_key": "generic.client", "label": "Gateway", "group": "ingress", "role": "gateway"},
                {"id": "store", "icon_key": "generic.database", "label": "Store", "group": "cloud", "role": "managed_service"},
            ],
            "edges": [
                {"src": "client", "dst": "gateway", "kind": "https"},
                {"src": "gateway", "dst": "dns", "kind": "dns"},
                {"src": "gateway", "dst": "store", "kind": "private_connection"},
            ],
            "abbreviations": {"DNS": "Domain Name System"},
        }
    )
    layout = compute_network_layout(d)
    assert layout.graph_bb == (0.0, 0.0, 990.0, 810.0)
    cloud = next(c.bb for c in layout.clusters if c.id == "cloud")
    network = next(c.bb for c in layout.clusters if c.id == "vnet")
    assert layout.nodes["client"].cx < cloud[0]
    assert CANVAS_H - layout.nodes["dns"].cy < CANVAS_H - network[3]
    assert CANVAS_H - layout.nodes["store"].cy > CANVAS_H - network[1]


def _layered_diagram() -> Diagram:
    """A layered board with an uneven Source tier (4) vs. the rest (2)."""
    return Diagram(
        title="Layered",
        direction="LR",
        layered=True,
        groups=[Group(id="src", label="Source"), Group(id="ingest", label="Ingest"), Group(id="serve", label="Serve")],
        nodes=[
            Node(id="s1", icon_key="generic.client", label="S1", group="src"),
            Node(id="s2", icon_key="generic.client", label="S2", group="src"),
            Node(id="s3", icon_key="generic.client", label="S3", group="src"),
            Node(id="s4", icon_key="generic.client", label="S4", group="src"),
            Node(id="i1", icon_key="generic.client", label="I1", group="ingest"),
            Node(id="i2", icon_key="generic.client", label="I2", group="ingest"),
            Node(id="v1", icon_key="generic.client", label="V1", group="serve"),
            Node(id="v2", icon_key="generic.client", label="V2", group="serve"),
        ],
        edges=[
            Edge(src="s1", dst="i1"), Edge(src="s2", dst="i1"),
            Edge(src="s3", dst="i2"), Edge(src="s4", dst="i2"),
            Edge(src="i1", dst="v1"), Edge(src="i2", dst="v2"),
        ],
        abbreviations={},
    )


def test_layered_pins_each_tier_to_one_rank():
    # Every multi-node tier gets a rank=same constraint; a non-layered build gets none.
    dot = build_dot(_layered_diagram())
    assert dot.count("rank=same") == 3
    plain = _layered_diagram()
    plain.layered = False
    assert "rank=same" not in build_dot(plain)


def test_layered_card_size_shrinks_for_denser_tiers():
    # Renderer-owned sizing: a 6-node tier must yield shorter cards than a 2-node one.
    def _tier(n: int) -> Diagram:
        nodes = [Node(id=f"n{i}", icon_key="generic.client", label=f"N{i}", group="g") for i in range(n)]
        return Diagram(
            title="t", direction="LR", layered=True,
            groups=[Group(id="g", label="G")],
            nodes=nodes,
            edges=[Edge(src=f"n{i}", dst=f"n{i+1}") for i in range(n - 1)],
            abbreviations={},
        )

    _, h_sparse = layered_card_size(_tier(2), 11.5, 5.3)
    _, h_dense = layered_card_size(_tier(6), 11.5, 5.3)
    assert h_dense < h_sparse


@pytest.mark.skipif(not shutil.which("dot"), reason="graphviz dot not installed")
def test_layered_tiers_align_into_columns():
    layout, engine = compute_flow_layout(_layered_diagram(), node_w=1.6, node_h=0.7)
    assert engine == "graphviz"
    # In LR each tier shares a rank -> its nodes line up at one center-x (a column).
    for gid, ids in {"src": ["s1", "s2", "s3", "s4"], "ingest": ["i1", "i2"], "serve": ["v1", "v2"]}.items():
        xs = {round(layout.nodes[i].cx, 1) for i in ids}
        assert len(xs) == 1, f"tier {gid} not column-aligned: {xs}"
