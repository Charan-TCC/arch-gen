"""Deterministic semantic layout for cloud-network architecture diagrams.

Unlike the generic flow engines, this layout uses architectural roles rather
than edge ranks: external clients sit left of the cloud boundary, DNS services
occupy the control plane above the virtual network, managed services sit below
it, and ingress/compute/management/private-access containers form a stable 2x2
grid inside the network. The spec still contains no coordinates or visual rules.
"""

from __future__ import annotations

from archgen.layout.graphviz_layout import ClusterLayout, EdgeLayout, Layout, NodeLayout
from archgen.spec.models import Diagram, GroupKind, GroupRole, NodeRole

CANVAS_W = 990.0
CANVAS_H = 810.0

# Fixed semantic zones in top-left/y-down points. Transforming to the shared
# y-up convention happens only when the public Layout is assembled.
_CLOUD = (291.0, 34.5, 975.0, 799.5)
_NETWORK = (310.5, 282.0, 892.5, 634.5)
_ROLE_BOXES = {
    GroupRole.ingress: (331.5, 324.0, 498.0, 450.0),
    GroupRole.compute: (539.25, 324.0, 870.75, 450.0),
    GroupRole.management: (331.5, 489.75, 498.0, 615.0),
    GroupRole.private_access: (539.25, 489.75, 870.75, 615.0),
}

_ROLE_NODE_CENTERS = {
    GroupRole.ingress: [(378.0, 396.0), (451.0, 396.0)],
    GroupRole.compute: [(632.5, 386.0), (784.5, 386.0)],
    GroupRole.management: [(416.5, 552.0)],
    GroupRole.private_access: [(584.0, 549.0), (664.5, 549.0), (748.0, 549.0), (830.0, 549.0)],
}

_ARCHITECTURE_ROLES = {
    GroupRole.web,
    GroupRole.storage,
    GroupRole.private_network,
    GroupRole.data,
    GroupRole.discovery,
    GroupRole.cost,
    GroupRole.deployment,
    GroupRole.integration,
}
_ARCH_CLOUD = (180.0, 25.0, 975.0, 790.0)
_ARCH_NETWORK = (215.0, 270.0, 945.0, 525.0)
_ARCH_ROLE_BOXES = {
    GroupRole.web: (215.0, 55.0, 610.0, 250.0),
    GroupRole.storage: (630.0, 55.0, 945.0, 250.0),
    GroupRole.private_network: (230.0, 295.0, 930.0, 510.0),
    GroupRole.data: (245.0, 320.0, 570.0, 495.0),
    GroupRole.discovery: (590.0, 320.0, 915.0, 495.0),
    GroupRole.cost: (215.0, 545.0, 450.0, 765.0),
    GroupRole.deployment: (470.0, 545.0, 700.0, 765.0),
    GroupRole.integration: (720.0, 545.0, 945.0, 765.0),
}


def compute_layout(diagram: Diagram) -> Layout:
    architecture_board = any(group.role in _ARCHITECTURE_ROLES for group in diagram.groups)
    boxes = _group_boxes(diagram)
    positions: dict[str, tuple[float, float, float, float]] = {}

    by_role: dict[NodeRole, list] = {}
    for node in diagram.nodes:
        by_role.setdefault(node.role, []).append(node)

    # Global semantic zones override ordinary group placement.
    if architecture_board:
        _place_vertical(positions, by_role.get(NodeRole.external, []), x=90.0, start_y=145.0, step=90.0)
        _place_vertical(positions, by_role.get(NodeRole.internet, []), x=150.0, start_y=175.0, step=62.0)
    else:
        _place_vertical(positions, by_role.get(NodeRole.external, []), x=146.5, start_y=383.0, step=145.0)
        _place_vertical(positions, by_role.get(NodeRole.internet, []), x=248.0, start_y=385.0, step=62.0)

    for i, node in enumerate(by_role.get(NodeRole.dns, [])):
        w, h = _node_size(node.role)
        positions[node.id] = (577.0 + i * 47.0, 70.0 + i * 42.0, w, h)

    for i, node in enumerate(by_role.get(NodeRole.network_anchor, [])):
        w, h = _node_size(node.role)
        positions[node.id] = (519.0 + i * 60.0, 278.0, w, h)

    platform = by_role.get(NodeRole.managed_service, []) + by_role.get(NodeRole.observability, [])
    for i, node in enumerate(platform):
        w, h = _node_size(node.role)
        positions[node.id] = (664.0 + i * 83.0, 686.0, w, h)

    # Remaining nodes are distributed inside their immediate semantic container.
    placed = set(positions)
    members: dict[str, list] = {}
    for node in diagram.nodes:
        if node.id not in placed and node.group is not None:
            members.setdefault(node.group, []).append(node)

    groups_by_id = {g.id: g for g in diagram.groups}
    for gid, nodes in members.items():
        box = boxes.get(gid, _NETWORK)
        group = groups_by_id.get(gid)
        if architecture_board and group is not None and group.role in _ARCHITECTURE_ROLES:
            _place_grid(positions, nodes, box)
            continue
        prescribed = _ROLE_NODE_CENTERS.get(group.role if group else GroupRole.default, [])
        if prescribed and len(nodes) <= len(prescribed):
            for node, (cx, cy) in zip(nodes, prescribed):
                w, h = _node_size(node.role)
                positions[node.id] = (cx, cy, w, h)
            continue
        x0, y0, x1, y1 = box
        usable_left, usable_right = x0 + 12.0, x1 - 12.0
        center_y = (y0 + y1) / 2.0 + 8.0
        widths = [_node_size(n.role)[0] for n in nodes]
        total_w = sum(widths)
        gap = max(8.0, min(30.0, (usable_right - usable_left - total_w) / max(len(nodes) - 1, 1)))
        row_w = total_w + gap * max(len(nodes) - 1, 0)
        cursor = (x0 + x1 - row_w) / 2.0
        for node, width in zip(nodes, widths):
            w, h = _node_size(node.role)
            positions[node.id] = (cursor + width / 2.0, center_y, w, h)
            cursor += width + gap

    # Any ungrouped service nodes use a final safe lane inside the network.
    remainder = [n for n in diagram.nodes if n.id not in positions]
    for i, node in enumerate(remainder):
        w, h = _node_size(node.role)
        positions[node.id] = (340.0 + i * 110.0, 300.0, w, h)

    nodes = {
        nid: NodeLayout(id=nid, cx=cx, cy=CANVAS_H - cy_down, w=w, h=h)
        for nid, (cx, cy_down, w, h) in positions.items()
    }
    groups = {g.id: g for g in diagram.groups}
    clusters = []
    for gid, (x0, y0, x1, y1) in boxes.items():
        g = groups[gid]
        clusters.append(
            ClusterLayout(id=gid, label=g.label, bb=(x0, CANVAS_H - y1, x1, CANVAS_H - y0))
        )

    return Layout(
        graph_bb=(0.0, 0.0, CANVAS_W, CANVAS_H),
        nodes=nodes,
        clusters=clusters,
        edges=[EdgeLayout(src=e.src, dst=e.dst, points=[]) for e in diagram.edges],
    )


def _group_boxes(diagram: Diagram) -> dict[str, tuple[float, float, float, float]]:
    out: dict[str, tuple[float, float, float, float]] = {}
    architecture_board = any(group.role in _ARCHITECTURE_ROLES for group in diagram.groups)
    fallback_index = 0
    for group in diagram.groups:
        if architecture_board:
            if group.kind == GroupKind.cloud or group.role == GroupRole.boundary:
                out[group.id] = _ARCH_CLOUD
            elif group.role in _ARCH_ROLE_BOXES:
                out[group.id] = _ARCH_ROLE_BOXES[group.role]
            elif group.kind == GroupKind.network:
                out[group.id] = _ARCH_NETWORK
            else:
                col, row = fallback_index % 3, fallback_index // 3
                x0 = 215.0 + col * 245.0
                y0 = 545.0 + row * 115.0
                out[group.id] = (x0, y0, x0 + 225.0, y0 + 100.0)
                fallback_index += 1
            continue
        if group.kind == GroupKind.cloud or group.role == GroupRole.boundary:
            out[group.id] = _CLOUD
        elif group.kind == GroupKind.network:
            out[group.id] = _NETWORK
        elif group.role in _ROLE_BOXES:
            out[group.id] = _ROLE_BOXES[group.role]
        else:
            # Stable fallback grid for additional network subcontainers.
            col, row = fallback_index % 2, fallback_index // 2
            x0 = 331.5 + col * 207.75
            y0 = 324.0 + row * 165.75
            out[group.id] = (x0, y0, x0 + 166.5, y0 + 126.0)
            fallback_index += 1
    return out


def _place_vertical(positions, nodes, *, x: float, start_y: float, step: float) -> None:
    offset = -step * (len(nodes) - 1) / 2.0
    for i, node in enumerate(nodes):
        w, h = _node_size(node.role)
        positions[node.id] = (x, start_y + offset + i * step, w, h)


def _place_grid(positions, nodes, box) -> None:
    """Place a dense architecture section in a stable, readable icon grid."""
    x0, y0, x1, y1 = box
    count = len(nodes)
    if not count:
        return
    width = x1 - x0
    # Choose columns so each cell is wide enough (~90pt) to hold a node label
    # without colliding with its neighbour; narrow sections drop to fewer columns.
    max_cols = max(1, min(4, int((width - 24.0) / 90.0)))
    cols = min(max_cols, count)
    if count == 5:
        cols = min(cols, 3)
    rows = (count + cols - 1) // cols
    left, right = x0 + 12.0, x1 - 12.0
    top, bottom = y0 + 30.0, y1 - 10.0
    cell_w = (right - left) / cols
    cell_h = (bottom - top) / rows
    for i, node in enumerate(nodes):
        col, row = i % cols, i // cols
        w, h = _node_size(node.role)
        cx = left + cell_w * (col + 0.5)
        cy = top + cell_h * (row + 0.5)
        positions[node.id] = (cx, cy, w, h)


def _node_size(role: NodeRole) -> tuple[float, float]:
    return {
        NodeRole.external: (116.0, 78.0),
        NodeRole.internet: (79.0, 52.0),
        NodeRole.dns: (270.0, 36.0),
        NodeRole.gateway: (60.0, 70.0),
        NodeRole.workload: (108.0, 48.0),
        NodeRole.management: (66.0, 66.0),
        NodeRole.private_endpoint: (72.0, 62.0),
        NodeRole.network_anchor: (54.0, 34.0),
        NodeRole.managed_service: (90.0, 72.0),
        NodeRole.observability: (90.0, 72.0),
        NodeRole.service: (66.0, 60.0),
    }[role]
