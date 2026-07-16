"""Layout dispatcher: prefer Graphviz `dot`; fall back to the pure-Python layout."""

from __future__ import annotations

import shutil

from archgen.layout.graphviz_layout import Layout
from archgen.spec.models import Diagram


def compute_layout(
    diagram: Diagram,
    engine: str = "auto",
    *,
    node_w: float | None = None,
    node_h: float | None = None,
) -> tuple[Layout, str]:
    """Return ``(layout, engine_used)``.

    ``engine``: ``"auto"`` (dot if available else simple), ``"graphviz"``, or ``"simple"``.
    ``node_w``/``node_h`` let the renderer override card proportions (used for the
    layered auto-sizing path); they only apply to the graphviz flow layout.
    """
    if diagram.kind.value == "network":
        from archgen.layout.network_layout import compute_layout as network

        return network(diagram), "network"

    if engine in ("auto", "graphviz") and shutil.which("dot"):
        from archgen.layout.graphviz_layout import compute_layout as gv

        return gv(diagram, node_w=node_w, node_h=node_h), "graphviz"
    if engine == "graphviz":
        raise RuntimeError("engine='graphviz' requested but `dot` is not on PATH.")

    # The pure-Python fallback renders uniform cards; layered banding is
    # graphviz-only, so a layered spec degrades to the ordinary flow layout here.
    from archgen.layout.simple_layout import compute_layout as simple

    return simple(diagram), "simple"
