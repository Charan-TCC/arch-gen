"""Layout diagnostics: the deterministic text report that replaces
render-and-look iteration (scale collapse, wrap risk, auto-direction)."""

from __future__ import annotations

import shutil

import pytest

from archgen.layout.diagnostics import COLLAPSE_SCALE, MIN_READABLE_SCALE, diagnose, has_errors
from archgen.layout.graphviz_layout import Layout, NodeLayout
from archgen.spec.models import Diagram, Edge, Group, Node

AVAIL_W = 10_000_000  # EMU
AVAIL_H = 6_000_000


def _diagram(labels: dict[str, str] | None = None, direction: str = "TB", page: str = "16:9") -> Diagram:
    labels = labels or {"a": "API", "b": "DB"}
    return Diagram(
        title="t",
        kind="flow",
        direction=direction,
        page=page,
        groups=[Group(id="g", label="G")],
        nodes=[Node(id=i, icon_key="generic.service", label=lb, group="g") for i, lb in labels.items()],
        edges=[Edge(src="a", dst="b")],
    )


def _layout(node_w_pt: float = 80.0) -> Layout:
    nodes = {
        i: NodeLayout(id=i, cx=100.0 * (k + 1), cy=100.0, w=node_w_pt, h=60.0)
        for k, i in enumerate(("a", "b"))
    }
    return Layout(graph_bb=(0, 0, 400, 250), nodes=nodes, clusters=[], edges=[])


def test_healthy_scale_is_clean():
    findings = diagnose(_diagram(), _layout(), scale=1.0, avail_w_emu=AVAIL_W, avail_h_emu=AVAIL_H)
    assert not has_errors(findings)


def test_collapsed_scale_is_an_error_with_direction_advice():
    findings = diagnose(
        _diagram(), _layout(), scale=COLLAPSE_SCALE - 0.1, avail_w_emu=AVAIL_W, avail_h_emu=AVAIL_H
    )
    assert has_errors(findings)
    msg = " ".join(f.message for f in findings if f.severity == "error")
    assert "LR" in msg  # TB spec -> advises the flip


def test_cramped_but_readable_scale_is_a_warn_not_error():
    findings = diagnose(
        _diagram(),
        _layout(),
        scale=(COLLAPSE_SCALE + MIN_READABLE_SCALE) / 2,
        avail_w_emu=AVAIL_W,
        avail_h_emu=AVAIL_H,
    )
    assert not has_errors(findings)
    assert any("scaled to" in f.message for f in findings if f.severity == "warn")


def test_no_circular_flip_advice_after_auto_direction():
    findings = diagnose(
        _diagram(),
        _layout(),
        scale=COLLAPSE_SCALE - 0.1,
        avail_w_emu=AVAIL_W,
        avail_h_emu=AVAIL_H,
        auto_direction_tried=True,
    )
    msg = " ".join(f.message for f in findings if f.severity == "error")
    assert "try direction" not in msg
    assert "reshape" in msg


def test_unbreakable_word_warns():
    findings = diagnose(
        _diagram(labels={"a": "Observability", "b": "DB"}),
        _layout(node_w_pt=50.0),
        scale=1.0,
        avail_w_emu=AVAIL_W,
        avail_h_emu=AVAIL_H,
    )
    assert any("Observability" in f.message for f in findings if f.severity == "warn")


def test_direction_flip_is_reported_as_info():
    findings = diagnose(
        _diagram(direction="LR"),
        _layout(),
        scale=1.0,
        avail_w_emu=AVAIL_W,
        avail_h_emu=AVAIL_H,
        direction_flipped=True,
    )
    assert any(f.severity == "info" and "auto-direction" in f.message for f in findings)


def _chain(n: int, page: str) -> Diagram:
    ids = [f"n{i}" for i in range(n)]
    return Diagram(
        title="chain",
        kind="flow",
        direction="TB",
        page=page,
        groups=[Group(id="g", label="G")],
        nodes=[Node(id=i, icon_key="generic.service", label=i, group="g") for i in ids],
        edges=[Edge(src=a, dst=b) for a, b in zip(ids, ids[1:])],
    )


@pytest.mark.skipif(not shutil.which("dot"), reason="graphviz dot not installed")
def test_auto_direction_rescues_a_deep_chain(tmp_path):
    """A 7-rank chain declared TB collapses; auto_direction must flip it to LR."""
    from archgen.render.renderer import Renderer

    r = Renderer()
    r.render(_chain(7, page="16:9"), tmp_path / "chain.pptx", auto_direction=True)
    # The flip to LR must happen and be reported; the chain may still be
    # cramped (honest report), but the advice must not suggest flipping again.
    assert any("auto-direction" in f.message for f in r.diagnostics)
    assert not any("try direction" in f.message for f in r.diagnostics)


@pytest.mark.skipif(not shutil.which("dot"), reason="graphviz dot not installed")
def test_free_page_grows_slide_instead_of_shrinking_text(tmp_path):
    from archgen.render.renderer import Renderer

    r = Renderer()
    r.render(_chain(5, page="free"), tmp_path / "free.pptx")
    default_w, default_h = int(r.theme.slide_w), int(r.theme.slide_h)
    assert (r._slide_w, r._slide_h) != (default_w, default_h)
    assert any("slide expanded" in f.message for f in r.diagnostics)
    # No scale finding may remain once the slide has grown to the floor.
    assert not any(
        "scaled to" in f.message for f in r.diagnostics if f.severity in ("error", "warn")
    )


@pytest.mark.skipif(not shutil.which("dot"), reason="graphviz dot not installed")
def test_free_page_growth_is_capped_and_reported_honestly(tmp_path):
    """A pathological 9-rank chain hits the 3x growth cap: the slide stops at
    the cap and the residual cramping stays visible in the report."""
    from archgen.render.renderer import Renderer

    r = Renderer()
    r.render(_chain(9, page="free"), tmp_path / "capped.pptx")
    assert r._slide_h == 3 * int(r.theme.slide_h)
    assert any("scaled to" in f.message for f in r.diagnostics if f.severity == "warn")


@pytest.mark.skipif(not shutil.which("dot"), reason="graphviz dot not installed")
def test_169_page_locks_slide_and_demands_reshaping(tmp_path):
    from archgen.render.renderer import Renderer

    r = Renderer()
    r.render(_chain(9, page="16:9"), tmp_path / "deck.pptx")
    assert (r._slide_w, r._slide_h) == (int(r.theme.slide_w), int(r.theme.slide_h))
    assert any("will NOT grow" in f.message for f in r.diagnostics)


def test_lint_requires_page_for_flow_specs():
    from archgen.spec.lint import lint

    d = _diagram().model_copy(update={"page": None})
    assert any("page format not declared" in e for e in lint(d))
    assert not any("page format" in e for e in lint(_diagram()))
