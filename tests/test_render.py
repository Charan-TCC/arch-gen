"""Golden render test: the sample spec (all four icon sources) renders to a pptx
whose labels are editable text, whose icons are picture objects, and whose shapes
stay on-slide."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from archgen.registry.resolver import IconResolver
from archgen.render.renderer import render_spec
from archgen.render.renderer import Renderer
from archgen.spec.io import load_spec
from archgen.spec.lint import lint

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_spec.yaml"
NETWORK_SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "azure_private_aks.yaml"


def _walk(shapes):
    """Yield every shape, descending into group shapes (nodes are grouped:
    one picture + one label per group), so leaf pictures/labels are visible."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(shape.shapes)
        else:
            yield shape


@pytest.fixture(scope="module")
def rendered(tmp_path_factory) -> Path:
    out = tmp_path_factory.mktemp("out") / "sample.pptx"
    diagram = load_spec(SAMPLE)
    assert lint(diagram) == []
    return render_spec(diagram, out)


def test_sample_spans_all_four_sources():
    d = load_spec(SAMPLE)
    r = IconResolver()
    providers = {r.get(n.icon_key).provider for n in d.nodes if r.exists(n.icon_key)}
    assert {"azure", "aws", "gcp", "tech"} <= providers


def test_file_opens_and_has_one_slide(rendered):
    prs = Presentation(str(rendered))
    assert len(prs.slides) == 1


def test_every_node_has_a_picture(rendered):
    prs = Presentation(str(rendered))
    pics = [s for s in _walk(prs.slides[0].shapes) if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    n_nodes = len(load_spec(SAMPLE).nodes)
    assert len(pics) == n_nodes  # one editable picture object per node


def test_labels_are_editable_text(rendered):
    prs = Presentation(str(rendered))
    texts = {s.text_frame.text for s in _walk(prs.slides[0].shapes) if s.has_text_frame}
    # Node labels and container titles are real, editable text (not rasterized).
    for expected in ["API on AKS", "Workers on GKE", "Compute Tier (Azure)", "Data Tier"]:
        assert expected in texts


def test_edges_have_arrowheads(rendered):
    prs = Presentation(str(rendered))
    xml = prs.slides[0].shapes._spTree.xml
    diagram = load_spec(SAMPLE)
    semantics = {(edge.style.value, edge.kind.value) for edge in diagram.edges}
    n_legend_lines = len(semantics) if len(semantics) > 1 else 0
    expected = len(diagram.edges) + n_legend_lines
    assert xml.count("<p:cxnSp>") == expected
    assert xml.count("tailEnd") == expected


def test_legends_present(rendered):
    prs = Presentation(str(rendered))
    texts = {s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame}
    assert "Providers" in texts
    assert "Abbreviations" in texts


def test_flow_connection_legend_is_derived_from_multiple_line_semantics(rendered):
    prs = Presentation(str(rendered))
    texts = {s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame}
    assert "Connections" in texts
    assert "Request / data flow" in texts


def test_all_shapes_on_slide(rendered):
    prs = Presentation(str(rendered))
    W, H = prs.slide_width, prs.slide_height
    for s in prs.slides[0].shapes:
        if s.left is None or s.width is None:
            continue
        assert s.left >= -1 and s.top >= -1
        assert s.left + s.width <= W + 1 and s.top + (s.height or 0) <= H + 1


def test_network_example_lints_and_exposes_connection_semantics():
    diagram = load_spec(NETWORK_SAMPLE)
    assert lint(diagram) == []
    assert diagram.kind.value == "network"
    assert {edge.kind.value for edge in diagram.edges} == {
        "https", "diagnostics", "dns", "private_connection", "virtual_network_link"
    }


def test_network_flattening_removes_inherited_shape_effects():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100000, 100000)
    assert shape._element.find(qn("p:style")) is not None
    Renderer()._remove_effects(shape)
    assert shape._element.find(qn("p:style")) is None
    assert shape._element.spPr.find(qn("a:effectLst")) is not None


def test_legend_reserve_grows_with_abbreviations():
    from archgen.render.legend import legends_reserve_height
    from archgen.render.theme import load_theme
    from archgen.spec.models import Diagram, Edge, Node

    theme = load_theme()
    base = dict(
        title="t",
        page="16:9",
        nodes=[
            Node(id="a", icon_key="generic.service", label="A"),
            Node(id="b", icon_key="generic.service", label="B"),
        ],
        edges=[Edge(src="a", dst="b")],
    )
    small = Diagram(**base)
    many = Diagram(
        **{**base, "title": "CDN DNS VPC IAM ECS ECR"},
        abbreviations={
            k: k for k in ("CDN", "DNS", "VPC", "IAM", "ECS", "ECR")
        },
    )
    assert legends_reserve_height(many, theme) > legends_reserve_height(small, theme)


def test_two_column_abbrev_grid():
    from archgen.render.legend import _abbrev_grid

    assert _abbrev_grid(4) == (4, 1)
    assert _abbrev_grid(10) == (5, 2)


def test_cluster_title_wrap_for_dot():
    from archgen.layout.graphviz_layout import _wrap_title

    # Narrow cluster: two words wrap to two dot label lines.
    assert _wrap_title("Backend Engine", 1.2) == "Backend\\nEngine"
    # Wide cluster: single line.
    assert _wrap_title("Backend Engine", 3.0) == "Backend Engine"
