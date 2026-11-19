"""geomcheck catches misalignment, oversized text boxes and arrow problems —
and compose's helpers produce layouts it passes."""

from pptx import Presentation
from pptx.util import Inches

from archgen.shared import compose, geomcheck


def _slide():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs, prs.slides.add_slide(prs.slide_layouts[6])


def _kinds(slide):
    return [(f.severity, f.kind) for f in geomcheck.check_slide(slide, 1)]


def test_near_miss_row_is_flagged():
    _, s = _slide()
    for i, y in enumerate([2.0, 2.0, 2.0, 2.04]):          # last card 2.9pt low
        compose.rrect(s, 1 + i * 2, y, 1.5, 1, fill="EEEEEE", line=None)
    assert ("warn", "Near-miss alignment") in _kinds(s)


def test_row_helper_is_aligned():
    _, s = _slide()
    for x in compose.row(4, 1, 10, 1.5):
        compose.rrect(s, x, compose.snap(2.0), 1.5, 1, fill="EEEEEE", line=None)
    assert "Near-miss alignment" not in [k for _, k in _kinds(s)]


def test_oversized_text_box_is_flagged_and_auto_height_is_not():
    _, s = _slide()
    compose.text(s, 1, 1, 3, 2.0, "Short label", size=12)
    assert ("error", "Text box much taller than its text") in _kinds(s)
    _, s2 = _slide()
    compose.text(s2, 1, 1, 3, None, "Short label", size=12)   # h=None -> fits
    assert "Text box much taller than its text" not in [k for _, k in _kinds(s2)]


def test_arrow_through_unrelated_text_is_flagged():
    _, s = _slide()
    a = compose.rrect(s, 1, 3, 1, 1, fill="EEEEEE", line=None)
    b = compose.rrect(s, 7, 3, 1, 1, fill="EEEEEE", line=None)
    compose.text(s, 3.5, 3.2, 2.4, None, "An unrelated paragraph of card text that the arrow cuts through", size=12)
    compose.link(s, a, b)
    assert ("error", "Arrow crosses a shape it doesn't connect") in _kinds(s)


def test_own_edge_label_is_not_an_obstacle():
    _, s = _slide()
    a = compose.rrect(s, 1, 3, 1, 1, fill="EEEEEE", line=None)
    b = compose.rrect(s, 7, 3, 1, 1, fill="EEEEEE", line=None)
    compose.link(s, a, b)
    compose.text(s, 4, 3.4, 1.0, None, "calls", size=10)    # centred on the arrow
    assert "Arrow crosses a shape it doesn't connect" not in [k for _, k in _kinds(s)]


def test_stacked_arrows_flagged_lanes_fix_it():
    _, s = _slide()
    a = compose.rrect(s, 1, 3, 1, 1, fill="EEEEEE", line=None)
    b = compose.rrect(s, 7, 3, 1, 1, fill="EEEEEE", line=None)
    compose.connector(s, 2, 3.5, 7, 3.5)
    compose.connector(s, 2.5, 3.5, 6.5, 3.5)
    assert ("error", "Arrows overlap") in _kinds(s)
    _, s2 = _slide()
    a = compose.rrect(s2, 1, 3, 1, 1, fill="EEEEEE", line=None)
    b = compose.rrect(s2, 7, 3, 1, 1, fill="EEEEEE", line=None)
    compose.link(s2, a, b, lane=-0.2)
    compose.link(s2, b, a, a_side="left", b_side="right", lane=0.2)
    assert "Arrows overlap" not in [k for _, k in _kinds(s2)]


def test_link_is_glued_to_shapes():
    _, s = _slide()
    a = compose.rrect(s, 1, 3, 1, 1, fill="EEEEEE", line=None)
    b = compose.rrect(s, 7, 3, 1, 1, fill="EEEEEE", line=None)
    cn = compose.link(s, a, b)
    xml = cn._element.xml
    assert "stCxn" in xml and "endCxn" in xml


def test_fix_slide_shrinks_oversized_text_and_snaps_alignment():
    prs, s = _slide()
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(3))
    tb.text_frame.text = "Short line"
    for i, dy in enumerate((0, 0, 0.03)):
        compose.rrect(s, 1 + i * 2.5, 5 + dy, 2, 1, fill="DDDDDD", line=None)
    assert any(k == "Text box much taller than its text" for _, k in _kinds(s))
    assert geomcheck.fix_slide(s) >= 2
    kinds = [k for _, k in _kinds(s)]
    assert "Text box much taller than its text" not in kinds
    assert "Near-miss alignment" not in kinds
    assert abs(tb.top - Inches(1)) < 10          # top anchor kept: text didn't move


def test_vertical_elbow_routes_vertical_first_and_is_modelled():
    prs, s = _slide()
    a = compose.rrect(s, 1, 1, 2, 1, fill="DDDDDD", line=None)
    b = compose.rrect(s, 5, 4, 2, 1, fill="DDDDDD", line=None)
    cn = compose.link(s, a, b, a_side="bottom", b_side="top", bend=0.3)
    path = geomcheck._conn_path(cn)
    (x0, y0), (x1, y1), (x2, y2), (x3, y3) = path
    assert abs(x0 - 2) < 0.01 and abs(y0 - 2) < 0.01      # leaves a's bottom centre
    assert abs(x3 - 6) < 0.01 and abs(y3 - 4) < 0.01      # enters b's top centre
    assert abs(x1 - x0) < 0.01 and abs(y1 - 2.6) < 0.01   # first run is vertical, bend at 30%
    assert not [k for sev, k in _kinds(s) if sev == "error"]
