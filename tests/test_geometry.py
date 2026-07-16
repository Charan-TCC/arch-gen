"""Geometry unit tests: points -> EMU with correct scale, centering, and Y-flip."""

from __future__ import annotations

from archgen.layout.geometry import EMU_PER_POINT, Transform


def _slide():
    # 10in x 10in slide, 1in margin -> 8in x 8in available.
    return dict(slide_w_emu=int(10 * 914400), slide_h_emu=int(10 * 914400), margin_emu=int(1 * 914400))


def test_uniform_scale_and_fit():
    # Graph 72x72 pt = 1in x 1in -> scaled to fill 8in available -> 8x.
    tf = Transform((0, 0, 72, 72), **_slide())
    w, h = tf.size(72, 72)
    assert abs(w - 8 * 914400) < 2 and abs(h - 8 * 914400) < 2


def test_y_is_flipped():
    tf = Transform((0, 0, 72, 72), **_slide())
    # Top of graph (y=72, y-up) must map ABOVE bottom (y=0) on the slide (smaller EMU y).
    _, y_top = tf.point(0, 72)
    _, y_bottom = tf.point(0, 0)
    assert y_top < y_bottom


def test_centering_square_on_square():
    tf = Transform((0, 0, 72, 72), **_slide())
    # A centered graph point maps to the slide center.
    cx, cy = tf.point(36, 36)
    assert abs(cx - 5 * 914400) < 2 and abs(cy - 5 * 914400) < 2


def test_node_rect_is_centered_on_pos():
    tf = Transform((0, 0, 100, 100), **_slide())
    rect = tf.node_rect(50, 50, 20, 10)
    # Rect center should equal the transform of the node center.
    px, py = tf.point(50, 50)
    assert abs(rect.cx - px) <= 1 and abs(rect.cy - py) <= 1


def test_bbox_rect_orientation():
    tf = Transform((0, 0, 100, 100), **_slide())
    rect = tf.bbox_rect(10, 10, 90, 90)
    assert rect.width > 0 and rect.height > 0
