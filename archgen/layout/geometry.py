"""Coordinate transform: Graphviz space -> PowerPoint EMU.

This is the most error-prone step in the pipeline (see plan risk note), so it
lives in its own module with a dedicated unit test.

Graphviz coordinate contract (from `dot -Tjson`):
  * units are **points** (72 per inch);
  * origin is **bottom-left**, y grows **upward**;
  * a node `pos` is the node **center**; `bb` is "x0,y0,x1,y1".

PowerPoint contract:
  * units are **EMU** (914400 per inch);
  * origin is **top-left**, y grows **downward**.

So we (1) scale the whole graph bounding box to fit the slide content area
(uniform scale, preserving aspect ratio), (2) center it, and (3) flip Y.
"""

from __future__ import annotations

from dataclasses import dataclass

EMU_PER_INCH = 914400
POINTS_PER_INCH = 72
EMU_PER_POINT = EMU_PER_INCH / POINTS_PER_INCH  # 12700


@dataclass(frozen=True)
class Rect:
    """Axis-aligned rectangle in EMU, top-left origin (PowerPoint convention)."""

    left: int
    top: int
    width: int
    height: int

    @property
    def cx(self) -> int:
        return self.left + self.width // 2

    @property
    def cy(self) -> int:
        return self.top + self.height // 2


class Transform:
    """Maps Graphviz points -> slide EMU with uniform scale, centering, Y-flip."""

    def __init__(
        self,
        graph_bb: tuple[float, float, float, float],
        slide_w_emu: int,
        slide_h_emu: int,
        margin_emu: int,
        top_reserve_emu: int = 0,
        bottom_reserve_emu: int = 0,
    ):
        self._x0, self._y0, self._x1, self._y1 = graph_bb
        self._gw = max(self._x1 - self._x0, 1.0)  # graph width in points
        self._gh = max(self._y1 - self._y0, 1.0)  # graph height in points

        # Reserved bands (e.g. a legend strip) shrink the drawable area so
        # content is scaled to fit *below/above* them instead of overlapping.
        top = margin_emu + top_reserve_emu
        bottom = margin_emu + bottom_reserve_emu
        avail_w = slide_w_emu - 2 * margin_emu
        avail_h = slide_h_emu - top - bottom
        # Uniform scale (EMU per point) to fit within the available area.
        self._scale = min(avail_w / (self._gw * EMU_PER_POINT), avail_h / (self._gh * EMU_PER_POINT))

        drawn_w = self._gw * EMU_PER_POINT * self._scale
        drawn_h = self._gh * EMU_PER_POINT * self._scale
        # Center the drawn graph within the available band.
        self._off_x = (slide_w_emu - drawn_w) / 2
        self._off_y = top + (avail_h - drawn_h) / 2
        self._slide_h = slide_h_emu

    @property
    def scale(self) -> float:
        """Dimensionless fit factor: 1.0 draws the graph at its natural Graphviz
        size; below ~0.75 node cards shrink past label readability."""
        return self._scale

    def _emu_x(self, x_pt: float) -> float:
        return self._off_x + (x_pt - self._x0) * EMU_PER_POINT * self._scale

    def _emu_y_flip(self, y_pt: float) -> float:
        # Flip: distance from the top of the graph, measured downward.
        return self._off_y + (self._y1 - y_pt) * EMU_PER_POINT * self._scale

    def point(self, x_pt: float, y_pt: float) -> tuple[int, int]:
        """Transform a single Graphviz point (center-origin, y-up) to EMU."""
        return int(round(self._emu_x(x_pt))), int(round(self._emu_y_flip(y_pt)))

    def size(self, w_pt: float, h_pt: float) -> tuple[int, int]:
        """Scale a width/height in points to EMU (no translation, no flip)."""
        return (
            int(round(w_pt * EMU_PER_POINT * self._scale)),
            int(round(h_pt * EMU_PER_POINT * self._scale)),
        )

    def node_rect(self, cx_pt: float, cy_pt: float, w_pt: float, h_pt: float) -> Rect:
        """Graphviz node (center + size in points) -> EMU Rect (top-left origin)."""
        w_emu, h_emu = self.size(w_pt, h_pt)
        cx_emu, cy_emu = self.point(cx_pt, cy_pt)
        return Rect(left=cx_emu - w_emu // 2, top=cy_emu - h_emu // 2, width=w_emu, height=h_emu)

    def bbox_rect(self, x0: float, y0: float, x1: float, y1: float) -> Rect:
        """Graphviz bounding box "x0,y0,x1,y1" (y-up) -> EMU Rect (top-left)."""
        left, bottom = self.point(x0, y0)
        right, top = self.point(x1, y1)  # y-flip makes y1 the *top*
        return Rect(left=left, top=top, width=right - left, height=bottom - top)
