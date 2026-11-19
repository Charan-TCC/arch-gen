"""geomcheck — layout-geometry QA for any .pptx (python-pptx only).

SHARED FILE. Identical copy in Arch Gen (archgen/shared/geomcheck.py) and PPT Gen
(pptgen/shared/geomcheck.py); the parity test fails on drift. Contract:
docs/COMPOSE_API.md.

Catches the three defects that slip past eyeballing:

  near-miss alignment   shapes meant to line up (same row / column) that are off
                        by a hair — tops, centres, bottoms, lefts, rights;
  oversized text boxes  a text box far taller (or wider) than its text needs, so
                        it overlaps neighbours and breaks vertical centring;
  arrow problems        connectors that cut through text/cards they do not
                        connect, that run on top of each other, or that end
                        short of / past their target.

Findings are (severity, slide_no, kind, message). "error" = almost certainly
wrong; "warn" = probably wrong, confirm.  CLI:  python -m pptgen.shared.geomcheck  (or archgen.shared.geomcheck) deck.pptx
"""

from __future__ import annotations

import math
import sys
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

EMU_IN = 914400

# Tolerances (inches)
ALIGN_SNAP = 0.055      # differences up to ~4pt are "meant to align but missed"
ALIGN_MIN = 0.006       # below this it's rounding, not a defect
TEXT_SLACK_IN = 0.30    # allowed empty height beyond the text's needs
TEXT_SLACK_RATIO = 1.8  # ...or this multiple of the needed height
END_TOL = 0.10          # connector end within this of a shape edge = attached
GAP_WARN = 0.15         # connector end this far from the nearest shape = floating


@dataclass
class Finding:
    severity: str
    slide: int
    kind: str
    message: str

    def __str__(self) -> str:
        return f"[{self.severity.upper()}] slide {self.slide} · {self.kind}: {self.message}"


@dataclass
class _Box:
    name: str
    l: float
    t: float
    w: float
    h: float
    text: str
    kind: str          # text | shape | picture | group
    sizes: tuple = ()
    anchor: str = "top"
    wrap: bool = True
    group: int | None = None
    shape: object = field(default=None, repr=False, compare=False)

    @property
    def r(self): return self.l + self.w
    @property
    def b(self): return self.t + self.h
    @property
    def cx(self): return self.l + self.w / 2
    @property
    def cy(self): return self.t + self.h / 2

    def label(self) -> str:
        return f"{self.name!r}" + (f" ({self.text[:28]!r})" if self.text else "")


@dataclass
class _Conn:
    name: str
    x1: float
    y1: float
    x2: float
    y2: float
    elbow: bool
    path: tuple = ()

    def segments(self):
        if self.path:
            return [(*self.path[i], *self.path[i + 1]) for i in range(len(self.path) - 1)]
        if self.elbow and abs(self.x1 - self.x2) > 1e-3 and abs(self.y1 - self.y2) > 1e-3:
            xm = (self.x1 + self.x2) / 2
            return [(self.x1, self.y1, xm, self.y1), (xm, self.y1, xm, self.y2),
                    (xm, self.y2, self.x2, self.y2)]
        return [(self.x1, self.y1, self.x2, self.y2)]


def _conn_path(sh) -> tuple | None:
    """Drawn route of a connector in inches, honouring rot / flips / adj1
    (python-pptx's begin_x/end_x ignore rotation)."""
    sp = sh._element.spPr
    xf = sp.find(qn("a:xfrm"))
    if xf is None or xf.find(qn("a:off")) is None:
        return None
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    ox, oy = int(off.get("x")) / EMU_IN, int(off.get("y")) / EMU_IN
    w, h = int(ext.get("cx")) / EMU_IN, int(ext.get("cy")) / EMU_IN
    rot = int(xf.get("rot", "0")) / 60000.0
    fh, fv = xf.get("flipH") in ("1", "true"), xf.get("flipV") in ("1", "true")
    geom = sp.find(qn("a:prstGeom"))
    prst = geom.get("prst") if geom is not None else "line"
    if prst.startswith("bentConnector") and w > 1e-4 and h > 1e-4:
        a = 0.5
        for gd in geom.iter(qn("a:gd")):
            if gd.get("name") == "adj1":
                a = int(gd.get("fmla").split()[-1]) / 100000.0
        local = [(0, 0), (a * w, 0), (a * w, h), (w, h)]
    else:
        local = [(0, 0), (w, h)]
    cx, cy = ox + w / 2, oy + h / 2
    cs, sn = math.cos(math.radians(rot)), math.sin(math.radians(rot))
    out = []
    for u, v in local:
        if fh:
            u = w - u
        if fv:
            v = h - v
        du, dv = u - w / 2, v - h / 2
        out.append((cx + du * cs - dv * sn, cy + du * sn + dv * cs))
    return tuple(out)


# ------------------------------------------------------------------ extraction

def _in(v) -> float:
    return 0.0 if v is None else int(v) / EMU_IN


def _walk(shapes, gid=None, counter=[0]):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            counter[0] += 1
            yield from _walk(sh.shapes, counter[0], counter)
        else:
            yield sh, gid


def _collect(slide):
    boxes, conns = [], []
    for sh, gid in _walk(slide.shapes):
        if sh.left is None or sh.width is None:
            continue
        if sh.shape_type == MSO_SHAPE_TYPE.LINE or sh.__class__.__name__ == "Connector":
            try:
                path = _conn_path(sh)
                if path:
                    (bx, by), (ex, ey) = path[0], path[-1]
                    conns.append(_Conn(sh.name, bx, by, ex, ey, len(path) > 2, path))
                else:
                    conns.append(_Conn(sh.name, _in(sh.begin_x), _in(sh.begin_y),
                                       _in(sh.end_x), _in(sh.end_y),
                                       "bentConnector" in sh._element.xml))
            except Exception:
                pass
            continue
        text, sizes, anchor, wrap = "", (), "top", True
        if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
            tf = sh.text_frame
            text = tf.text.strip()
            sizes = tuple(
                (r.font.size or p.font.size).pt for p in tf.paragraphs for r in p.runs
                if (r.font.size or p.font.size) is not None) or (18.0,)
            a = tf.vertical_anchor
            anchor = "top" if a is None else str(a).split(".")[-1].split(" ")[0].lower()
            wrap = tf.word_wrap is not False
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            kind = "picture"
        elif sh.shape_type == MSO_SHAPE_TYPE.TEXT_BOX or (
                sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and text):
            kind = "text"
        else:
            kind = "shape"
        boxes.append(_Box(sh.name, _in(sh.left), _in(sh.top), _in(sh.width),
                          _in(sh.height), text, kind, sizes, anchor, wrap, gid, sh))
    return boxes, conns


# ------------------------------------------------------------------ text fit

def _needed_height(b: _Box) -> float:
    """Estimated height (in) the text needs at its width. Average glyph ≈ 0.52em."""
    lines = 0
    size = max(b.sizes) if b.sizes else 18.0
    usable = max(b.w - 0.1, 0.2)
    for para in b.text.split("\n"):
        chars_per_line = max(1, int(usable * 72 / (size * 0.52)))
        lines += max(1, -(-len(para) // chars_per_line)) if b.wrap else 1
    return lines * size * 1.2 / 72 + 0.1


def check_text_boxes(sn, boxes, out):
    for b in boxes:
        if b.kind != "text" or not b.text:
            continue
        need = _needed_height(b)
        slack = b.h - need
        if slack > TEXT_SLACK_IN and b.h > need * TEXT_SLACK_RATIO:
            sev = "error" if b.anchor in ("top", "bottom") else "warn"
            out.append(Finding(sev, sn, "Text box much taller than its text",
                               f"{b.label()} is {b.h:.2f}in tall, text needs ~{need:.2f}in "
                               f"(anchor {b.anchor}). Shrink the height to the text so it "
                               "centres and doesn't overlap neighbours."))
        # one short line in a very wide box (non-wrapping labels usually)
        if "\n" not in b.text and b.w > 1.5:
            size = max(b.sizes) if b.sizes else 18.0
            need_w = len(b.text) * size * 0.52 / 72 + 0.1
            if b.w > need_w * 3 and b.w - need_w > 1.0:
                bumps = [o for o in boxes if o is not b and o.kind in ("text", "picture")
                         and o.l < b.r and o.r > b.l and o.t < b.b and o.b > b.t]
                if bumps:
                    out.append(Finding("warn", sn, "Text box much wider than its text",
                                       f"{b.label()} is {b.w:.2f}in wide for ~{need_w:.2f}in "
                                       f"of text and overlaps {bumps[0].label()}."))


# ------------------------------------------------------------------ alignment

def check_alignment(sn, boxes, out):
    """Peers (same kind, similar size, same row / column) whose edges or centres
    almost — but not exactly — match. Reported once per outlier shape, not per
    pair, so one stray card in a row of five is one finding."""
    items = [b for b in boxes if b.kind in ("shape", "picture", "text") and b.w > 0.05 and b.h > 0.05]
    pairs = []   # (outlier candidates a, b, what, d)
    for i, a in enumerate(items):
        for b in items[i + 1:]:
            if a.kind != b.kind:
                continue
            same_row = abs(a.h - b.h) <= max(0.12, 0.25 * max(a.h, b.h)) and \
                abs(a.cy - b.cy) < max(a.h, b.h)
            same_col = abs(a.w - b.w) <= max(0.12, 0.25 * max(a.w, b.w)) and \
                abs(a.cx - b.cx) < max(a.w, b.w)
            checks = []
            if same_row:
                checks += [("top", a.t, b.t), ("vertical centre", a.cy, b.cy)]
            if same_col:
                checks += [("left", a.l, b.l), ("horizontal centre", a.cx, b.cx)]
            for what, va, vb in checks:
                d = abs(va - vb)
                if ALIGN_MIN < d <= ALIGN_SNAP:
                    pairs.append((a, b, what, d))
                    break
    # greedy: blame the shape involved in the most near-misses, report once
    while pairs:
        count = {}
        for a, b, what, d in pairs:
            for x in (a, b):
                count[id(x)] = count.get(id(x), 0) + 1
        worst = max(count, key=count.get)
        mine = [p for p in pairs if id(p[0]) == worst or id(p[1]) == worst]
        a, b, what, d = mine[0]
        shape = a if id(a) == worst else b
        others = [(p[1] if p[0] is shape else p[0]) for p in mine]
        out.append(Finding("warn", sn, "Near-miss alignment",
                           f"{shape.label()} is {d * 72:.1f}pt off the {what} of "
                           f"{len(others)} peer(s) e.g. {others[0].label()} — snap it."))
        pairs = [p for p in pairs if p[0] is not shape and p[1] is not shape]


# ------------------------------------------------------------------ connectors

def _seg_hits_box(x1, y1, x2, y2, b: _Box, pad=0.02) -> bool:
    """Liang–Barsky segment/rectangle intersection (box shrunk by pad)."""
    l, t, r, bt = b.l + pad, b.t + pad, b.r - pad, b.b - pad
    if r <= l or bt <= t:
        return False
    dx, dy = x2 - x1, y2 - y1
    p = [-dx, dx, -dy, dy]
    q = [x1 - l, r - x1, y1 - t, bt - y1]
    u1, u2 = 0.0, 1.0
    for pi, qi in zip(p, q):
        if abs(pi) < 1e-12:
            if qi < 0:
                return False
        else:
            u = qi / pi
            if pi < 0:
                u1 = max(u1, u)
            else:
                u2 = min(u2, u)
    return u1 < u2


def _dist_to_box(x, y, b: _Box) -> float:
    dx = max(b.l - x, 0, x - b.r)
    dy = max(b.t - y, 0, y - b.b)
    return (dx * dx + dy * dy) ** 0.5


def _contains(b: _Box, x, y, tol=END_TOL) -> bool:
    return b.l - tol <= x <= b.r + tol and b.t - tol <= y <= b.b + tol


def _seg_dist(px, py, x1, y1, x2, y2) -> float:
    dx, dy = x2 - x1, y2 - y1
    L = dx * dx + dy * dy
    u = 0.0 if L == 0 else max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / L))
    qx, qy = x1 + u * dx, y1 + u * dy
    return ((px - qx) ** 2 + (py - qy) ** 2) ** 0.5


def _conn_dist(c, x, y) -> float:
    return min(_seg_dist(x, y, *s) for s in c.segments())


def _is_own_label(c, b, conns) -> bool:
    """A small text box centred on this connector — and closer to it than to any
    other connector — is that connector's edge label, not an obstacle."""
    if b.w > 2.0 or b.h > 0.6:
        return False
    d = _conn_dist(c, b.cx, b.cy)
    if d > max(0.15, b.h):
        return False
    return all(_conn_dist(o, b.cx, b.cy) >= d - 1e-6 for o in conns if o is not c)


def _shares_end(a, b, tol=0.03) -> bool:
    pa = [(a.x1, a.y1), (a.x2, a.y2)]
    pb = [(b.x1, b.y1), (b.x2, b.y2)]
    return any(abs(x - u) < tol and abs(y - v) < tol for x, y in pa for u, v in pb)


def check_connectors(sn, boxes, conns, out):
    solid = [b for b in boxes if b.kind in ("text", "picture") or
             (b.kind == "shape" and b.w < 4 and b.h < 3)]   # skip big containers
    for c in conns:
        ends = [(c.x1, c.y1), (c.x2, c.y2)]
        attached = [b for b in boxes if any(_contains(b, x, y) for x, y in ends)]
        attached_groups = {b.group for b in attached if b.group is not None}
        # 1. passes through something it does not connect
        hits = []
        for b in solid:
            if b in attached or (b.group is not None and b.group in attached_groups):
                continue
            if b.kind == "text" and _is_own_label(c, b, conns):
                continue
            if any(_seg_hits_box(*s, b) for s in c.segments()):
                hits.append(b)
        if hits:
            sev = "error" if any(h.kind == "text" for h in hits) else "warn"
            out.append(Finding(sev, sn, "Arrow crosses a shape it doesn't connect",
                               f"{c.name!r} runs through " +
                               ", ".join(h.label() for h in hits[:3]) +
                               ". Re-route it (elbow / different side) or move the shape."))
        # 2. floating end — not near any shape
        for x, y in ends:
            near = min((_dist_to_box(x, y, b) for b in boxes if b.kind != "text"), default=99)
            if near > GAP_WARN:
                out.append(Finding("warn", sn, "Arrow end not attached",
                                   f"{c.name!r} ends at ({x:.2f}, {y:.2f})in, {near:.2f}in "
                                   "from the nearest shape — it stops short or overshoots."))
                break
    # 3. connectors drawn on top of each other
    for i, a in enumerate(conns):
        for b in conns[i + 1:]:
            # fan-in/fan-out and chained segments share a short stub by design
            min_run = 0.35 if _shares_end(a, b) else 0.15
            for (ax1, ay1, ax2, ay2) in a.segments():
                for (bx1, by1, bx2, by2) in b.segments():
                    if abs(ay1 - ay2) < 1e-3 and abs(by1 - by2) < 1e-3 and abs(ay1 - by1) < 0.04:
                        lo = max(min(ax1, ax2), min(bx1, bx2))
                        hi = min(max(ax1, ax2), max(bx1, bx2))
                        if hi - lo > min_run:
                            out.append(Finding("error", sn, "Arrows overlap",
                                               f"{a.name!r} and {b.name!r} share {hi - lo:.2f}in "
                                               "of the same horizontal line. Offset one."))
                            break
                    if abs(ax1 - ax2) < 1e-3 and abs(bx1 - bx2) < 1e-3 and abs(ax1 - bx1) < 0.04:
                        lo = max(min(ay1, ay2), min(by1, by2))
                        hi = min(max(ay1, ay2), max(by1, by2))
                        if hi - lo > min_run:
                            out.append(Finding("error", sn, "Arrows overlap",
                                               f"{a.name!r} and {b.name!r} share {hi - lo:.2f}in "
                                               "of the same vertical line. Offset one."))
                            break
                else:
                    continue
                break


# ------------------------------------------------------------------ entry points

# ------------------------------------------------------------------ auto-fix

def _plain_text_box(b: _Box) -> bool:
    """A text box with no visible fill/outline and no autofit, outside any group
    — shrinking it changes nothing visible, only the selection box."""
    sh = b.shape
    if sh is None or b.group is not None or b.kind != "text" or not b.text:
        return False
    xml = sh._element.xml
    if "fontScale" in xml:          # PowerPoint already shrank the text; leave it
        return False
    try:
        if sh.fill.type not in (None, 5):      # 5 = MSO_FILL.BACKGROUND
            return False
        if sh.line.fill.type not in (None, 5):
            return False
    except Exception:
        return False
    return True


def fix_slide(slide) -> int:
    """Safe, visually neutral repairs. Returns the number of shapes changed.

    1. Text boxes much taller than their text are shrunk to the text, keeping
       the anchored edge (top/bottom) or the centre, so the text doesn't move.
    2. Near-miss alignment (0.5-4pt) between peers is snapped to the value most
       peers share.
    """
    boxes, _ = _collect(slide)
    changed = 0
    for b in boxes:
        if not _plain_text_box(b):
            continue
        need = _needed_height(b)
        if b.h - need > TEXT_SLACK_IN and b.h > need * TEXT_SLACK_RATIO:
            pad = 1.25 if "normAutofit" in b.shape._element.xml else 1.1
            new_h = min(b.h, need * pad)
            if b.anchor == "bottom":
                top = b.b - new_h
            elif b.anchor in ("middle", "ctr", "center"):
                top = b.cy - new_h / 2
            else:
                top = b.t
            b.shape.top, b.shape.height = Emu(int(top * EMU_IN)), Emu(int(new_h * EMU_IN))
            b.t, b.h = top, new_h
            changed += 1
    # alignment: per dimension, cluster near-equal values and snap to the mode
    items = [b for b in boxes if b.shape is not None and b.group is None
             and b.kind in ("shape", "picture", "text") and b.w > 0.05 and b.h > 0.05]
    for dim in ("t", "l"):
        size = "h" if dim == "t" else "w"
        for i, a in enumerate(items):
            peers = [o for o in items if o is not a and o.kind == a.kind
                     and abs(getattr(o, size) - getattr(a, size)) <= 0.02
                     and ALIGN_MIN < abs(getattr(o, dim) - getattr(a, dim)) <= ALIGN_SNAP]
            if not peers:
                continue
            exact = [o for o in items if o is not a and o.kind == a.kind
                     and abs(getattr(o, size) - getattr(a, size)) <= 0.02
                     and abs(getattr(o, dim) - getattr(a, dim)) <= ALIGN_MIN]
            # move a only when more peers sit at the other value than with a
            target = getattr(peers[0], dim)
            agree = [o for o in peers if abs(getattr(o, dim) - target) <= ALIGN_MIN]
            if len(agree) > len(exact):
                if dim == "t":
                    a.shape.top = Emu(int(target * EMU_IN))
                else:
                    a.shape.left = Emu(int(target * EMU_IN))
                setattr(a, dim, target)
                changed += 1
    return changed


def fix_presentation(prs) -> int:
    """Run fix_slide over every slide of an open Presentation (call before save)."""
    return sum(fix_slide(s) for s in prs.slides)


def check_slide(slide, sn: int) -> list[Finding]:
    boxes, conns = _collect(slide)
    out: list[Finding] = []
    check_text_boxes(sn, boxes, out)
    check_alignment(sn, boxes, out)
    check_connectors(sn, boxes, conns, out)
    return out


def check_pptx(path) -> list[Finding]:
    prs = Presentation(str(path))
    out: list[Finding] = []
    for n, s in enumerate(prs.slides, 1):
        out += check_slide(s, n)
    return out


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print("usage: geomcheck [--fix] <deck.pptx> [...]", file=sys.stderr)
        return 2
    if argv[0] == "--fix":
        for p in argv[1:]:
            prs = Presentation(str(p))
            n = fix_presentation(prs)
            prs.save(str(p))
            print(f"fixed {n} shape(s) in {p}")
        argv = argv[1:]
    errors = 0
    for p in argv:
        fs = check_pptx(p)
        errors += sum(f.severity == "error" for f in fs)
        print(f"== {p}: {sum(f.severity == 'error' for f in fs)} error(s), "
              f"{sum(f.severity == 'warn' for f in fs)} warning(s)")
        for f in fs:
            print(f"  {f}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
