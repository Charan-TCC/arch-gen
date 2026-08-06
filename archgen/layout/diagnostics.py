"""Post-layout diagnostics: deterministic layout-quality findings.

Token economy for agent-driven use: the agent iterates on this TEXT report
(scale collapse, cramped cards, wrap risk, direction advice) instead of
rasterizing and visually inspecting a PNG after every render. The single
final PNG view stays mandatory (CLAUDE.md §4); everything detectable here
must be fixed *before* that one look.

All checks are geometric heuristics computed from the Layout and the fit
scale — no rendering, no rasterizing. The wrap/width thresholds assume the
default theme font; they are advisory (warn), while scale collapse is the
one hard failure (error) because it reliably produces unreadable output.
"""

from __future__ import annotations

from dataclasses import dataclass

from archgen.layout.graphviz_layout import Layout
from archgen.spec.models import Diagram

# Thresholds calibrated against the committed gallery renders (2026-07):
# 44% scale shows mid-word label breaks (rag_chatbot); 57% is readable with
# wasted whitespace (serverless_api); 25-35% is unusable (collapse).
COLLAPSE_SCALE = 0.50  # below: labels reliably break -> error
MIN_READABLE_SCALE = 0.75  # below: cramped, whitespace-heavy -> warn
# Drawn card width under which a typical two-word label cannot fit.
MIN_CARD_WIDTH_IN = 0.70
# Rough average glyph advance for the default theme font at label size.
CHAR_WIDTH_IN = 0.082
# Fraction of the slide's off-axis left empty before we call it unbalanced.
LOW_FILL_RATIO = 0.55

POINTS_PER_INCH = 72.0
EMU_PER_INCH = 914400


@dataclass(frozen=True)
class Finding:
    severity: str  # "error" | "warn" | "info"
    message: str

    def __str__(self) -> str:  # pragma: no cover - formatting only
        return f"[{self.severity}] {self.message}"


def has_errors(findings: list[Finding]) -> bool:
    return any(f.severity == "error" for f in findings)


def diagnose(
    diagram: Diagram,
    layout: Layout,
    *,
    scale: float,
    avail_w_emu: int,
    avail_h_emu: int,
    label_cap_in: float = 1.4,
    direction_flipped: bool = False,
    auto_direction_tried: bool = False,
    quality: str = "draft",
) -> list[Finding]:
    """Return layout-quality findings for a *flow* layout (network boards are
    hand-tuned and skip this).

    ``quality`` picks the readability bar. In ``"draft"`` (default) a cramped
    layout is a *warning* — the tool still emits a usable working diagram. In
    ``"release"`` the same conditions are *errors*: a formal presentation slide
    scaled below the readable floor (or with hard-wrapping labels) is not
    shippable, so ``--strict`` should refuse it rather than pass unreadable
    output. Release never *lowers* a severity; it only promotes readability
    warnings to errors."""
    findings: list[Finding] = []
    release = quality == "release"

    if direction_flipped:
        findings.append(
            Finding(
                "info",
                f"auto-direction: rendered {diagram.direction} — it fits better than "
                "the spec's declared direction (update the spec to silence this).",
            )
        )

    if scale < MIN_READABLE_SCALE:
        flip = "LR" if diagram.direction == "TB" else "TB"
        # Once auto-direction has already compared both orientations, advising
        # the flip again would send the agent in a circle — advise reshaping.
        # Layered mode owns direction and banding, so neither hint applies there.
        if diagram.layered:
            direction_hint = "rebalance the tiers: "
        elif auto_direction_tried:
            direction_hint = "both directions were tried — reshape the spec: "
        else:
            direction_hint = f"try direction {flip} (or `--auto-direction`), then "
        collapsed = scale < COLLAPSE_SCALE
        # Below the collapse floor is always an error; between collapse and the
        # readable floor is a warning in draft but an error in release (a formal
        # slide with sub-readable text is not shippable).
        severity = "error" if (collapsed or release) else "warn"
        effect = (
            "labels will collapse and wrap mid-word"
            if collapsed
            else "cards are cramped and much of the slide is empty"
        )
        reshape = (
            "spread nodes more evenly across groups or merge sparse tiers."
            if diagram.layered
            else "split the longest node chain across groups, or set `layered: true` for banded tiers."
        )
        if diagram.page == "16:9":
            reshape += (
                ' The page is "16:9" (locked deck slide) — the slide will NOT grow; '
                "only reshaping fixes this (or the user opts into page: \"free\")."
            )
        findings.append(
            Finding(
                severity,
                f"layout is scaled to {scale:.0%} of its natural size — {effect}. "
                f"Fix the shape, not the render: {direction_hint}{reshape}",
            )
        )

    # Cards drawn too narrow for a label even at healthy scale.
    narrow = [
        (n.id, n.w * scale / POINTS_PER_INCH)
        for n in layout.nodes.values()
        if n.w * scale / POINTS_PER_INCH < MIN_CARD_WIDTH_IN
    ]
    if narrow:
        worst_id, worst_w = min(narrow, key=lambda t: t[1])
        findings.append(
            Finding(
                "error" if release else "warn",
                f"{len(narrow)} node card(s) drawn under {MIN_CARD_WIDTH_IN}\" wide "
                f"(worst: '{worst_id}' at {worst_w:.2f}\") — labels will wrap hard.",
            )
        )

    # Words that cannot fit the label box cap break mid-word regardless of scale.
    label_w_in = {
        n.id: max(layout.nodes[n.id].w * scale / POINTS_PER_INCH, 0.0)
        for n in diagram.nodes
        if n.id in layout.nodes
    }
    for n in diagram.nodes:
        if n.id not in label_w_in:
            continue
        cap = min(max(label_w_in[n.id], MIN_CARD_WIDTH_IN), label_cap_in)
        longest = max((w for w in n.label.split()), key=len, default="")
        if len(longest) * CHAR_WIDTH_IN > cap:
            findings.append(
                Finding(
                    "error" if release else "warn",
                    f"node '{n.id}': word '{longest}' is wider than its label box "
                    f"(~{cap:.2f}\") and will break mid-word — abbreviate it "
                    f"(declare in `abbreviations`).",
                )
            )

    # Unbalanced slide: uniform scaling fills one axis; report when the other
    # axis is mostly empty so the agent knows the composition looks sparse.
    x0, y0, x1, y1 = layout.graph_bb
    drawn_w = (x1 - x0) * scale * EMU_PER_INCH / POINTS_PER_INCH
    drawn_h = (y1 - y0) * scale * EMU_PER_INCH / POINTS_PER_INCH
    fill_w = drawn_w / avail_w_emu if avail_w_emu else 1.0
    fill_h = drawn_h / avail_h_emu if avail_h_emu else 1.0
    low_axis = min(fill_w, fill_h)
    if low_axis < LOW_FILL_RATIO:
        axis = "width" if fill_w < fill_h else "height"
        findings.append(
            Finding(
                "info",
                f"diagram uses only {low_axis:.0%} of the slide's {axis} — layout is "
                f"unbalanced; consider the other direction or regrouping nodes.",
            )
        )

    return findings
