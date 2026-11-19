# Shared compose kit & style presets — cross-repo contract

This file is **identical** in Arch Gen (`docs/COMPOSE_API.md`) and PPT Gen
(`docs/COMPOSE_API.md`). Arch Gen and PPT Gen stay separate repos, but these
shared pieces are kept byte-identical so the repos can be merged later by
deleting one copy and fixing imports.

| Shared piece | Arch Gen | PPT Gen |
|---|---|---|
| Drawing kit | `archgen/shared/compose.py` | `pptgen/shared/compose.py` |
| Geometry QA | `archgen/shared/geomcheck.py` | `pptgen/shared/geomcheck.py` |
| Style presets | `styles/*.json` | `styles/*.json` |
| This contract | `docs/COMPOSE_API.md` | `docs/COMPOSE_API.md` |

A parity test in each repo (`tests/test_shared_parity.py`) compares the copies
when the sibling repo is checked out next to it, and fails on any drift. **Edit
one copy, then copy it over the other.**

## Rules for shared files
- `compose.py` imports **only** python-pptx and the stdlib — never `archgen` or `pptgen`.
- All coordinates and sizes are **inches**.
- Colours/fonts/shape language come in via `Style`; icons via an `IconSource`
  (`.png(key, size_in, tone="auto"|"inverse") -> path | None`).
- Repo-specific glue lives outside the shared file:
  `archgen/kit.py` (`RegistryIcons`, `style(preset)`) and
  `pptgen/kit.py` (`AssetIcons`, `style_from_theme`, `add_cover`) +
  `pptgen/brandkit.py` (brand mode).

## API (inches)
| Function | Purpose |
|---|---|
| `Style.from_tokens(dict)` | Style from a preset / theme.json dict |
| `set_background(slide, hex)` | Solid slide ground |
| `rrect(slide, x,y,w,h, fill=, line=, line_pt=, radius=)` | Rounded/sharp rectangle |
| `text(slide, x,y,w,h, paras \| str, font=, size=, color=, ...)` | Non-growing text box |
| `zone(slide, st, x,y,w,h, title)` | Labelled container |
| `band(slide, st, x,y,w,h, title)` | Full-width strip |
| `connector(slide, x1,y1,x2,y2, color=, dashed=, elbow=, head=, tail=)` | Arrow (tail=True for bidirectional) |
| `callout(slide, st, cx,cy, n)` | Numbered marker |
| `badge(slide, st, x,y, label)` | Pill label |
| `icon_card(slide, st, icons, key, label, x,y,w,h, sublabel=)` | Icon + label as **one group** |
| `text_height(paras, w, size)` | Height text needs — `text(..., h=None)` uses it so boxes hug text |
| `snap(v)` | Snap to a 1/32in grid so peers share identical coordinates |
| `row(n, x, w, item_w)` / `column(n, y, h, item_h)` | Evenly spaced, snapped positions for a row/column of peers |
| `port(shape, side, offset)` | Exact edge midpoint of a shape |
| `link(slide, a, b, a_side=, b_side=, lane=)` | Arrow attached to shape edges; **glued** (moves with the shapes) when `lane=0`; `lane=±0.2` offsets parallel arrows so they never stack |
| `cover(slide, st, icons, heading=, subheading=, logos=, visual=, visual_keys=, footer=, layout=, motif_png=)` | The cover slide |

Cover visuals (`COVER_VISUALS`): `hub-spoke`, `ring`, `icon-cluster`, `stack`, `flow-strip`.
Cover layouts (`COVER_LAYOUTS`): `split-right-visual`, `left-text-right-visual`,
`centered-visual`, `color-block-left` (default comes from the preset's `style.cover`).

## The cover rule (both repos)
A cover holds **only**: one heading (≤ ~60 chars, ≤ 2 lines), one one-line
subheading (≤ ~90 chars), logos (≤ 3), one visual (infographic / icon cluster, or a
brand motif in brand mode), and optionally one small footer line. No bullets,
agenda, stats or paragraphs — those go to slide 2. PPT Gen's `check.py` enforces it.

## Geometry QA (`geomcheck`)
Runs on any `.pptx` — `archgen qa out.pptx`, `make qa OUT=...` (Arch Gen), and inside
`make check` (PPT Gen). It flags:

| Finding | Severity | Fix |
|---|---|---|
| Near-miss alignment — peers in a row/column off by 0.5–4pt (reported once per outlier) | warn | `row()` / `column()` / `snap()` |
| Text box much taller than its text (top/bottom anchored) | error | `text(..., h=None)` |
| Text box much wider than its text **and** overlapping something | warn | narrow it |
| Arrow runs through text/a card it does not connect (its own edge label is ignored) | error (text) / warn | re-route, `link(..., a_side=, b_side=)` |
| Arrows overlap — two connectors share a run of the same line | error | `link(..., lane=±0.2)` |
| Arrow end not attached — stops short of / overshoots its shape | warn | `link()` instead of raw coordinates |

Heuristic: text width is estimated (~0.52em/char), so treat warnings as "look here".

## Style presets
Eight generic presets, shared format (`styles/<name>.json`): palette (all tokens
contrast-checked), `font` (`primary`/`heading`/`body`/`mono`), `type_scale`,
`space`, and `style` (`corner_radius`, `line_pt`, `card`, `divider`, `cover`,
`decoration`, `density`, `icon_style`).

| Preset | Mode | Best for |
|---|---|---|
| corporate-clean | light | client proposals, status reports, general business |
| modern-tech | dark | platform, AI, engineering and product decks |
| editorial | light | strategy narratives, thought leadership |
| warm-minimal | light | people, culture, sustainability, workshops |
| bold-contrast | light | pitches, launches, keynote-style statements |
| consulting-classic | light | analysis-heavy decks, steering committees |
| data-dark | dark | dashboards, KPI reviews, analytics readouts |
| pastel-soft | light | education, onboarding, internal comms |

**Generic is the default. Brand mode is opt-in** — only when the user explicitly
asks for a brand (PPT Gen `brands/<name>/brand.yaml`).
