# CLAUDE.md — how to work in this repo (read before generating any diagram)

archgen turns a declarative spec into a **native, editable** PowerPoint architecture
diagram. The agent writes only the spec; a deterministic Python renderer owns the
aesthetics. Follow [DESIGN_GUIDELINES.md](docs/DESIGN_GUIDELINES.md), [SPEC.md](docs/SPEC.md),
and [docs/SPEC_FORMAT.md](docs/SPEC_FORMAT.md).

## 0. Before you iterate — read this, it will save you an hour

The expensive mistake is **iterating on rendered output by trial-and-error** instead
of understanding constraints first. Do these three things up front:

1. **Ask the user TWO questions before writing any spec — every time, no assuming:**
   - **Page format:** *formal presentation (16:9, fixed 13.33"×7.5" slide) or free-form
     (slide grows so text never shrinks below readable size)?* This is the spec's
     required `page` field; lint refuses a flow spec without it. On `"16:9"` the
     content must be reshaped until it fits readably — the slide will not grow.
   - **Fidelity bar** (only if the user gave a reference image): *must it match that
     image closely, or is a clean archgen-style diagram fine?* This answer picks the
     whole path (see §3).
2. **Check capabilities:** run `make doctor`. It reports the layout engine and icon
   converter actually available in THIS environment.
3. **Know the layout constraints** (§2) before choosing a `direction`/`kind`. The
   geometry transform (`archgen/layout/geometry.py`) scales the whole graph
   **uniformly** to fit the slide — tall node stacks get shrunk until labels collapse.

## 1. Environment facts (this machine)

- Custom Homebrew prefix: `/Users/charan/homebrew`.
- **Graphviz `dot` IS installed** (`/Users/charan/homebrew/bin/dot`) → archgen uses its
  real graphviz layout for `kind: flow`. `make doctor` should say `layout engine: graphviz`.
  If it ever says `simple (pure-Python fallback)`, `dot` fell off PATH — fix that first;
  the fallback is much weaker (see §2).
- Icon converter is **qlmanage** (macOS QuickLook → PNG). No Inkscape/LibreOffice, so
  no vector EMF and no `soffice`. Icons embed as high-res PNG. This is fine.
- QuickLook substitutes the Windows font "Segoe UI" with a wider fallback, so **PNG
  previews wrap text slightly more than real PowerPoint** — judge layout, not exact wrap.
- **`out/` is archgen's scratch output dir** — the default target for ad-hoc
  renders (`OUT ?= out.pptx` in the Makefile writes there when you pass a bare
  filename with a directory). It's gitignored and `make clean` removes it freely;
  never treat anything in it as precious. Every spec render AND every bespoke
  build writes there — nothing is ever saved to the repo root.
- **Folder map:** `archgen/` package (`archgen/shared/` = modules byte-identical
  with PPT Gen) · `examples/` specs · `bespoke/<name>/` one folder per bespoke
  diagram (`build.py`, optional `spec.yaml`, `reference/` images) · `docs/`
  (guides + `gallery/`) · `styles/` presets · `archive/` superseded files and
  `archive/outputs-legacy/` (old renders with no reproducible source — gitignored,
  don't delete without checking).

## 2. Layout engines & their constraints

- `kind: flow` → **graphviz** (`dot`) when present: clustered ranks, crossing
  minimization. Best general choice. Edges use `splines=ortho` and the renderer
  draws graphviz's own routes (they go around nodes) — don't reintroduce
  renderer-invented elbows. Route ends snap to the visible icon+label group;
  ≤1-bend routes are glued native connectors, longer ones are freeforms.
  Edge labels sit on the route (or beside short arrows), never over titles.
  A long dependency chain (5+ ranks) makes a tall, sparse layout — use
  `layered: true` (one band per group) and show a mirrored region ONCE. Pure-Python `simple_layout.py` is the fallback
  ONLY when `dot` is absent — and it has sharp edges:
  - Cross-group edges must form a **DAG**. A group-level cycle (e.g. A→B and B→A across
    groups) makes longest-path layering diverge → collapsed/overlapping containers.
  - It uniform-scales to fit: a column of 5–6 stacked nodes is ~10in tall → scaled to
    ~0.58 → node label boxes collapse and text clips. Prefer `direction: TB` (wide
    horizontal bands) over `LR` (tall narrow columns); keep labels short (≤~8 chars/word).
- `kind: network` → deterministic semantic zones (`network_layout.py`). Two modes:
  the 2×2 role grid, and the **architecture board** (fixed role boxes: web/storage/
  private_network/data/discovery/cost/deployment/integration) — great for dense grouped
  boards, but the box positions are **hard-tuned to one AWS topology**; don't expect an
  arbitrary zone count to map cleanly.

## 3. Choosing an approach for a NEW / different architecture

```
reference image the user wants matched closely?
├─ NO  → archgen spec. flow(graphviz) for pipelines; network for zonal boards.
│        Keep it in-tool: authentic, spec-driven, minimal code.
└─ YES → can an archgen profile (flow / network board) actually express that shape?
         ├─ YES → archgen spec, then nudge in PowerPoint if needed.
         └─ NO  → BESPOKE build with python-pptx (see §5). Do NOT fight archgen's
                  generic layouts to force a hand-designed composition.
```

Signs you need the bespoke path: 3+ side-by-side zones **plus** a full-width band,
header/annotation arrows, numbered callouts, or left-aligned icon+caption cards.
archgen has no template for those — `bespoke/azure_data_platform/build.py`
is the worked example.

## 4. Visual QA — text report first, PNG last (token budget is part of the job)

Iterating by rendering PNGs and eyeballing them is the expensive failure mode: every
image view costs far more tokens than a text report, and visual diagnosis re-discovers
facts the tool already knows. The loop below is ordered by cost — exhaust each stage
before escalating to the next:

1. **Lint** (`make lint SPEC=...`) — spec-level problems.
2. **Render with the layout + geometry report**:
   `archgen render <spec> -o <out> --auto-direction --strict`.
   `--auto-direction` lays out both TB and LR and keeps the better fit (a decision the
   tool makes deterministically — never burn an iteration discovering it yourself).
   The printed **layout report** flags scale collapse, cramped cards, words that will
   break mid-word, and unbalanced slides — with the concrete fix for each. Iterate on
   the SPEC against this text report until it has **no error-level findings**
   (`--strict` exits 3 while any remain). Do NOT rasterize or view a PNG during this
   stage. **Budget: 3 spec-fix iterations** — if errors persist after 3, the shape is
   wrong for this profile: change `kind`/`layered` or switch path per §3; don't grind.
   The render also prints a **geometry report** (`archgen/shared/geomcheck.py`): near-miss
   alignment, text boxes taller than their text, arrows crossing labels/cards or stacked on
   each other. Bespoke builds: run `archgen qa <out.pptx>` (or `make qa OUT=...`) and fix
   errors before looking at a PNG. Build with `compose.row/snap`, `text(h=None)` and
   `compose.link(..., lane=)` so these don't happen in the first place.
3. **One mandatory PNG view** — the report can't see everything (icon quality,
   connector collisions, overall balance). Rasterize (`make preview SPEC=...` or
   `qlmanage -t -s 1600 -o <dir> <file.pptx>`) and **actually view the PNG** (Read the
   image). Check: labels clipping/wrapping mid-word, text spilling outside containers,
   overlapping shapes, collided connectors, empty/placeholder icons, off-slide elements.
4. If something is visually wrong → fix, re-render (report must stay clean), view again.
   **Budget: 2 fix-and-view rounds after the first look** (3 PNG views total). Still
   flawed after that? The layout profile can't express the composition — switch path
   (§3, usually bespoke) or surface the tradeoff to the user. Never close a task on an
   unviewed or flawed render — but never chase pixel perfection past the budget either.

- Committed reference renders live in **[docs/gallery/](docs/gallery/)** — look there
  first instead of re-running the pipeline to remember what a spec produces.
- To preview one spec: `make preview SPEC=... OUT=...` (renders + opens a QuickLook PNG).
- After ANY change to renderer/layout/theme/icons or a new example: `make gallery`
  (regenerates every PNG in one pass), then eyeball the images. That is the required
  visual QA — one command, all examples, no repeated one-off renders.
- Always lint before render: `make lint SPEC=...`.
- Remember the §1 caveat: qlmanage substitutes a wider font, so PNG previews wrap text
  slightly more than real PowerPoint. Judge layout and overflow, not exact wrap points.

## 5. Bespoke image-faithful builds (python-pptx + archgen icons)

When §3 routes to bespoke, reuse archgen's real icon assets so branding stays authentic
and output stays native/editable:

```python
from archgen.registry.resolver import IconResolver
from archgen.registry import emf
e = IconResolver().get("azure.databricks")
recolor = f"#{e.brand_hex}" if e.recolorable and e.brand_hex else None  # brand icons: never recolor
png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)   # returns a PNG path
# add_picture(png, ...) in python-pptx; keep everything as native shapes/text/connectors.
```

**Use the shared kit, don't copy helpers.** `archgen/shared/compose.py` (identical to PPT Gen's
`pptgen/compose.py`, contract in [docs/COMPOSE_API.md](docs/COMPOSE_API.md)) has `rrect`,
`text`, `zone`, `band`, `connector` (incl. bidirectional), `callout`, `badge`, `icon_card`
(icon + label as ONE group — §7) and `cover`. Glue: `archgen/kit.py`:

```python
from archgen.shared import compose
from archgen.kit import RegistryIcons, style
st, icons = style("corporate-clean"), RegistryIcons()   # any preset from styles/
compose.icon_card(slide, st, icons, "azure.databricks", "Databricks", x, y, 1.6, 0.9)
```

A new bespoke diagram = content list + coordinates on top of this kit. The older
`bespoke/*/build.py` still carry their own copied helpers — migrate them onto the kit when
you touch them. Never edit only one copy of `compose.py`: copy it to the sibling repo
(the parity test fails otherwise).

**Styles:** `styles/*.json` are the same 8 presets as PPT Gen. Spec `theme: <preset>` or
`archgen render --theme <preset>`. Provider icons keep official colours.

## 6. Gotcha already fixed — don't reintroduce

`graphviz_layout.build_dot` once emitted `subgraph cluster_"gid" {` (quotes inside the
name) → `dot` syntax error. It was dormant because `dot` used to be absent. Fixed to
`subgraph "cluster_<gid>" {`. If you touch DOT generation, keep the whole subgraph name
quoted and keep the `cluster_` prefix (the JSON parser matches `name.startswith("cluster_")`).
Run `make test` — the graphviz path is now exercised.

## 7. Rendering invariants — keep these true (renderer AND bespoke builds)

These are hard rules for every node/card drawn, in `archgen/render/renderer.py` and in
any `bespoke/*/build.py` bespoke composition. Breaking one is a visual bug caught by §4.

- **Group icon + text into ONE shape.** Every node is an icon *and* its label bundled in
  a single PowerPoint group (`slide.shapes.add_group_shape()` → add icon + label into the
  group). They must move, align, and stay together as a unit — never leave a label as a
  loose textbox drifting from its icon. See `_draw_nodes` / `_place_icon` / `_place_label`.
- **Text boxes hug their text.** A label/text box is sized to its content, not oversized:
  - Width fits the text and is **capped** (`node.label_max_width_in`) so a label never
    bleeds into a neighbouring node; long labels wrap to ≤2 lines *inside* that cap.
  - The box stays **inside its card** — the icon slot is card-relative (`_icon_slot`) so
    the label never spills past the card bottom into the container below.
  - Do **NOT** use `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` for height: it grows the box past the
    card border and collides with the container. Keep `auto_size = NONE` with a clamped,
    card-bounded height.
  - Keep labels short at the source (the linter caps at `MAX_LABEL_LEN`); hugging a box is
    no substitute for an over-long label — abbreviate instead (declare it in
    `abbreviations` so the legend covers it).
