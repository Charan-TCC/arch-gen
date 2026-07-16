# CLAUDE.md — how to work in this repo (read before generating any diagram)

archgen turns a declarative spec into a **native, editable** PowerPoint architecture
diagram. The agent writes only the spec; a deterministic Python renderer owns the
aesthetics. Follow [DESIGN_GUIDELINES.md](DESIGN_GUIDELINES.md), [SPEC.md](SPEC.md),
and [docs/SPEC_FORMAT.md](docs/SPEC_FORMAT.md).

## 0. Before you iterate — read this, it will save you an hour

The expensive mistake is **iterating on rendered output by trial-and-error** instead
of understanding constraints first. Do these three things up front:

1. **Confirm the fidelity bar.** If the user gave a reference image, ask ONE question:
   *must it match that image closely, or is a clean archgen-style diagram fine?* This
   single answer picks the whole path (see §3). Don't assume.
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

## 2. Layout engines & their constraints

- `kind: flow` → **graphviz** (`dot`) when present: clustered ranks, crossing
  minimization. Best general choice. Pure-Python `simple_layout.py` is the fallback
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
archgen has no template for those — `build_shine.py` is the worked example.

## 4. Visual QA — DO NOT re-render blindly to "check"

**MANDATORY before closing ANY diagram task — the agent MUST look at the output as an
image, not just trust that the pipeline exited 0.** The loop is not optional:

1. Render the spec to a `.pptx`, then rasterize it to PNG (`make preview SPEC=...`, or
   `qlmanage -t -s 1600 -o <dir> <file.pptx>`).
2. **Actually open and view the PNG** (Read the image file). Check for: labels clipping
   or wrapping mid-word, icons or text spilling outside their card/container, overlapping
   shapes, collided connectors, empty/placeholder icons, off-slide elements.
3. If ANYTHING is wrong or "just doesn't look good" → go back, fix the spec (or the
   renderer if it's a rendering bug), re-render, and **view again**. Repeat until the
   image is clean. Never hand back or close a task on an unviewed or flawed render.
4. Only once the rendered image looks correct do you close the task.

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

`build_shine.py` (repo root) is the scaffold: helpers `card`, `zone`, `connector`,
`icon`, `text`. A new bespoke diagram = new content list + coordinates, same helpers.
Regenerate its PNG via `make gallery`.

## 6. Gotcha already fixed — don't reintroduce

`graphviz_layout.build_dot` once emitted `subgraph cluster_"gid" {` (quotes inside the
name) → `dot` syntax error. It was dormant because `dot` used to be absent. Fixed to
`subgraph "cluster_<gid>" {`. If you touch DOT generation, keep the whole subgraph name
quoted and keep the `cluster_` prefix (the JSON parser matches `name.startswith("cluster_")`).
Run `make test` — the graphviz path is now exercised.

## 7. Rendering invariants — keep these true (renderer AND bespoke builds)

These are hard rules for every node/card drawn, in `archgen/render/renderer.py` and in
any `custom_builds/*.py` bespoke composition. Breaking one is a visual bug caught by §4.

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
