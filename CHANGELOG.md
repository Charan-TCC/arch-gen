# Changelog

All notable changes to archgen. Format loosely follows
[Keep a Changelog](https://keepachangelog.com/); this project uses semantic versioning.

## [Unreleased]

### Fixed
- **Group titles no longer overlap the first node.** `build_dot` pre-wraps cluster labels
  at the narrowest possible cluster width and reserves headroom at a fontsize that still
  clears the drawn title at low layout scales; the drawn title font now scales down with
  the layout (floor 9pt) and its textbox hugs the wrapped text.
- **Content no longer hides behind legends.** The bottom legend reserve is computed from
  the legends that actually render (abbreviation panel rows included) instead of a fixed
  band; abbreviation lists past 6 entries render in two columns, and the connections
  panel centers in the free gap between the other panels.
- **Node label boxes hug their text** (width and height, capped inside the card) instead
  of spanning the full card width — what you grab in PowerPoint now matches the label.
  Node label fonts scale down with the layout (floor 8pt) so labels stop breaking
  mid-word in scale-shrunk cards.
- **Edge labels avoid nodes.** Flow edge labels walk the elbow route for a collision-free
  spot (vs. node cards and other edge labels) and shove sideways past whatever card the
  midpoint hits when the connector runs straight through a node stack.
- **Layered TB sizing** used LR gap math (nodesep/ranksep swapped); tiers-as-rows now
  budget margins, title bands, and ranksep correctly, and `--auto-direction` now covers
  layered diagrams (card size recomputed per orientation).

### Added
- **Layered flow mode** (`layered: true`) — renders groups as clean parallel tier bands
  (columns in LR, rows in TB) instead of staggering by edge depth. The renderer owns all
  sizing: it counts nodes per tier and fits the *densest* tier to the slide at a legibility
  floor, so many-node tiers (e.g. a six-source rail) stay readable without breaking labels
  mid-word. The spec stays purely declarative — no coordinates or sizes. Flow + Graphviz
  only; the pure-Python fallback degrades to the ordinary flow layout. See
  `examples/layered_data_platform.yaml`. Flow card icon/label placement now scales with the
  card in this mode (icon = fraction of height, label font clamped to a readable floor).
- **New example specs** — small, realistic, agent-origin (prompt recorded in each header):
  `serverless_api.yaml` (AWS serverless), `event_driven_pipeline.yaml` (Kafka pipeline),
  `rag_chatbot.yaml` (RAG/LLM app). Each renders clean with color icons and elbow connectors.
- **Full-color generic icons** — 28 architecture concepts (`generic.user`, `generic.database`,
  `generic.loadbalancer`, `generic.cache`, `generic.queue`, `generic.function`, …) sourced from
  Icons8 **Flat Color Icons** (MIT), replacing the flat-gray glyph placeholders.
- **True-color technology logos** — `tech.*` icons resolve in order **Gil Barbara "logos" (CC0)
  → Devicon (MIT) → Simple Icons** (monochrome, tinted to the official brand hex as a fallback).
  All served via the Iconify API. `archgen icons add` is now color-first; `apply_color_logos()`
  upgrades the whole cached set in one pass.
- **Native PowerPoint elbow connectors** for flow-diagram edges — real, editable right-angle
  connectors chosen from pure geometry (dominant axis → facing sides), so any graph routes cleanly
  instead of cutting diagonally across the slide.

### Fixed
- Icons no longer sit on a white box: the QuickLook raster path now flood-fills the opaque white
  thumbnail canvas to transparent from the borders (preserving white *inside* a logo).
- Node-label overlaps in dense architecture-board sections — the grid now picks its column count for
  a minimum ~90pt cell so adjacent labels don't collide.
- Edge labels no longer wrap into stacked characters (internal text-box margins were eating the width)
  and are clamped on-slide.
- The network connection legend no longer overlaps the first container (a top band is reserved for it).

### Changed
- `tech.*` icons render in their brand color instead of a flat gray; manifest entries carry
  `brand_hex` and a `source` (`logos`/`devicon`/`simple-icons`) that drives `recolorable`.
- Official AWS 48px architecture-service icons, nested Azure service icons, and GCP
  service icons are now auto-discovered into the registry with deterministic keys.
- Existing AWS/Azure/GCP short keys now prefer official assets through stable aliases
  and fall back to placeholders only when the corresponding pack is absent.
- `make icons` rebuilds from local assets without making network requests.
- SVG conversion now falls through to the next available engine when an installed
  vector or raster converter rejects a particular official icon.

### Added
- `kind: network` diagrams with deterministic role-based placement for cloud boundaries,
  virtual networks, ingress, compute, management, private-access, DNS, and managed services.
- Semantic node roles, group kinds/roles, and edge kinds while keeping all coordinates and
  visual decisions inside Archgen.
- Native orthogonal connection rendering and an automatically derived connection legend for
  HTTPS, diagnostics, DNS, private connections, and virtual-network links.
- A flat, tall architecture-board network profile with explicit Manhattan routes and inherited
  Office shadow effects disabled for clean reference-style diagrams.
- Reusable architecture-board zones for web, storage, private networks, data, discovery,
  cost, deployment, and integration sections, plus dense icon-grid placement.
- QuickLook raster cleanup now trims and refits oversized white thumbnail canvases so official
  AWS 48px icons render at full size in PowerPoint.
- A complete private-AKS example at `examples/azure_private_aks.yaml` and regression coverage
  for its layout and connection semantics.

## [0.1.0] — 2026-07-16

Initial working MVP: spec → editable `.pptx`, end to end.

### Added
- **Spec** — pydantic models (`Diagram`/`Node`/`Group`/`Edge`), fail-loud linter,
  JSON-Schema export, YAML loader.
- **Icon registry** — manifest + resolver; fetcher for CC0 tech logos (simple-icons)
  and generated placeholder cloud icons; official GCP pack wired via `REAL_FILE`.
- **SVG→picture converter** — Inkscape/LibreOffice (EMF) with macOS QuickLook / cairosvg
  (PNG) fallbacks, content-hash cache, fixed-size normalization, CC0-only recolor.
- **Layout** — Graphviz `dot` adapter (`-Tjson`) + group-aware pure-Python fallback that
  guarantees non-overlapping containers; geometry transform (points→EMU, Y-flip, fit).
- **Renderer** — deterministic; tinted grouped containers, node cards with uniform-scaled
  icon pictures, auto-fit contained labels, edge connectors with arrowheads + dash styles.
- **Legends** — auto color + abbreviation legends derived from spec content.
- **Agent** — system prompt (schema + icon catalog + rules) and prompt→spec→lint→retry
  loop; optional Anthropic backend.
- **CLI** — `render`, `lint`, `prompt`, `generate`, `icons`, `schema`, `doctor`.
- **Tests** — 27 tests: linter, geometry, non-overlapping layout, golden render, licensing.
- **Docs** — README, SPEC.md, docs/ARCHITECTURE.md, docs/SPEC_FORMAT.md, CONTRIBUTING.md.

### Known limitations
- Full official AWS/Azure packs not yet ingested (placeholders until dropped in + mapped).
- Icons embed as PNG unless Inkscape/LibreOffice is installed (then vector EMF).
- Edges are straight border-to-border (Graphviz spline points parsed but not yet drawn).
- Single-slide output; no IaC→spec import. See SPEC.md §9 Roadmap.
