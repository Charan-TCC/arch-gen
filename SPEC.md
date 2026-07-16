# archgen — Project Specification

**Status:** v0.1.0 (working MVP) · **Owner:** charan · **Last updated:** 2026-07-16

---

## 1. Problem & goal

Architecture diagrams are built by hand today: hunting logos on the Azure/AWS/GCP
icon sites and misc. logo sites, then placing everything manually in PowerPoint. The
result is inconsistent spacing, sizing, colors, and arrows, and it is slow.

**Goal:** a pipeline where an **agent (Codex/Claude) writes only a declarative spec**,
and a **deterministic renderer** turns any valid spec into a **neat, editable `.pptx`
slide** — consistent sizing/spacing/colors/backgrounds, grouped containers, tidy
arrows, auto-fit labels (no stretched text boxes), and auto-generated color +
abbreviation legends.

### 1.1 Load-bearing design principle

> **The agent handles meaning; the renderer handles aesthetics.**

The agent never generates PPTX code. It emits spec YAML only. All visual rules live as
fixed code in the renderer, so every diagram comes out consistent regardless of what
the agent wrote. This separation is what guarantees quality.

### 1.2 Why not an existing tool

- **mingrammer/diagrams** has the icon packs but exports flat PNG/SVG — not editable.
- **D2** can emit `.pptx`, but each slide is a single rendered image, not native shapes.
- **Structurizr / Mermaid / Excalidraw** export images or text formats; none emit
  editable PowerPoint shapes.

archgen occupies the unfilled niche: **editable PPTX with native shapes** (boxes,
arrows, labels, legends are all editable objects; only the brand logos are pictures).

---

## 2. Scope

### In scope (v0.1)
- Spec → editable `.pptx` for a single slide/diagram.
- Azure + AWS + GCP + generic tech logos (simple-icons), all from day one.
- Automatic layout (Graphviz `dot` when present; pure-Python fallback otherwise) plus
  a deterministic semantic layout for cloud-network diagrams.
- Grouped containers (incl. nesting), directed edges (solid/dashed/dotted), auto legends.
- Agent prompt generation + optional direct LLM generation with a lint/retry loop.

### Out of scope (v0.1) — see [Roadmap](#9-roadmap)
- Multi-slide decks; IaC (Terraform/Bicep) → spec; general-purpose curved edge routing;
  provider-specific AWS resource/group variants beyond the service-level catalog.

---

## 3. Tech stack

| Concern | Choice | Notes |
|---|---|---|
| Language | **Python ≥ 3.11** (dev on 3.14) | |
| PPTX generation | **python-pptx** | native shapes, pictures, connectors |
| Spec schema/validation | **pydantic v2** | strict models, JSON-Schema export |
| Spec format | **YAML** (PyYAML) | human- and agent-friendly |
| Layout (primary) | **Graphviz `dot`** via `-Tjson` | optional; cluster support for groups |
| Layout (fallback) | **pure-Python** group-aware layered | zero-dependency; guarantees non-overlap |
| SVG→picture (vector) | **Inkscape** / **LibreOffice** → EMF | optional; best fidelity/editability |
| SVG→picture (raster) | **macOS QuickLook** (`qlmanage`) / **cairosvg** | zero-install / needs libcairo; white canvas flood-filled to transparent |
| Icons — tech logos | **logos** (CC0) → **Devicon** (MIT) → **simple-icons** (CC0) | full color; SI monochrome tinted to brand hex as fallback |
| Icons — generic | **Flat Color Icons** (MIT) | 28 color concept icons |
| Icons — cloud | official Azure/AWS/GCP packs | brand icons, never altered |
| CLI | **argparse** | `archgen <cmd>` |
| Tests | **pytest** | linter, geometry, layout, golden render, licensing |
| LLM (optional) | **anthropic** SDK | injected callable; model-agnostic |

**Runtime capability detection:** `archgen doctor` reports which layout engine and icon
converter are active on the current machine.

---

## 4. Architecture

```
prompt ─► [agent] ─► spec.yaml ─► [linter] ─► [layout] ─► [geometry] ─► [renderer] ─► out.pptx
            meaning only          fail loud    coords      pts→EMU       aesthetics (deterministic)
```

Five layers (see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the file-by-file map):

1. **registry** — canonical icon key → SVG file → embeddable picture (EMF/PNG), cached.
2. **spec** — pydantic models, fail-loud linter, JSON-Schema export, YAML loader.
3. **layout** — Graphviz adapter + pure-Python fallback; geometry transform to EMU.
4. **render** — `theme.yaml` design tokens + deterministic renderer + auto legends.
5. **agent** — system prompt (schema + icon catalog + rules) + prompt→spec→lint→retry.

---

## 5. Data model (the spec)

Full field reference: [docs/SPEC_FORMAT.md](docs/SPEC_FORMAT.md). Summary:

- `Diagram{ title, kind, provider_default, direction, nodes[], groups[], edges[], abbreviations{} }`
- `Node{ id, icon_key, label, group? }`
- `Group{ id, label, parent?, provider? }`
- `Edge{ src, dst, label?, style }`

The JSON Schema is generated from the pydantic models (`archgen schema`) and embedded
in the agent prompt, so the agent's output validates against the same source of truth.

---

## 6. Invariants (what must always hold)

These are enforced by the linter and/or tests; treat them as contracts when changing code:

1. **Every `icon_key` resolves** to a file in the manifest.
2. **No dangling/self edges; no orphan nodes** (a node is connected or grouped).
3. **Labels ≤ 28 chars** (`MAX_LABEL_LEN`); longer must be abbreviated.
4. **Every abbreviation used in a label is declared** in `abbreviations` (feeds the legend).
5. **Group containers never overlap** (guaranteed by both layout engines).
6. **Brand icons are placed uniform-scale, never recolored or distorted** — only CC0
   `tech.*` icons carry `recolorable: true`. (AWS CC BY-ND / Azure terms.)
7. **Output is editable**: nodes/labels/arrows/legends are native shapes; only logos are
   pictures. Verified by the golden render test.
8. **Legends are derived from spec content**, so they never go stale.

---

## 7. Pipeline stages (contracts)

| Stage | Input | Output | Fails when |
|---|---|---|---|
| agent | NL description | spec YAML | (model output; retried on lint failure) |
| linter | `Diagram` | `list[str]` errors | any invariant 1–4 violated |
| layout | `Diagram` | `Layout` (pts, y-up) | Graphviz requested but `dot` absent |
| geometry | `Layout` + slide dims | EMU rects | — |
| renderer | `Diagram` + `Layout` | `.pptx` | icon file missing |

---

## 8. Coordinate contract (the fragile part)

- **Graphviz/layout space:** points (72/in), origin bottom-left, y grows up, node `pos`
  is the center, `bb` = `x0,y0,x1,y1`.
- **PowerPoint:** EMU (914400/in), origin top-left, y grows down.
- `layout/geometry.Transform` does: uniform scale to fit the slide content area →
  center → **flip Y**. Covered by `tests/test_geometry.py`.

---

## 9. Roadmap

- **P1** — Optionally expose provider-specific AWS resource/group variants without
  polluting or destabilizing the service-level key catalog.
- **P1** — Vector EMF by default (bundle/instruct Inkscape or LibreOffice).
- **P2** — Smarter edge routing: flow edges now use native PowerPoint elbow (orthogonal)
  connectors and the network profile uses lane-based Manhattan routes; remaining work is
  channel separation so dense diagrams never share a lane.
- **P2** — Native `GroupShape` binding so a container + children move as one unit.
- **P3** — IaC → spec (parse Terraform/Bicep to derive nodes/edges/groups).
- **P3** — Multi-slide decks; theme variants beyond `default`.

---

## 10. Verification

```bash
pytest                                   # lint, geometry, layout, render, registry, licensing
archgen render examples/sample_spec.yaml -o out.pptx
qlmanage -t -s 1600 -o /tmp out.pptx     # eyeball the slide (macOS)
```

Golden render asserts: file opens, one picture per node, labels are editable text,
one connector + arrowhead per edge, both legends present, all shapes on-slide.
