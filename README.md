# archgen

Generate **neat, editable PowerPoint** cloud-architecture diagrams from a declarative
spec. An LLM agent (Codex/Claude) writes *only* the spec; a deterministic Python
renderer owns all the aesthetics — consistent sizing, spacing, colors, grouped
containers, tidy arrows, auto-fit labels, and auto-generated legends.

Agents and contributors must follow [DESIGN_GUIDELINES.md](docs/DESIGN_GUIDELINES.md)
before generating or delivering a slide. It defines the asset-only logo policy,
placeholder disclosure, flat/no-shadow treatment, connector integrity, text fitting,
grouping, and required visual QA.

```
prompt ─► [agent] ─► spec.yaml ─► [linter] ─► [layout] ─► [renderer] ─► out.pptx
            meaning only          fail loud     coords     aesthetics (deterministic)
```

> **Why not just use mingrammer/diagrams or D2?** They export flat PNG/SVG (or, for
> D2, a PPTX with the whole board as a single image). archgen emits **native editable
> shapes** — you can grab any box, arrow, or label in PowerPoint and change it. Icons
> are embedded as picture objects (vector EMF when a converter is available).

## Install

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e .
python -m archgen.registry.fetch     # populate the icon registry + manifest
```

Optional, for best results:
- `brew install graphviz` — tighter grouped/crossing-minimized layouts (a pure-Python
  fallback is used automatically when `dot` is absent).
- `brew install --cask inkscape` — embed icons as **vector EMF** (crisp, editable).
  Without it, icons are embedded as high-res PNG via macOS QuickLook (zero-install) or
  cairosvg. Run `archgen doctor` to see what's active.

## Use

```bash
# 1. Render a hand-written or agent-written spec:
archgen render examples/serverless_api.yaml -o out.pptx

# 2. Validate a spec without rendering:
archgen lint examples/serverless_api.yaml

# 3. Get the agent prompt to paste into Codex/Claude, then render its YAML:
archgen prompt "3-tier web app on Azure with a Kafka bus and an RDS database"
#   (paste prompt into your agent -> save its output as spec.yaml)
archgen render spec.yaml -o out.pptx

# Network architecture profile (VNet/subnets/DNS/private endpoints):
archgen render examples/azure_private_aks.yaml -o private-aks.pptx

# 4. Or let archgen call the model directly (needs ANTHROPIC_API_KEY):
archgen generate "..." -o out.pptx --save-spec spec.yaml

archgen icons     # list valid icon keys
archgen icons search "Next.js"                 # search Simple Icons online
archgen icons add nextdotjs --key tech.nextjs  # cache SVG + rebuild manifest
archgen schema    # dump the spec JSON Schema
archgen doctor    # report layout/converter capabilities
```

### Add a missing logo for future use

Archgen can search the trusted [Simple Icons](https://simpleicons.org/) catalog and
cache an exact result locally:

```bash
archgen icons search "WorkOS"
archgen icons add workos --key tech.workos --category auth
```

The SVG is downloaded from a version-pinned Simple Icons package, validated to reject
active/external content, saved under `assets/icons/tech/`, and recorded in
`assets/icons/simple-icons-downloads.json` with source, package version, license, and
brand-guideline metadata. The manifest is rebuilt automatically, so `tech.workos`
works offline in later specs and renders. Existing files are preserved unless
`--force` is supplied.

This online workflow is intentionally limited to Simple Icons; it does not scrape
arbitrary websites. AWS, Azure, and GCP service artwork must still come from the
providers' official icon packs. Simple Icons' project license does not waive brand
trademark rights, so review the displayed per-icon license and usage guidelines before
publishing a diagram.

## The spec

```yaml
title: "Multi-Cloud 3-Tier Web Application"
kind: flow                    # flow (default) or network
page: "16:9"                  # required for flow — 16:9 (locked deck) or free (slide grows); the user's call
direction: LR                 # LR (flow) or TB (stack)
abbreviations:                # every ABBR used in a label must be declared
  AKS: "Azure Kubernetes Service"
groups:
  - {id: compute, label: "Compute Tier (Azure)", provider: azure}
nodes:
  - {id: api, icon_key: azure.aks, label: "API on AKS", group: compute}
  - {id: bus, icon_key: tech.kafka, label: "Event Bus", group: compute}
edges:
  - {src: api, dst: bus, style: dashed, kind: traffic}
```

### Examples

Small, realistic specs to copy from (each generated from a one-line prompt — the prompt
is recorded in the file header):

| Example | Shows |
|---|---|
| [`serverless_api.yaml`](examples/serverless_api.yaml) | AWS serverless REST API (CloudFront → API Gateway → Lambda → DynamoDB) |
| [`event_driven_pipeline.yaml`](examples/event_driven_pipeline.yaml) | Event-driven pipeline (Kafka → consumer → Postgres/Redis, dead-letter queue) |
| [`rag_chatbot.yaml`](examples/rag_chatbot.yaml) | RAG/LLM app using color generic-concept icons + tech logos |
| [`sample_spec.yaml`](examples/sample_spec.yaml) | Multi-cloud 3-tier flow spanning Azure/AWS/GCP/tech icons |
| [`azure_private_aks.yaml`](examples/azure_private_aks.yaml) | `kind: network` profile — VNet, subnets, DNS zones, private endpoints |
| [`aws_discovery_platform.yaml`](examples/aws_discovery_platform.yaml) | Dense architecture-board profile (many grouped sections) |

For cloud network diagrams, set `kind: network` and describe architectural meaning
with node `role`, group `kind`/`role`, and edge `kind`. Archgen then owns the zones,
orthogonal routing, connection colors, and connection legend.

The **linter** rejects a spec (before rendering) for: unknown icon key, dangling/self
edge, orphan node, over-long label, bad group nesting, or an undeclared abbreviation.

## Architecture

| Module | Responsibility |
|---|---|
| `archgen/registry/` | icon manifest, key→file resolver, official-pack auto-discovery, color-logo + Simple Icons fetchers, SVG→picture converter (Inkscape/LibreOffice EMF → transparent QuickLook/cairosvg PNG) |
| `archgen/spec/` | pydantic models, JSON-Schema export, fail-loud linter, YAML loader |
| `archgen/layout/` | Graphviz `dot`, a group-aware pure-Python fallback, and a deterministic semantic network layout; geometry transform (points→EMU, Y-flip, fit-to-slide) |
| `archgen/render/` | `theme.yaml` design tokens + deterministic renderer (native elbow connectors, brand-color icon tinting) + auto legends |
| `archgen/agent/` | system prompt (schema + icon catalog + rules) and a prompt→spec→lint→retry loop |

## Icons & licensing

Every icon renders in **full color** and is embedded with a **transparent background**
(the QuickLook raster path flood-fills the white thumbnail canvas), so icons sit cleanly
on tinted containers rather than on white boxes.

- **Generic concept icons** (`generic.user`, `generic.database`, `generic.queue`,
  `generic.loadbalancer`, `generic.cache`, `generic.function`, … — 28 in all) come from
  [Icons8 Flat Color Icons](https://github.com/icons8/flat-color-icons) (MIT). They are
  multi-color and rendered as-is (never tinted).
- **Technology logos** (`tech.*`) resolve to real full-color artwork, tried in order:
  [Gil Barbara "logos"](https://github.com/gilbarbara/logos) (CC0) →
  [Devicon](https://github.com/devicons/devicon) (MIT) →
  [simple-icons](https://github.com/simple-icons/simple-icons) monochrome tinted to the
  brand's official color. All fetched via the [Iconify](https://iconify.design) API.
  Search/cache more with `archgen icons search/add` (color-first).
- **Cloud brand icons**: drop the official
  [Azure](https://learn.microsoft.com/en-us/azure/architecture/icons/) /
  [AWS](https://aws.amazon.com/architecture/icons/) /
  [GCP](https://cloud.google.com/icons) SVG packs into `assets/icons/<provider>/`,
  preserving the pack directories, then run `make icons`. Service icons are discovered
  automatically and receive canonical keys; a small alias table preserves familiar keys
  such as `aws.ec2` and `azure.aks`. Until a pack is present, the curated compatibility
  keys use clean provider-colored **placeholders**.
- **Guardrail:** cloud brand icons (AWS/Azure/GCP) are placed **uniform-scale only, never
  recolored or distorted** — required by the AWS CC BY-ND / Azure terms, enforced by the
  `recolorable` flag (never true for cloud packs). Simple Icons / logos / Devicon are
  brand trademarks: CC0/MIT covers the artwork, not the trademark, so don't imply
  endorsement — review brand guidelines before publishing.

## Tests

```bash
pytest        # linter, geometry, non-overlapping layout, golden render, licensing
```
