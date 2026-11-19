# Diagram spec format

The spec is the **only** thing an agent produces. It is YAML (JSON also works) and is
validated against the pydantic models in `archgen/spec/models.py`. Dump the machine
schema with `archgen schema`.

## Top level — `Diagram`

| Field | Type | Req | Default | Notes |
|---|---|---|---|---|
| `title` | string | ✅ | — | Slide title (top-left of the slide). |
| `kind` | `flow` \| `network` | | `flow` | Selects the ordinary graph layout or semantic cloud-network layout. |
| `page` | `16:9` \| `free` | ✅ (flow) | — | **The user's call — ask, don't assume.** `16:9` locks the standard 13.33"×7.5" deck slide (formal presentation; content must be reshaped to fit readably). `free` lets the slide grow so text never shrinks below a readable size (working diagrams/posters). Required for `flow`; the linter rejects a flow spec without it. `archgen render`/`generate` accept `--page` to set it from the CLI. |
| `direction` | `LR` \| `TB` | | `LR` | Flow: left-to-right or top-to-bottom. |
| `layered` | bool | | `false` | Flow only. Render groups as ordered tier bands (columns in LR, rows in TB) instead of staggering by edge depth; the renderer sizes cards to fit the densest tier, so many-node tiers stay legible. Requires Graphviz; ignored by `network`. |
| `provider_default` | provider | | `generic` | Fallback tint for containers with no clear provider. |
| `nodes` | `Node[]` | | `[]` | The boxes. |
| `groups` | `Group[]` | | `[]` | Containers (tiers/VPCs/resource groups). |
| `edges` | `Edge[]` | | `[]` | Directed connections. |
| `abbreviations` | map | | `{}` | `ABBR → expansion`; every abbreviation used in a label must appear here. |
| `theme` | string | | `default` | Style preset. `default` (theme.yaml) or a name from `styles/`: `corporate-clean`, `modern-tech`, `editorial`, `warm-minimal`, `bold-contrast`, `consulting-classic`, `data-dark`, `pastel-soft` — the same presets PPT Gen uses. Also `archgen render --theme <name>`. Provider icons keep official colours. |

**provider** ∈ `azure` · `aws` · `gcp` · `tech` · `generic`.

## `Node`

| Field | Type | Req | Notes |
|---|---|---|---|
| `id` | string | ✅ | Unique; referenced by edges and `group`. |
| `icon_key` | string | ✅ | Canonical key from the registry (`archgen icons`). |
| `label` | string | ✅ | ≤ 28 chars; use declared abbreviations for long names. |
| `group` | string | | Id of the containing group. |
| `role` | node role | | `service` | Semantic placement in a `network` diagram. |

**node role** ∈ `service` · `external` · `internet` · `dns` · `gateway` · `workload` ·
`management` · `private_endpoint` · `network_anchor` · `managed_service` ·
`observability`.

## `Group`

| Field | Type | Req | Notes |
|---|---|---|---|
| `id` | string | ✅ | Unique. |
| `label` | string | ✅ | Container title (top-left, inside the box). |
| `parent` | string | | Parent group id (nesting). |
| `provider` | provider | | Override the container tint; else inferred from members. |
| `kind` | group kind | | `tier` | `tier`, `cloud`, `network`, `subnet`, or `cluster`. |
| `role` | group role | | `default` | Placement zone for a network diagram. |

**group role** ∈ `default` · `boundary` · `ingress` · `compute` · `management` ·
`private_access` · `web` · `storage` · `private_network` · `data` · `discovery` ·
`cost` · `deployment` · `integration`.

The latter roles activate the dense architecture-board zones used for grouped AWS-style
platform diagrams. See `examples/aws_discovery_platform.yaml`.

## `Edge`

| Field | Type | Req | Default | Notes |
|---|---|---|---|---|
| `src` | string | ✅ | — | Source node id. |
| `dst` | string | ✅ | — | Destination node id. No self-loops. |
| `label` | string | | — | Optional short edge label. |
| `style` | `solid`\|`dashed`\|`dotted` | | `solid` | Line style. |
| `kind` | edge kind | | `traffic` | Connection meaning; drives routing, color, and the derived connection legend. |

**edge kind** ∈ `traffic` · `https` · `diagnostics` · `dns` ·
`private_connection` · `virtual_network_link`.

## Validation rules (linter)

A spec is rejected **before rendering** if any of these fail (`archgen lint <spec>`):

1. `icon_key` not in the manifest (a close-match suggestion is offered).
2. Edge `src`/`dst` references a non-existent node; or a self-loop.
3. A node is an **orphan** (no edges and no group).
4. A `label` exceeds 28 characters.
5. A `group`/`parent` references a missing group, or parents form a cycle.
6. An abbreviation (2+ uppercase run, e.g. `AKS`, `RDS`, `DB`) appears in a label/title
   but is **not declared** in `abbreviations`.
7. Duplicate node or group ids.

## Minimal example

```yaml
title: "API with a queue"
page: "16:9"          # required for flow — 16:9 (locked deck) or free (slide grows)
direction: LR
abbreviations:
  AKS: "Azure Kubernetes Service"
  API: "Application Programming Interface"
groups:
  - {id: compute, label: "Compute (Azure)", provider: azure}
nodes:
  - {id: api,  icon_key: azure.aks,  label: "API on AKS", group: compute}
  - {id: bus,  icon_key: tech.kafka, label: "Event Bus",  group: compute}
  - {id: db,   icon_key: aws.rds,    label: "DB"}
edges:
  - {src: api, dst: bus, style: dashed}
  - {src: api, dst: db}
```

See [../examples/sample_spec.yaml](../examples/sample_spec.yaml) for a fuller,
multi-tier, multi-cloud example.

For a network-profile example that recreates a private AKS deployment with DNS zones,
subnets, private endpoints, and diagnostics, see
[../examples/azure_private_aks.yaml](../examples/azure_private_aks.yaml).

## Icon keys

Naming: `<provider>.<service>` (`azure.aks`, `aws.lambda`, `gcp.gke`, `tech.kafka`,
`generic.user`). Official service packs are auto-discovered, with normalized keys such
as `aws.api-gateway`, `azure.openai`, and `gcp.cloud-run`; legacy short aliases remain
valid. `tech.*` logos render in full brand color; `generic.*` covers 28 vendor-neutral
concepts in color (`user`, `client`, `mobile`, `internet`, `identity`, `database`,
`storage`, `queue`, `cache`, `function`, `container`, `api`, `gateway`, `loadbalancer`,
`monitoring`, `analytics`, …) — use these when no exact brand logo fits. List everything
currently available with:

```bash
archgen icons
```
