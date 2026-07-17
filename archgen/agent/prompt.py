"""Build the system prompt an LLM agent (Codex/Claude/…) uses to emit a spec.

The agent produces *only* a YAML spec — never PPTX code. The prompt embeds the
JSON Schema, the exact set of valid icon keys, and the house rules the linter
enforces, so the agent's output validates on the first try as often as possible.
"""

from __future__ import annotations

from archgen.registry.resolver import IconResolver
from archgen.spec.lint import MAX_LABEL_LEN
from archgen.spec.schema import json_schema_str

_RULES = f"""\
You convert a plain-English description of a system into a diagram SPEC for the
`archgen` renderer. Output the SPEC only — no prose, no PPTX, no code fences.

Hard rules (the spec is auto-validated and rejected if any are broken):
1. Output valid YAML matching the JSON Schema below.
2. Every node `icon_key` MUST be one of the VALID ICON KEYS listed below. If no
   exact match exists, choose the closest generic key (e.g. `generic.database`,
   `generic.queue`, `generic.client`).
3. Keep every `label` <= {MAX_LABEL_LEN} characters. Prefer abbreviations.
4. Any abbreviation you use in a label (2+ uppercase letters, e.g. AKS, RDS, DB)
   MUST be declared in `abbreviations` with its full expansion.
5. Every edge `src`/`dst` must reference an existing node `id`. No self-loops.
6. Every node must either be in a `group` or be connected by at least one edge.
7. Group related nodes into `groups` (tiers, VPCs, resource groups). Set a
   group's `provider` when it clearly belongs to one cloud.
8. Choose `direction: LR` for request/data flows, `TB` for layered stacks.
8b. `page` is REQUIRED for flow diagrams and is the user's decision — ask them:
    `"16:9"` for a formal deck (fixed 13.33"x7.5" slide; keep the diagram compact
    enough to fit) or `"free"` for a working diagram (the slide grows so text
    stays readable). If the description doesn't say, ask; do not guess.
9. Use `kind: network` for cloud-network architecture with VNets/VPCs, subnets,
   DNS zones, gateways, private endpoints, or managed services. Assign meaningful
   node `role` and group `kind`/`role`; never invent coordinates or visual styles.
10. Set each edge `kind` to its connection meaning (`https`, `diagnostics`, `dns`,
    `private_connection`, `virtual_network_link`, or ordinary `traffic`).
"""


def build_system_prompt(resolver: IconResolver | None = None) -> str:
    resolver = resolver or IconResolver()
    keys = resolver.keys()
    catalog = "\n".join(
        f"  {k}  ({resolver.get(k).display_name})" if resolver.exists(k) else f"  {k}"
        for k in keys
    )
    return (
        _RULES
        + "\n\nVALID ICON KEYS:\n"
        + catalog
        + "\n\nJSON SCHEMA:\n"
        + json_schema_str()
    )


def build_user_prompt(description: str) -> str:
    return f"Produce the archgen spec for this system:\n\n{description}"
