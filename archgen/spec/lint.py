"""Spec linter — fail loud *before* rendering.

Catches the agent's mistakes (dangling edges, unknown icons, undeclared
abbreviations, over-long labels, bad group nesting) so they never reach the
slide. Returns a list of human-readable errors; empty list == valid.
"""

from __future__ import annotations

import re

from archgen.registry.resolver import IconResolver
from archgen.spec.models import Diagram

# Labels longer than this wrap poorly / stretch boxes; the agent should
# abbreviate instead. Kept in sync with the renderer's max box width.
MAX_LABEL_LEN = 28

# Token used for abbreviation detection: 2+ char all-caps runs (AKS, RDS, VPC),
# optionally with digits (S3, EC2, K8S).
_ABBREV_RE = re.compile(r"\b[A-Z][A-Z0-9]{1,}\b")


class SpecError(ValueError):
    """Raised by :func:`validate` when ``raise_on_error`` is set."""


def lint(diagram: Diagram, resolver: IconResolver | None = None) -> list[str]:
    """Return a list of validation errors (empty == valid)."""
    errors: list[str] = []
    resolver = resolver or IconResolver()

    node_ids = [n.id for n in diagram.nodes]
    group_ids = [g.id for g in diagram.groups]

    # --- uniqueness -------------------------------------------------------
    errors += _dupes("node id", node_ids)
    errors += _dupes("group id", group_ids)

    node_id_set = set(node_ids)
    group_id_set = set(group_ids)

    # --- icons ------------------------------------------------------------
    for n in diagram.nodes:
        if not resolver.exists(n.icon_key):
            suggestion = resolver.suggest(n.icon_key)
            hint = f" Did you mean '{suggestion}'?" if suggestion else ""
            if n.icon_key.startswith("tech."):
                query = n.icon_key.split(".", 1)[1]
                hint += (
                    f" Search with `archgen icons search {query}` and cache an exact match "
                    "with `archgen icons add <slug>`."
                )
            errors.append(f"node '{n.id}': unknown icon_key '{n.icon_key}'.{hint}")

    # --- labels -----------------------------------------------------------
    for n in diagram.nodes:
        if len(n.label) > MAX_LABEL_LEN:
            errors.append(
                f"node '{n.id}': label '{n.label}' is {len(n.label)} chars "
                f"(max {MAX_LABEL_LEN}); shorten or use a declared abbreviation."
            )

    # --- group membership + nesting --------------------------------------
    for n in diagram.nodes:
        if n.group is not None and n.group not in group_id_set:
            errors.append(f"node '{n.id}': references missing group '{n.group}'.")
    for g in diagram.groups:
        if g.parent is not None and g.parent not in group_id_set:
            errors.append(f"group '{g.id}': references missing parent '{g.parent}'.")
    errors += _group_cycles(diagram)

    # --- edges ------------------------------------------------------------
    for i, e in enumerate(diagram.edges):
        if e.src not in node_id_set:
            errors.append(f"edge #{i} ({e.src}->{e.dst}): source node '{e.src}' does not exist.")
        if e.dst not in node_id_set:
            errors.append(f"edge #{i} ({e.src}->{e.dst}): destination node '{e.dst}' does not exist.")
        if e.src == e.dst:
            errors.append(f"edge #{i}: self-loop on '{e.src}' is not supported.")

    # --- orphan nodes -----------------------------------------------------
    connected = {e.src for e in diagram.edges} | {e.dst for e in diagram.edges}
    for n in diagram.nodes:
        if n.id not in connected and n.group is None:
            errors.append(
                f"node '{n.id}': orphan (no edges and no group); connect it or place it in a group."
            )

    # --- abbreviations ----------------------------------------------------
    declared = set(diagram.abbreviations)
    used: set[str] = set()
    for text in [diagram.title, *(n.label for n in diagram.nodes), *(g.label for g in diagram.groups)]:
        used |= set(_ABBREV_RE.findall(text))
    for abbr in sorted(used - declared):
        errors.append(
            f"abbreviation '{abbr}' is used in a label but not declared in Diagram.abbreviations "
            f"(needed for the abbreviation legend)."
        )

    return errors


def validate(diagram: Diagram, resolver: IconResolver | None = None, raise_on_error: bool = True) -> list[str]:
    """Lint and (optionally) raise :class:`SpecError` on any problem."""
    errors = lint(diagram, resolver)
    if errors and raise_on_error:
        raise SpecError("Spec failed validation:\n  - " + "\n  - ".join(errors))
    return errors


def _dupes(kind: str, ids: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for i in ids:
        if i in seen:
            out.append(f"duplicate {kind}: '{i}'.")
        seen.add(i)
    return out


def _group_cycles(diagram: Diagram) -> list[str]:
    parent = {g.id: g.parent for g in diagram.groups}
    out: list[str] = []
    for start in parent:
        seen: set[str] = set()
        cur: str | None = start
        while cur is not None:
            if cur in seen:
                out.append(f"group '{start}': parent chain forms a cycle.")
                break
            seen.add(cur)
            cur = parent.get(cur)
    return out
