"""Declarative diagram spec — the *only* artifact the agent produces.

Everything here describes **meaning** (what nodes/groups/edges exist), never
**aesthetics** (positions, colors, sizes). The renderer owns all aesthetics.

The models are intentionally strict (``extra="forbid"``) so the linter and the
agent get loud, early errors instead of silently-ignored typos.
"""

from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class Provider(str, Enum):
    """Icon/palette provider. Drives the color legend and default theming."""

    azure = "azure"
    aws = "aws"
    gcp = "gcp"
    tech = "tech"  # generic tech logos (simple-icons), e.g. kafka, postgres
    generic = "generic"  # provider-agnostic boxes (client, user, internet)


class EdgeStyle(str, Enum):
    solid = "solid"
    dashed = "dashed"
    dotted = "dotted"


class DiagramKind(str, Enum):
    """Semantic diagram family; selects the deterministic layout strategy."""

    flow = "flow"
    network = "network"


class NodeRole(str, Enum):
    """A node's architectural role, not its visual treatment."""

    service = "service"
    external = "external"
    internet = "internet"
    dns = "dns"
    gateway = "gateway"
    workload = "workload"
    management = "management"
    private_endpoint = "private_endpoint"
    network_anchor = "network_anchor"
    managed_service = "managed_service"
    observability = "observability"


class GroupKind(str, Enum):
    """What a container represents in the architecture."""

    tier = "tier"
    cloud = "cloud"
    network = "network"
    subnet = "subnet"
    cluster = "cluster"


class GroupRole(str, Enum):
    """Functional responsibility used to place network containers."""

    default = "default"
    boundary = "boundary"
    ingress = "ingress"
    compute = "compute"
    management = "management"
    private_access = "private_access"
    web = "web"
    storage = "storage"
    private_network = "private_network"
    data = "data"
    discovery = "discovery"
    cost = "cost"
    deployment = "deployment"
    integration = "integration"


class EdgeKind(str, Enum):
    """Connection meaning; the renderer maps it to line treatment and legends."""

    traffic = "traffic"
    https = "https"
    diagnostics = "diagnostics"
    dns = "dns"
    private_connection = "private_connection"
    virtual_network_link = "virtual_network_link"


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Node(_Strict):
    """A single icon+label box in the diagram."""

    id: str = Field(..., description="Unique node id, referenced by edges and group membership.")
    icon_key: str = Field(
        ...,
        description="Canonical icon key resolved against the registry manifest, e.g. 'azure.aks', 'aws.lambda', 'gcp.gke', 'tech.kafka'.",
    )
    label: str = Field(..., description="Short display label. Keep terse; use abbreviations declared in Diagram.abbreviations.")
    group: str | None = Field(None, description="Optional id of the Group this node belongs to.")
    role: NodeRole = Field(
        NodeRole.service,
        description="Semantic architectural role used by network diagrams; ignored by ordinary flow diagrams.",
    )


class Group(_Strict):
    """A container (VPC / subnet / resource-group / tier) drawn as a tinted box."""

    id: str = Field(..., description="Unique group id.")
    label: str = Field(..., description="Container title shown at the top-left of the box.")
    parent: str | None = Field(None, description="Optional parent group id for nesting.")
    provider: Provider | None = Field(
        None, description="Optional provider override for this container's tint; defaults to the diagram/provider of its members."
    )
    kind: GroupKind = Field(GroupKind.tier, description="Semantic container type.")
    role: GroupRole = Field(GroupRole.default, description="Functional role within a network diagram.")


class Edge(_Strict):
    """A directed connection between two nodes."""

    src: str = Field(..., description="Source node id.")
    dst: str = Field(..., description="Destination node id.")
    label: str | None = Field(None, description="Optional short edge label.")
    style: EdgeStyle = Field(EdgeStyle.solid, description="Line style.")
    kind: EdgeKind = Field(EdgeKind.traffic, description="Semantic connection type used for routing and derived legends.")


class Diagram(_Strict):
    """Root spec object."""

    title: str = Field(..., description="Slide/diagram title.")
    kind: DiagramKind = Field(
        DiagramKind.flow,
        description="Diagram family. 'network' activates the deterministic cloud-network layout.",
    )
    provider_default: Provider = Field(
        Provider.generic, description="Fallback provider for nodes/containers that don't imply one from their icon_key."
    )
    direction: Literal["TB", "LR"] = Field(
        "LR", description="Primary layout flow: top-to-bottom or left-to-right."
    )
    layered: bool = Field(
        False,
        description=(
            "Flow diagrams only. Declare that the groups are ordered tiers and should "
            "render as clean parallel bands (columns in LR, rows in TB) rather than "
            "staggering by edge depth. The renderer owns all sizing: it fits the "
            "densest tier to the slide, so many-node tiers (e.g. a source rail) stay "
            "legible. Use for layered architectures (source->ingest->store->serve->...). "
            "Requires Graphviz; ignored by network diagrams."
        ),
    )
    nodes: list[Node] = Field(default_factory=list)
    groups: list[Group] = Field(default_factory=list)
    edges: list[Edge] = Field(default_factory=list)
    abbreviations: dict[str, str] = Field(
        default_factory=dict,
        description="Abbreviation -> expansion, e.g. {'AKS': 'Azure Kubernetes Service'}. Every abbreviation appearing in a label MUST be declared here (the linter enforces it), and they populate the abbreviation legend.",
    )
    theme: str = Field("default", description="Theme name (reserved; only 'default' ships today).")
