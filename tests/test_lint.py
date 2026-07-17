"""Linter unit tests: each malformed spec fails with a clear message; valid passes."""

from __future__ import annotations

import pytest

from archgen.spec.lint import lint
from archgen.spec.models import Diagram, Edge, Group, Node


def _valid() -> Diagram:
    return Diagram(
        title="Test",
        page="16:9",
        nodes=[
            Node(id="a", icon_key="azure.aks", label="API on AKS", group="g"),
            Node(id="b", icon_key="tech.kafka", label="Bus"),
        ],
        groups=[Group(id="g", label="Tier")],
        edges=[Edge(src="a", dst="b")],
        abbreviations={"AKS": "Azure Kubernetes Service", "API": "Application Programming Interface"},
    )


def test_valid_spec_passes():
    assert lint(_valid()) == []


def test_unknown_icon_key():
    d = _valid()
    d.nodes[0].icon_key = "azure.does_not_exist"
    errs = lint(d)
    assert any("unknown icon_key" in e for e in errs)


def test_dangling_edge():
    d = _valid()
    d.edges.append(Edge(src="a", dst="ghost"))
    errs = lint(d)
    assert any("ghost" in e and "does not exist" in e for e in errs)


def test_self_loop():
    d = _valid()
    d.edges.append(Edge(src="a", dst="a"))
    assert any("self-loop" in e for e in lint(d))


def test_label_too_long():
    d = _valid()
    d.nodes[1].label = "x" * 40
    assert any("chars" in e for e in lint(d))


def test_undeclared_abbreviation():
    d = _valid()
    d.nodes[1].label = "RDS store"  # RDS not declared
    assert any("RDS" in e and "not declared" in e for e in lint(d))


def test_orphan_node():
    d = _valid()
    d.nodes.append(Node(id="lonely", icon_key="tech.redis", label="Cache"))
    assert any("orphan" in e for e in lint(d))


def test_missing_group_reference():
    d = _valid()
    d.nodes[0].group = "nope"
    assert any("missing group" in e for e in lint(d))


def test_group_cycle():
    d = _valid()
    d.groups = [Group(id="g", label="G", parent="h"), Group(id="h", label="H", parent="g")]
    assert any("cycle" in e for e in lint(d))
