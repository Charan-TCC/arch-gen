"""Shared files must stay byte-identical across Arch Gen and PPT Gen
(docs/COMPOSE_API.md). Skipped when the sibling repo is not checked out."""

from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SIBLING = ROOT.parent / "PPT Gen"
OWN_COMPOSE = ROOT / ("archgen" if (ROOT / "archgen").is_dir() else "pptgen") / "shared" / "compose.py"
SIB_COMPOSE = SIBLING / ("archgen" if (SIBLING / "archgen").is_dir() else "pptgen") / "shared" / "compose.py"

needs_sibling = pytest.mark.skipif(not SIBLING.is_dir(), reason="sibling repo not checked out")


@needs_sibling
def test_compose_identical():
    assert OWN_COMPOSE.read_bytes() == SIB_COMPOSE.read_bytes(), (
        f"{OWN_COMPOSE} and {SIB_COMPOSE} drifted — copy the edited one over the other")


@needs_sibling
def test_geomcheck_identical():
    own = OWN_COMPOSE.with_name("geomcheck.py")
    sib = SIB_COMPOSE.with_name("geomcheck.py")
    assert own.read_bytes() == sib.read_bytes(), "geomcheck.py drifted between repos"


@needs_sibling
def test_contract_doc_identical():
    a, b = ROOT / "docs/COMPOSE_API.md", SIBLING / "docs/COMPOSE_API.md"
    assert a.read_bytes() == b.read_bytes()


@needs_sibling
def test_style_presets_identical():
    own = {p.name: p.read_bytes() for p in (ROOT / "styles").glob("*.json")}
    sib = {p.name: p.read_bytes() for p in (SIBLING / "styles").glob("*.json")}
    assert own.keys() == sib.keys(), "preset sets differ"
    for name in own:
        assert own[name] == sib[name], f"styles/{name} drifted"


def test_compose_has_no_repo_imports():
    import re
    src = OWN_COMPOSE.read_text()
    bad = re.findall(r"^\s*(?:from|import)\s+(archgen|pptgen)\b", src, re.M)
    assert not bad, f"compose.py must not import repo packages: {bad}"
