"""Load a Diagram spec from YAML/JSON (file path or string).

Loading can fail three ways *before* the semantic linter ever runs: the file
is missing, the YAML is malformed, or the data doesn't match the schema. Each
is turned into a :class:`SpecLoadError` with a readable message instead of a
raw traceback — see :mod:`archgen.spec.lint` for errors in an otherwise valid
spec.
"""

from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import ValidationError

from archgen.spec.models import Diagram


class SpecLoadError(ValueError):
    """A spec could not be read, parsed, or validated against the schema."""


def load_spec(path: str | Path) -> Diagram:
    p = Path(path)
    try:
        text = p.read_text()
    except FileNotFoundError as exc:
        raise SpecLoadError(f"spec file not found: {p}") from exc
    except OSError as exc:
        raise SpecLoadError(f"cannot read spec file {p}: {exc}") from exc
    return _parse(text, source=str(p))


def loads_spec(text: str) -> Diagram:
    return _parse(text, source="<string>")


def _parse(text: str, source: str) -> Diagram:
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise SpecLoadError(_yaml_message(exc, source)) from exc
    if data is None:
        raise SpecLoadError(f"{source}: spec is empty.")
    if not isinstance(data, dict):
        raise SpecLoadError(f"{source}: spec must be a YAML mapping (got {type(data).__name__}).")
    try:
        return Diagram.model_validate(data)
    except ValidationError as exc:
        raise SpecLoadError(_schema_message(exc)) from exc


def _yaml_message(exc: yaml.YAMLError, source: str) -> str:
    problem = getattr(exc, "problem", None) or "could not parse YAML"
    mark = getattr(exc, "problem_mark", None)
    if mark is not None:
        return f"{source}: {problem} (line {mark.line + 1}, column {mark.column + 1})."
    return f"{source}: {problem}."


def _schema_message(exc: ValidationError) -> str:
    lines = [f"{'.'.join(str(p) for p in e['loc']) or '(root)'}: {e['msg']}" for e in exc.errors()]
    n = len(lines)
    head = f"spec does not match the schema ({n} problem{'' if n == 1 else 's'}):"
    return head + "\n  - " + "\n  - ".join(lines)
