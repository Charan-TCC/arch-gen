"""Export the spec's JSON Schema — consumed by the agent so it emits only valid specs."""

from __future__ import annotations

import json

from archgen.spec.models import Diagram


def json_schema() -> dict:
    """Return the JSON Schema for a :class:`Diagram` spec."""
    return Diagram.model_json_schema()


def json_schema_str(indent: int = 2) -> str:
    return json.dumps(json_schema(), indent=indent)


if __name__ == "__main__":
    print(json_schema_str())
