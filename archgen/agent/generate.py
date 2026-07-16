"""Prompt -> validated spec, with a lint/retry loop.

The LLM is injected as a ``complete(system, user) -> str`` callable so archgen
stays model-agnostic (use Codex, Claude, or any agent). A convenience
:func:`anthropic_complete` is provided when ``ANTHROPIC_API_KEY`` is set.
"""

from __future__ import annotations

import re
from typing import Callable, Protocol

import yaml

from archgen.agent.prompt import build_system_prompt, build_user_prompt
from archgen.registry.resolver import IconResolver
from archgen.spec.lint import lint
from archgen.spec.models import Diagram

Complete = Callable[[str, str], str]


class GenerationError(RuntimeError):
    pass


def parse_spec_text(text: str) -> Diagram:
    """Parse an LLM response into a Diagram, tolerating ```yaml fences."""
    cleaned = _strip_fences(text)
    data = yaml.safe_load(cleaned)
    if not isinstance(data, dict):
        raise GenerationError("model output did not parse to a mapping/spec")
    return Diagram.model_validate(data)


def generate_spec(
    description: str,
    complete: Complete,
    resolver: IconResolver | None = None,
    max_retries: int = 2,
) -> Diagram:
    """Ask the model for a spec; on lint failure, feed errors back and retry."""
    resolver = resolver or IconResolver()
    system = build_system_prompt(resolver)
    user = build_user_prompt(description)
    last_errors: list[str] = []
    for _ in range(max_retries + 1):
        raw = complete(system, user)
        try:
            diagram = parse_spec_text(raw)
            errors = lint(diagram, resolver)
        except Exception as exc:  # noqa: BLE001 - malformed output is retryable
            errors = [f"could not parse/validate output: {exc}"]
            diagram = None
        if diagram is not None and not errors:
            return diagram
        last_errors = errors
        user = (
            build_user_prompt(description)
            + "\n\nYour previous spec was rejected for these reasons — fix them and "
            "return the corrected spec only:\n  - " + "\n  - ".join(errors)
        )
    raise GenerationError("spec still invalid after retries:\n  - " + "\n  - ".join(last_errors))


def anthropic_complete(model: str = "claude-sonnet-5") -> Complete:
    """Return a ``complete`` backed by the Anthropic API (requires the SDK + key)."""

    def _complete(system: str, user: str) -> str:
        import anthropic  # lazy; optional dependency

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=model,
            max_tokens=4096,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        return "".join(block.text for block in msg.content if getattr(block, "type", "") == "text")

    return _complete


def _strip_fences(text: str) -> str:
    m = re.search(r"```(?:ya?ml|json)?\s*(.*?)```", text, re.DOTALL)
    return m.group(1).strip() if m else text.strip()
