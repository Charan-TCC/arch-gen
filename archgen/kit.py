"""Arch Gen adapter for the shared compose kit (archgen/shared/compose.py).

compose.py is a shared, dependency-free file — identical in PPT Gen as
pptgen/compose.py. This module supplies the Arch Gen side: an IconSource backed
by the icon registry, and Style from a preset name. Bespoke builds in
bespoke/<name>/build.py should use:

    from archgen.shared import compose
    from archgen.kit import RegistryIcons, style
    st, icons = style("corporate-clean"), RegistryIcons()
    compose.icon_card(slide, st, icons, "azure.databricks", "Databricks", x, y, w, h)
"""

from __future__ import annotations

from archgen.shared import compose
from archgen.registry import emf
from archgen.registry.resolver import IconResolver
from archgen.render.theme import load_preset_tokens


class RegistryIcons:
    """compose.IconSource over archgen's registry. Brand icons are never recoloured
    (only recolourable monochrome logos get their official brand_hex)."""

    def __init__(self, resolver: IconResolver | None = None):
        self.r = resolver or IconResolver()

    def png(self, key: str, size_in: float, tone: str = "auto"):
        if not key:
            return None
        e = self.r.get(key)
        recolor = f"#{e.brand_hex}" if (e.recolorable and e.brand_hex) else None
        if tone == "inverse" and e.recolorable and key.startswith("generic."):
            recolor = "#FFFFFF"
        return str(emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False))


def style(preset: str = "corporate-clean") -> compose.Style:
    """compose.Style for a shared preset name (styles/*.json)."""
    return compose.Style.from_tokens(load_preset_tokens(preset))
