"""Load and expose design tokens from theme.yaml."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import yaml
from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

_THEME_PATH = Path(__file__).resolve().parent / "theme.yaml"


class Theme:
    def __init__(self, data: dict):
        self._d = data

    # --- raw access -------------------------------------------------------
    def __getitem__(self, key: str):
        return self._d[key]

    @property
    def font(self) -> str:
        return self._d["font_family"]

    def provider(self, name: str) -> dict:
        return self._d["providers"].get(name, self._d["providers"]["generic"])

    # --- helpers ----------------------------------------------------------
    @staticmethod
    def rgb(hex_str: str) -> RGBColor:
        return RGBColor.from_string(hex_str.lstrip("#"))

    @property
    def slide_w(self) -> Emu:
        return Inches(self._d["slide"]["width_in"])

    @property
    def slide_h(self) -> Emu:
        return Inches(self._d["slide"]["height_in"])

    @property
    def margin(self) -> Emu:
        return Inches(self._d["slide"]["margin_in"])


STYLES_DIR = Path(__file__).resolve().parents[2] / "styles"


def list_presets() -> list[str]:
    """Generic style presets (styles/*.json — same files and names as PPT Gen)."""
    return sorted(p.stem for p in STYLES_DIR.glob("*.json"))


def load_preset_tokens(name: str) -> dict:
    import json
    path = STYLES_DIR / f"{name}.json"
    if not path.is_file():
        raise ValueError(f"unknown theme preset {name!r}; one of: default, {', '.join(list_presets())}")
    return json.loads(path.read_text())


def _overlay_preset(data: dict, tok: dict) -> dict:
    """Map a shared preset onto archgen's diagram tokens. Provider colours stay
    official (brand icons are never recoloured); only neutrals, text, lines and
    fonts follow the preset."""
    import copy
    d = copy.deepcopy(data)
    c = {k: v.lstrip("#").upper() for k, v in tok["color"].items()}
    f = tok.get("font", {})
    st = tok.get("style", {})
    d["font_family"] = f.get("body", f.get("primary", d["font_family"]))
    d["background"] = c["background"]
    d["title_color"] = c["text-primary"]
    d["node"].update(box_fill=c["background"], box_line=c["border"], label_color=c["text-primary"])
    d["group"].update(title_color=c["text-secondary"])
    if "corner_radius" in st:
        d["group"]["corner_radius_in"] = round(0.03 + float(st["corner_radius"]) * 1.2, 3)
    d["edge"].update(color=c["text-secondary"], label_color=c["text-secondary"])
    d["legend"].update(box_fill=c["background"], box_line=c["border"])
    d["providers"]["generic"].update(tint=c["surface"], line=c["border"], hex=c["accent"])
    d["providers"]["tech"].update(tint=c["surface-alt"], line=c["border"])
    if tok.get("mode") == "dark":
        # Provider tints are light pastels; on a dark ground they would put light
        # label text on a light fill. Keep the brand line colour, darken the fill.
        for prov in d["providers"].values():
            prov["tint"] = c["surface"]
        d["network"].update(panel_fill=c["surface"], panel_line=c["border"])
    d["preset"] = tok.get("preset")
    return d


@lru_cache(maxsize=16)
def load_theme(path: Path | None = None, preset: str | None = None) -> Theme:
    p = path or _THEME_PATH
    data = yaml.safe_load(p.read_text())
    if preset and preset != "default":
        data = _overlay_preset(data, load_preset_tokens(preset))
    return Theme(data)
