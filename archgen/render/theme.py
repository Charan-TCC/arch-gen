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


@lru_cache(maxsize=4)
def load_theme(path: Path | None = None) -> Theme:
    p = path or _THEME_PATH
    return Theme(yaml.safe_load(p.read_text()))
