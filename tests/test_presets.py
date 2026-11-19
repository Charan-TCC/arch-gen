"""Style presets (styles/*.json, shared with PPT Gen) drive diagram styling."""

import pytest
from pptx import Presentation

from archgen.shared import compose
from archgen.kit import style
from archgen.render.renderer import Renderer
from archgen.render.theme import list_presets, load_theme
from archgen.spec.io import load_spec


def test_presets_available():
    assert len(list_presets()) >= 6


@pytest.mark.parametrize("name", list_presets())
def test_preset_overlay_and_style(name):
    t = load_theme(preset=name)
    assert t["preset"] == name and t["background"]
    st = style(name)
    assert isinstance(st, compose.Style)


def test_dark_preset_darkens_provider_tints():
    t = load_theme(preset="modern-tech")
    assert t.provider("aws")["tint"] == t["background"] or t.provider("aws")["tint"] != "FFF3E0"


def test_spec_theme_renders(tmp_path):
    d = load_spec("examples/serverless_api.yaml").model_copy(update={"theme": "data-dark"})
    out = Renderer().render(d, tmp_path / "x.pptx")
    slide = Presentation(str(out)).slides[0]
    assert str(slide.background.fill.fore_color.rgb) == load_theme(preset="data-dark")["background"]


def test_unknown_preset_errors():
    with pytest.raises(ValueError):
        load_theme(preset="nope")
