# Contributing / maintenance guide

How to set up, develop, extend, and release archgen. Architecture details:
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md). Project spec: [SPEC.md](docs/SPEC.md).

## Dev setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
python -m archgen.registry.fetch      # populate assets/icons + manifest.json
make doctor                           # or: archgen doctor
make test
```

Optional native tooling (graceful fallbacks exist without them):
- `brew install graphviz` — tighter grouped layouts (else pure-Python fallback).
- `brew install --cask inkscape` — vector EMF icons (else PNG via macOS QuickLook).

> **Note on this machine:** Homebrew uses a custom prefix that builds from source, so
> `graphviz`/`inkscape` are slow/unreliable to install here. The pure-Python layout and
> QuickLook rasterizer are the working defaults; everything runs without them.

## Common tasks

| Task | Command |
|---|---|
| Render a spec | `make render SPEC=examples/sample_spec.yaml OUT=out.pptx` |
| Preview a pptx (macOS) | `make preview OUT=out.pptx` |
| Lint a spec | `archgen lint <spec>` |
| Run tests | `make test` |
| Rebuild icon manifest | `make icons` (→ `python -m archgen.registry.fetch --no-download`) |
| Capability check | `make doctor` |

## How to extend

### Add an icon
- **Brand logo:** `archgen icons add <name> --key tech.<name>`. This is color-first —
  it caches full-color artwork from Iconify `logos`/`devicon`, else a Simple Icons
  monochrome logo tinted to its brand color. Provenance + `brand_hex` are recorded in
  `assets/icons/simple-icons-downloads.json` and survive a manifest rebuild.
- **Generic concept:** add an entry to `GENERIC_SEED` in `archgen/registry/fetch.py`
  (Flat Color Icons name → `generic.<key>`), then `make icons` (downloads it).
- **Curated cloud/tech seed:** add to `CLOUD_SEED`/`TECH_SEED`, and an
  `OFFICIAL_KEY_ALIAS` when it should point at an auto-discovered official service.
- **Local SVG:** drop it in `assets/icons/<provider>/` and run `make icons`.

Confirm any of the above with `archgen icons`. Full downloading pipeline (color logos +
Flat Color generic icons): `python -m archgen.registry.fetch`.

### Add a full official cloud pack (AWS/Azure/GCP)
1. Unzip the pack into `assets/icons/<provider>/`, preserving its directory layout.
2. Run `make icons`. AWS 48px architecture-service icons, nested Azure service SVGs,
   and flat GCP service SVGs are discovered automatically.
3. Review `archgen icons`. Add only genuinely useful short names to
   `OFFICIAL_KEY_ALIAS`; generated canonical keys remain the primary catalog.

AWS resource and architecture-group variants are intentionally excluded from automatic
discovery because they overlap service names and create ambiguous keys. All discovered
cloud icons are always written with `recolorable: false`.

### Change styling
Edit `archgen/render/theme.yaml` only — colors, fonts, sizes, spacing, palettes. No
code change needed.

### Add an edge style or node shape
`archgen/spec/models.py` (extend the enum) + `archgen/render/renderer.py` (draw it).

### Improve layout
Pure-Python path: `archgen/layout/simple_layout.py`. Graphviz path:
`archgen/layout/graphviz_layout.py`. Both must emit the same `Layout` convention
(points, y-up, node `pos` = center, `bb` = x0,y0,x1,y1).

## Conventions & invariants

- Keep the **agent-emits-spec / renderer-owns-aesthetics** split intact. No aesthetic
  decisions in the spec or the agent; no semantic decisions in the renderer.
- Preserve the invariants in [SPEC.md §6](docs/SPEC.md#6-invariants-what-must-always-hold) —
  especially: brand icons never recolored/distorted; output stays editable; containers
  never overlap.
- Add or update a test when you change behavior. The golden render test is the
  end-to-end guard.
- Coordinate math changes must keep `tests/test_geometry.py` green (points→EMU, Y-flip).

## Testing

```bash
pytest -q                 # full suite
pytest tests/test_render.py -q   # just the golden render
```

Manual visual check (macOS): render, then
`qlmanage -t -s 1600 -o /tmp out.pptx` and open the PNG.

## Release checklist

1. `make test` green.
2. Update `CHANGELOG.md` and bump `version` in `pyproject.toml`.
3. Verify `archgen render examples/sample_spec.yaml` opens cleanly in PowerPoint.
4. Tag the release.
