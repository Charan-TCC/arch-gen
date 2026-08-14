#!/usr/bin/env bash
# Render every example spec (+ the bespoke azure_modern_data_platform board) to a
# PNG in docs/gallery/. Run this ONCE after changing the renderer/layout/theme or
# adding an example. Purpose: keep committed reference images so nobody has to
# re-render just to see what a spec looks like. Reads specs from examples/,
# writes docs/gallery/<name>.png.
#
# Requires: the archgen venv (invoke via `make gallery`, which passes ARCHGEN/PY;
# running this script directly falls back to .venv/bin/*). Uses macOS `qlmanage`
# for rasterization (no Graphviz/LibreOffice needed in this env).
set -euo pipefail
cd "$(dirname "$0")/.."

ARCHGEN="${ARCHGEN:-.venv/bin/archgen}"
PY="${PY:-.venv/bin/python}"

OUT="docs/gallery"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
mkdir -p "$OUT"

render() {                       # render <spec.yaml> <name>
  local spec="$1" name="$2"
  echo "• $name"
  "$ARCHGEN" render "$spec" -o "$TMP/$name.pptx" --auto-direction >/dev/null
  rm -f "$OUT/$name.png"
  qlmanage -t -s 2000 -o "$TMP" "$TMP/$name.pptx" >/dev/null 2>&1
  mv "$TMP/$name.pptx.png" "$OUT/$name.png"
}

for spec in examples/*.yaml; do
  render "$spec" "$(basename "${spec%.yaml}")"
done

# Bespoke, layered Azure modern data platform (custom python-pptx, not the
# archgen renderer) — the only bespoke script wired into this gallery. If it's
# ever moved or renamed, this MUST fail loudly rather than silently skip: a
# quiet `if [ -f ... ]` guard is exactly how docs/gallery/shine_migration.png
# went stale (its script, build_shine.py, was referenced here for years but
# never actually existed in the repo).
BESPOKE=custom_builds/build_azure_data_platform.py
if [ ! -f "$BESPOKE" ]; then
  echo "error: $BESPOKE not found — fix this script or remove its gallery entry" >&2
  exit 1
fi
echo "• azure_modern_data_platform (bespoke)"
"$PY" "$BESPOKE" "$TMP/admp.pptx" >/dev/null
rm -f "$OUT/azure_modern_data_platform.png"
qlmanage -t -s 2000 -o "$TMP" "$TMP/admp.pptx" >/dev/null 2>&1
mv "$TMP/admp.pptx.png" "$OUT/azure_modern_data_platform.png"

echo "Gallery written to $OUT/"
