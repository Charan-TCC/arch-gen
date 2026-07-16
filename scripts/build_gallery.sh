#!/usr/bin/env bash
# Render every example spec (+ the bespoke shine board) to a PNG in docs/gallery/.
# Run this ONCE after changing the renderer/layout/theme or adding an example.
# Purpose: keep committed reference images so nobody has to re-render just to see
# what a spec looks like. Reads specs from examples/, writes docs/gallery/<name>.png.
#
# Requires: the archgen venv active (or run via `make gallery`). Uses macOS
# `qlmanage` for rasterization (no Graphviz/LibreOffice needed in this env).
set -euo pipefail
cd "$(dirname "$0")/.."

OUT="docs/gallery"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
mkdir -p "$OUT"

render() {                       # render <spec.yaml> <name>
  local spec="$1" name="$2"
  echo "• $name"
  archgen render "$spec" -o "$TMP/$name.pptx" >/dev/null
  rm -f "$OUT/$name.png"
  qlmanage -t -s 2000 -o "$TMP" "$TMP/$name.pptx" >/dev/null 2>&1
  mv "$TMP/$name.pptx.png" "$OUT/$name.png"
}

for spec in examples/*.yaml; do
  render "$spec" "$(basename "${spec%.yaml}")"
done

# Bespoke, image-faithful board (custom python-pptx, not the archgen renderer).
if [ -f build_shine.py ]; then
  echo "• shine_migration (bespoke)"
  python build_shine.py >/dev/null
  rm -f "$OUT/shine_migration.png"
  qlmanage -t -s 2000 -o "$TMP" shine_migration.pptx >/dev/null 2>&1
  mv "$TMP/shine_migration.pptx.png" "$OUT/shine_migration.png"
fi

# Bespoke, layered Azure modern data platform (custom python-pptx).
if [ -f custom_builds/build_azure_data_platform.py ]; then
  echo "• azure_modern_data_platform (bespoke)"
  python custom_builds/build_azure_data_platform.py "$TMP/admp.pptx" >/dev/null
  rm -f "$OUT/azure_modern_data_platform.png"
  qlmanage -t -s 2000 -o "$TMP" "$TMP/admp.pptx" >/dev/null 2>&1
  mv "$TMP/admp.pptx.png" "$OUT/azure_modern_data_platform.png"
fi

echo "Gallery written to $OUT/"
