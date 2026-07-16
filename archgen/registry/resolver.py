"""Icon registry resolver: canonical key -> on-disk SVG path + metadata.

The manifest (``manifest.json``) is the single source of truth, produced/refreshed
by ``registry/fetch.py``. The spec references canonical keys (``azure.aks``); the
resolver turns them into concrete file paths for the EMF converter.
"""

from __future__ import annotations

import difflib
import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

_REGISTRY_DIR = Path(__file__).resolve().parent
_ASSETS_DIR = _REGISTRY_DIR.parent.parent / "assets" / "icons"
_MANIFEST = _REGISTRY_DIR / "manifest.json"


@dataclass(frozen=True)
class IconEntry:
    key: str
    path: Path
    provider: str
    display_name: str
    category: str
    recolorable: bool  # True only for CC0 sources (simple-icons) — never recolor brand icons
    brand_hex: str = ""  # official brand color for recolorable logos (e.g. "#FF4438")


class IconResolver:
    """Resolves icon keys against the manifest. Cheap to construct (cached load)."""

    def __init__(self, manifest_path: Path | None = None, assets_dir: Path | None = None):
        self._manifest_path = manifest_path or _MANIFEST
        self._assets_dir = assets_dir or _ASSETS_DIR
        self._entries = _load_manifest(self._manifest_path, self._assets_dir)

    def exists(self, key: str) -> bool:
        entry = self._entries.get(key)
        return entry is not None and entry.path.is_file()

    def get(self, key: str) -> IconEntry:
        entry = self._entries.get(key)
        if entry is None:
            raise KeyError(f"unknown icon key '{key}'")
        if not entry.path.is_file():
            raise FileNotFoundError(
                f"icon '{key}' is in the manifest but its SVG is missing at {entry.path}. "
                f"Run `python -m archgen.registry.fetch` to (re)download icon packs."
            )
        return entry

    def keys(self) -> list[str]:
        return sorted(self._entries)

    def suggest(self, key: str) -> str | None:
        matches = difflib.get_close_matches(key, self._entries.keys(), n=1, cutoff=0.6)
        return matches[0] if matches else None


@lru_cache(maxsize=8)
def _load_manifest(manifest_path: Path, assets_dir: Path) -> dict[str, IconEntry]:
    if not manifest_path.is_file():
        return {}
    raw = json.loads(manifest_path.read_text())
    entries: dict[str, IconEntry] = {}
    for key, meta in raw.get("icons", {}).items():
        entries[key] = IconEntry(
            key=key,
            path=(assets_dir / meta["file"]).resolve(),
            provider=meta.get("provider", "generic"),
            display_name=meta.get("display_name", key),
            category=meta.get("category", ""),
            recolorable=bool(meta.get("recolorable", False)),
            brand_hex=str(meta.get("brand_hex", "")),
        )
    return entries
