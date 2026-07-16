"""Search and download reusable brand icons from the Simple Icons catalog.

This module intentionally does not accept arbitrary URLs.  Search metadata comes
from the published ``simple-icons`` npm package and SVGs are downloaded from a
version-pinned jsDelivr URL.  Downloaded assets and provenance are stored under
``assets/icons`` so later lint/render runs are fully offline.
"""

from __future__ import annotations

import difflib
import json
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

_REGISTRY_DIR = Path(__file__).resolve().parent
_ASSETS_DIR = _REGISTRY_DIR.parent.parent / "assets" / "icons"
DOWNLOADS_FILE = "simple-icons-downloads.json"

_NPM_LATEST_URL = "https://registry.npmjs.org/simple-icons/latest"
_CATALOG_URL = "https://cdn.jsdelivr.net/npm/simple-icons@{version}/data/simple-icons.json"
_ICON_URL = "https://cdn.jsdelivr.net/npm/simple-icons@{version}/icons/{slug}.svg"

# Full-color brand-logo sets, served (versioned + CORS-safe) by the Iconify API.
# Order = preference: Gil Barbara "logos" (CC0) first, then Devicon (MIT).
_ICONIFY_URL = "https://api.iconify.design/{prefix}/{name}.svg"
COLOR_LOGO_PREFIXES = ("logos", "devicon")
_SLUG_RE = re.compile(r"^[a-z0-9]+$")
_KEY_RE = re.compile(r"^tech\.[a-z0-9][a-z0-9-]*$")
_MAX_SVG_BYTES = 512 * 1024
_BANNED_SVG_ELEMENTS = {"script", "style", "foreignobject", "iframe", "object", "embed"}


class OnlineIconError(RuntimeError):
    """An online catalog, lookup, download, or SVG validation failure."""


@dataclass(frozen=True)
class CatalogIcon:
    title: str
    slug: str
    hex: str = ""
    source_url: str = ""
    guidelines_url: str = ""
    license_type: str = ""
    license_url: str = ""

    @property
    def license_display(self) -> str:
        return self.license_type or "not provided"


@dataclass(frozen=True)
class DownloadResult:
    key: str
    path: Path
    icon: CatalogIcon
    package_version: str
    already_present: bool


def fetch_catalog(session=None) -> tuple[str, list[CatalogIcon]]:
    """Return the latest published Simple Icons version and searchable catalog."""
    package = _get_json(_NPM_LATEST_URL, session)
    version = str(package.get("version", "")).strip()
    if not re.fullmatch(r"\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?", version):
        raise OnlineIconError("Simple Icons registry returned an invalid package version")

    payload = _get_json(_CATALOG_URL.format(version=version), session)
    raw_icons = payload.get("icons") if isinstance(payload, dict) else payload
    if not isinstance(raw_icons, list):
        raise OnlineIconError("Simple Icons catalog has an unexpected format")

    icons: list[CatalogIcon] = []
    for raw in raw_icons:
        if not isinstance(raw, dict):
            continue
        title = str(raw.get("title", "")).strip()
        slug = str(raw.get("slug", "")).strip()
        if not title or not _SLUG_RE.fullmatch(slug):
            continue
        license_meta = raw.get("license") if isinstance(raw.get("license"), dict) else {}
        icons.append(
            CatalogIcon(
                title=title,
                slug=slug,
                hex=str(raw.get("hex", "")).strip(),
                source_url=str(raw.get("source", "")).strip(),
                guidelines_url=str(raw.get("guidelines", "")).strip(),
                license_type=str(license_meta.get("type", "")).strip(),
                license_url=str(license_meta.get("url", "")).strip(),
            )
        )
    if not icons:
        raise OnlineIconError("Simple Icons catalog did not contain any usable entries")
    return version, icons


def search_catalog(query: str, catalog: Iterable[CatalogIcon], limit: int = 10) -> list[CatalogIcon]:
    """Fuzzy-search titles and slugs, with exact/prefix matches ranked first."""
    needle = _normal(query)
    if not needle:
        return []

    ranked: list[tuple[float, str, CatalogIcon]] = []
    for icon in catalog:
        title = _normal(icon.title)
        slug = _normal(icon.slug)
        if needle in (title, slug):
            score = 4.0
        elif title.startswith(needle) or slug.startswith(needle):
            score = 3.0
        elif needle in title or needle in slug:
            score = 2.0
        else:
            score = max(
                difflib.SequenceMatcher(None, needle, title).ratio(),
                difflib.SequenceMatcher(None, needle, slug).ratio(),
            )
            if score < 0.48:
                continue
        ranked.append((score, icon.title.casefold(), icon))
    ranked.sort(key=lambda item: (-item[0], item[1], item[2].slug))
    return [item[2] for item in ranked[: max(1, min(limit, 50))]]


def resolve_exact(query: str, catalog: Iterable[CatalogIcon]) -> CatalogIcon:
    """Resolve an exact title or slug; offer search candidates on ambiguity."""
    needle = _normal(query)
    exact = [icon for icon in catalog if needle in (_normal(icon.title), _normal(icon.slug))]
    if len(exact) == 1:
        return exact[0]
    candidates = search_catalog(query, catalog, limit=5)
    if not candidates:
        raise OnlineIconError(f"no Simple Icons result for {query!r}")
    options = ", ".join(f"{icon.title} ({icon.slug})" for icon in candidates)
    raise OnlineIconError(
        f"{query!r} is not an exact Simple Icons title or slug; choose one of: {options}"
    )


def download_simple_icon(
    query: str,
    *,
    key: str | None = None,
    category: str = "brand",
    force: bool = False,
    assets_dir: Path | None = None,
    session=None,
) -> DownloadResult:
    """Download one exact catalog match and persist its provenance metadata."""
    assets_dir = assets_dir or _ASSETS_DIR
    version, catalog = fetch_catalog(session)
    icon = resolve_exact(query, catalog)
    registry_key = key or f"tech.{icon.slug}"
    if not _KEY_RE.fullmatch(registry_key):
        raise OnlineIconError("downloaded icon key must match 'tech.<lowercase-name>'")

    dest = assets_dir / "tech" / f"{registry_key.split('.', 1)[1]}.svg"
    already_present = dest.is_file()
    if already_present and not force:
        recorded = load_downloaded_metadata(assets_dir).get(registry_key, {})
        if isinstance(recorded, dict) and recorded.get("simple_icons_slug") == icon.slug:
            return DownloadResult(registry_key, dest, icon, version, True)
        raise OnlineIconError(
            f"{dest} already exists without matching Simple Icons provenance; "
            "choose another --key or use --force to replace it"
        )

    # Prefer a full-color logo (Gil Barbara logos, then Devicon); fall back to
    # the monochrome Simple Icons artwork, which the renderer tints to brand hex.
    color = fetch_color_logo(color_logo_candidates(icon.slug, icon.title), session=session)
    if color is not None:
        svg = color.svg
        source, recolorable, color_name = color.prefix, False, color.name
    else:
        svg = _get_bytes(_ICON_URL.format(version=version, slug=icon.slug), session)
        _validate_svg(svg)
        source, recolorable, color_name = "simple-icons", True, ""
    dest.parent.mkdir(parents=True, exist_ok=True)
    temp = dest.with_suffix(".svg.tmp")
    temp.write_bytes(svg)
    temp.replace(dest)
    _record_download(
        assets_dir, registry_key, dest, icon, version, category,
        source=source, recolorable=recolorable, color_logo_name=color_name,
    )
    return DownloadResult(registry_key, dest, icon, version, False)


@dataclass(frozen=True)
class ColorLogo:
    prefix: str  # iconify collection: "logos" | "devicon"
    name: str    # icon name within that collection
    svg: bytes


def fetch_color_logo(
    candidates: list[str],
    prefixes: tuple[str, ...] = COLOR_LOGO_PREFIXES,
    session=None,
) -> ColorLogo | None:
    """Return the first full-color logo SVG found for any candidate name.

    Probes each collection (``logos`` then ``devicon``) for each candidate and
    returns the first valid, non-empty SVG. Returns ``None`` if nothing matches
    (the caller then keeps the recolored Simple Icons monochrome logo).
    """
    for prefix in prefixes:
        for name in candidates:
            if not name:
                continue
            url = _ICONIFY_URL.format(prefix=prefix, name=name)
            try:
                svg = _get_bytes(url, session)
                # Iconify returns a tiny 1x1 placeholder SVG for a miss.
                if b"<svg" not in svg or len(svg) < 80:
                    continue
                _validate_color_svg(svg)
            except Exception:  # noqa: BLE001 - color lookup is best-effort; any
                continue        # failure just falls through to the next source.
            return ColorLogo(prefix=prefix, name=name, svg=svg)
    return None


def color_logo_candidates(slug: str, title: str) -> list[str]:
    """Iconify names to try for a full-color logo, icon-only variants first."""
    words = re.findall(r"[a-z0-9]+", title.casefold())
    bases: list[str] = []
    for base in (slug, "".join(words), "-".join(words)):
        if base and base not in bases:
            bases.append(base)
    ordered = [f"{b}-icon" for b in bases] + bases
    seen: set[str] = set()
    out: list[str] = []
    for name in ordered:
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def save_downloaded_metadata(assets_dir: Path, icons: dict[str, dict[str, Any]]) -> None:
    """Persist the downloads ledger (used by the batch color-logo pass)."""
    path = assets_dir / DOWNLOADS_FILE
    payload = {
        "version": 1,
        "note": "Locally cached brand-logo downloads. Review brand guidelines and trademark rights before use.",
        "icons": dict(sorted(icons.items())),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    temp.replace(path)


def record_download(
    assets_dir: Path,
    key: str,
    dest: Path,
    icon: CatalogIcon,
    version: str,
    category: str,
    *,
    source: str = "simple-icons",
    recolorable: bool = True,
    color_logo_name: str = "",
) -> None:
    """Public wrapper so the seed fetcher can persist provenance + brand color."""
    _record_download(
        assets_dir, key, dest, icon, version, category,
        source=source, recolorable=recolorable, color_logo_name=color_logo_name,
    )


def load_downloaded_metadata(assets_dir: Path | None = None) -> dict[str, dict[str, Any]]:
    assets_dir = assets_dir or _ASSETS_DIR
    path = assets_dir / DOWNLOADS_FILE
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise OnlineIconError(f"could not read {path}: {exc}") from exc
    icons = raw.get("icons", {}) if isinstance(raw, dict) else {}
    return icons if isinstance(icons, dict) else {}


def _record_download(
    assets_dir: Path,
    key: str,
    dest: Path,
    icon: CatalogIcon,
    version: str,
    category: str,
    *,
    source: str = "simple-icons",
    recolorable: bool = True,
    color_logo_name: str = "",
) -> None:
    icons = load_downloaded_metadata(assets_dir)
    icons[key] = {
        "file": dest.relative_to(assets_dir).as_posix(),
        "provider": "tech",
        "display_name": icon.title,
        "category": category.strip() or "brand",
        # Full-color logos (source logos/devicon) keep their own colors;
        # only monochrome Simple Icons are recolorable to the brand hex.
        "recolorable": recolorable,
        "brand_hex": f"#{icon.hex}" if icon.hex else "",
        "source": source,
        "color_logo_name": color_logo_name,
        "simple_icons_slug": icon.slug,
        "simple_icons_version": version,
        "source_url": icon.source_url,
        "guidelines_url": icon.guidelines_url,
        "license": {
            "type": icon.license_type,
            "url": icon.license_url,
        },
    }
    save_downloaded_metadata(assets_dir, icons)


def _get_json(url: str, session=None) -> Any:
    data = _request(url, session)
    try:
        return json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise OnlineIconError(f"invalid JSON returned by {url}") from exc


def _get_bytes(url: str, session=None) -> bytes:
    return _request(url, session)


def _request(url: str, session=None) -> bytes:
    import requests  # lazy import so offline registry use has no network side effects

    getter = session.get if session is not None else requests.get
    try:
        response = getter(
            url,
            timeout=20,
            headers={"User-Agent": "archgen/0.1 icon-registry"},
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise OnlineIconError(f"could not fetch {url}: {exc}") from exc
    content = bytes(response.content)
    if not content:
        raise OnlineIconError(f"empty response from {url}")
    return content


def _validate_svg(svg: bytes) -> None:
    if len(svg) > _MAX_SVG_BYTES:
        raise OnlineIconError("downloaded SVG exceeds the 512 KiB safety limit")
    lowered = svg.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise OnlineIconError("downloaded SVG contains a forbidden document/entity declaration")
    try:
        root = ET.fromstring(svg)
    except ET.ParseError as exc:
        raise OnlineIconError("downloaded file is not valid SVG XML") from exc
    if _local_name(root.tag) != "svg" or "viewBox" not in root.attrib:
        raise OnlineIconError("downloaded file is not a viewBox-based SVG")
    for element in root.iter():
        if _local_name(element.tag) in _BANNED_SVG_ELEMENTS:
            raise OnlineIconError("downloaded SVG contains active or embedded content")
        for raw_name, raw_value in element.attrib.items():
            name = _local_name(raw_name)
            value = raw_value.strip().casefold()
            if name.startswith("on"):
                raise OnlineIconError("downloaded SVG contains an event handler")
            if name in {"href", "src"} and (
                value.startswith(("http:", "https:", "//", "data:", "file:"))
            ):
                raise OnlineIconError("downloaded SVG contains an external resource")


def _validate_color_svg(svg: bytes) -> None:
    """Like :func:`_validate_svg` but tolerant of gradients/inline styles that
    multi-color brand logos legitimately use. Still blocks active/embedded
    content, event handlers, and external references."""
    if len(svg) > _MAX_SVG_BYTES:
        raise OnlineIconError("downloaded SVG exceeds the 512 KiB safety limit")
    lowered = svg.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise OnlineIconError("downloaded SVG contains a forbidden document/entity declaration")
    try:
        root = ET.fromstring(svg)
    except ET.ParseError as exc:
        raise OnlineIconError("downloaded file is not valid SVG XML") from exc
    if _local_name(root.tag) != "svg" or "viewBox" not in root.attrib:
        raise OnlineIconError("downloaded file is not a viewBox-based SVG")
    banned = {"script", "foreignobject", "iframe", "object", "embed"}  # <style> allowed here
    for element in root.iter():
        if _local_name(element.tag) in banned:
            raise OnlineIconError("downloaded SVG contains active or embedded content")
        for raw_name, raw_value in element.attrib.items():
            name = _local_name(raw_name)
            value = raw_value.strip().casefold()
            if name.startswith("on"):
                raise OnlineIconError("downloaded SVG contains an event handler")
            if name in {"href", "src"} and value.startswith(
                ("http:", "https:", "//", "data:", "file:")
            ):
                raise OnlineIconError("downloaded SVG contains an external resource")


def _local_name(value: str) -> str:
    return value.rsplit("}", 1)[-1].casefold()


def _normal(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())
