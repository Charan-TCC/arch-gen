"""Populate the icon registry and (re)build the manifest.

Three concerns, kept separate:

* :func:`fetch_simple_icons` — download real CC0 generic tech logos (kafka,
  postgres, ...) from the jsDelivr CDN mirror of simple-icons.
* ``archgen icons search/add`` — discover additional Simple Icons online and
  cache their SVG + provenance for future offline renders.
* :func:`seed_placeholders` — generate clean, provider-colored placeholder SVGs
  for the curated compatibility keys when an official icon is unavailable.
* :func:`discover_official_icons` — scan installed AWS/Azure/GCP service packs and
  turn their filenames into stable canonical keys.
* :func:`build_manifest` — merge discovered services, compatibility aliases,
  generic icons, and CC0 tech icons into ``manifest.json``.

Run:  ``python -m archgen.registry.fetch``
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from archgen.registry.online import load_downloaded_metadata

_REGISTRY_DIR = Path(__file__).resolve().parent
_ASSETS_DIR = _REGISTRY_DIR.parent.parent / "assets" / "icons"
_MANIFEST = _REGISTRY_DIR / "manifest.json"

# provider -> brand hex (used to tint placeholders and drive the palette default)
PROVIDER_HEX = {
    "azure": "#0078D4",
    "aws": "#FF9900",
    "gcp": "#4285F4",
    "tech": "#333333",
    "generic": "#5B6470",
}

# Curated demo set of cloud services: key -> (display_name, category, abbrev-glyph)
CLOUD_SEED = {
    "azure.aks": ("Azure Kubernetes Service", "compute", "AKS"),
    "azure.appservice": ("Azure App Service", "compute", "APP"),
    "azure.sql": ("Azure SQL Database", "database", "SQL"),
    "azure.blob": ("Azure Blob Storage", "storage", "BLOB"),
    "aws.lambda": ("AWS Lambda", "compute", "λ"),
    "aws.ec2": ("Amazon EC2", "compute", "EC2"),
    "aws.rds": ("Amazon RDS", "database", "RDS"),
    "aws.s3": ("Amazon S3", "storage", "S3"),
    "gcp.gke": ("Google Kubernetes Engine", "compute", "GKE"),
    "gcp.run": ("Cloud Run", "compute", "RUN"),
    "gcp.spanner": ("Cloud Spanner", "database", "SPN"),
}

# Generic architecture concepts, drawn from Icons8 Flat Color Icons (MIT).
# These are multi-color flat icons — registered non-recolorable so their own
# colors are preserved (never tinted to the flat generic gray).
# key -> (flat-color-icons name, display_name, category)
GENERIC_SEED = {
    "generic.user": ("businessman", "User", "endpoint"),
    "generic.client": ("multiple_devices", "Client", "endpoint"),
    "generic.mobile": ("smartphone_tablet", "Mobile App", "endpoint"),
    "generic.internet": ("globe", "Internet", "network"),
    "generic.identity": ("key", "Identity", "security"),
    "generic.security": ("privacy", "Security", "security"),
    "generic.firewall": ("safe", "Firewall", "security"),
    "generic.database": ("database", "Database", "database"),
    "generic.storage": ("filing_cabinet", "Storage", "storage"),
    "generic.folder": ("opened_folder", "File Store", "storage"),
    "generic.document": ("document", "Document", "storage"),
    "generic.queue": ("serial_tasks", "Queue", "messaging"),
    "generic.workflow": ("workflow", "Workflow", "integration"),
    "generic.pipeline": ("flow_chart", "Pipeline", "integration"),
    "generic.function": ("process", "Function", "compute"),
    "generic.container": ("package", "Container", "compute"),
    "generic.service": ("services", "Service", "compute"),
    "generic.api": ("command_line", "API", "integration"),
    "generic.gateway": ("tree_structure", "Gateway", "network"),
    "generic.loadbalancer": ("parallel_tasks", "Load Balancer", "network"),
    "generic.cache": ("synchronize", "Cache", "cache"),
    "generic.monitoring": ("combo_chart", "Monitoring", "observability"),
    "generic.analytics": ("bar_chart", "Analytics", "analytics"),
    "generic.search": ("search", "Search", "service"),
    "generic.schedule": ("alarm_clock", "Scheduler", "integration"),
    "generic.settings": ("settings", "Config", "management"),
    "generic.support": ("headset", "Support", "service"),
    "generic.email": ("invite", "Email", "integration"),
    "generic.organization": ("organization", "Organization", "management"),
}

# Stable aliases for the small public catalog that shipped in v0.1. Auto-discovery
# exposes the rest of each installed service pack; these aliases preserve existing
# specs and choose a deterministic official file where a service has duplicates.
OFFICIAL_FILE = {
    "azure.aks": "azure/compute/10023-icon-service-Kubernetes-Services.svg",
    "azure.appservice": "azure/app services/10035-icon-service-App-Services.svg",
    "azure.sql": "azure/databases/10130-icon-service-SQL-Database.svg",
    "azure.blob": "azure/storage/10086-icon-service-Storage-Accounts.svg",
    "aws.lambda": "aws/Architecture-Service-Icons_04302026/Arch_Compute/48/Arch_AWS-Lambda_48.svg",
    "aws.ec2": "aws/Architecture-Service-Icons_04302026/Arch_Compute/48/Arch_Amazon-EC2_48.svg",
    "aws.rds": "aws/Architecture-Service-Icons_04302026/Arch_Databases/48/Arch_Amazon-RDS_48.svg",
    "aws.s3": "aws/Architecture-Service-Icons_04302026/Arch_Storage/48/Arch_Amazon-Simple-Storage-Service_48.svg",
    "gcp.gke": "gcp/Google-Kubernetes-Engine.svg",
    "gcp.run": "gcp/Cloud-Run.svg",
    "gcp.spanner": "gcp/Cloud-Spanner.svg",
}

# Backwards-compatible name for callers that imported the old extension point.
REAL_FILE = OFFICIAL_FILE

# Compatibility key -> auto-discovered catalog key. Unlike OFFICIAL_FILE, these do
# not contain a pack version or category path, so aliases survive future pack drops.
OFFICIAL_KEY_ALIAS = {
    "azure.aks": "azure.kubernetes-services",
    "azure.appservice": "azure.app-services",
    "azure.sql": "azure.sql-database",
    "azure.blob": "azure.storage-accounts",
    "aws.lambda": "aws.lambda",
    "aws.ec2": "aws.ec2",
    "aws.rds": "aws.rds",
    "aws.s3": "aws.simple-storage-service",
    "gcp.gke": "gcp.kubernetes-engine",
    "gcp.run": "gcp.cloud-run",
    "gcp.spanner": "gcp.cloud-spanner",
}

# Curated generic tech logos to pull from simple-icons (CC0 artwork).
# key -> (simple-icons slug, display_name, category)
TECH_SEED = {
    "tech.kafka": ("apachekafka", "Apache Kafka", "messaging"),
    "tech.postgres": ("postgresql", "PostgreSQL", "database"),
    "tech.redis": ("redis", "Redis", "cache"),
    "tech.nginx": ("nginx", "NGINX", "network"),
    "tech.react": ("react", "React", "frontend"),
    "tech.docker": ("docker", "Docker", "compute"),
}

_CDN = "https://cdn.jsdelivr.net/npm/simple-icons@latest/icons/{slug}.svg"
# Icons8 Flat Color Icons (MIT). Served from the project's GitHub repo.
_FCI_CDN = "https://raw.githubusercontent.com/icons8/flat-color-icons/master/svg/{name}.svg"


def fetch_flat_color_icons(seed: dict | None = None) -> list[str]:
    """Download multi-color generic concept icons (Icons8 Flat Color, MIT).

    Written into ``assets/icons/generic/``. On any failure we fall back to the
    provider-colored placeholder so a render never breaks offline.
    """
    import requests  # lazy import

    seed = seed or GENERIC_SEED
    out_dir = _ASSETS_DIR / "generic"
    out_dir.mkdir(parents=True, exist_ok=True)
    fetched: list[str] = []
    for key, (name, _display, _cat) in seed.items():
        local = key.split(".", 1)[1]
        dest = out_dir / f"{local}.svg"
        try:
            resp = requests.get(_FCI_CDN.format(name=name), timeout=15)
            resp.raise_for_status()
            if b"<svg" not in resp.content:
                raise ValueError("response was not an SVG")
            dest.write_bytes(resp.content)
            fetched.append(key)
        except Exception as exc:  # noqa: BLE001 - offline is acceptable; placeholder covers it
            print(f"  ! could not fetch {key} ({name}): {exc}; will use placeholder")
            _write_placeholder(dest, PROVIDER_HEX["generic"], local[:4].upper())
            fetched.append(key)
    return fetched


def apply_color_logos(assets_dir: Path | None = None) -> list[tuple[str, str, str]]:
    """Upgrade cached ``tech.*`` logos to full color where a color set has them.

    For every tech icon in the downloads ledger, try Gil Barbara ``logos`` then
    ``devicon`` (via :func:`fetch_color_logo`). On a hit, overwrite the SVG with
    the color artwork and mark it non-recolorable so its own colors are kept.
    Misses keep the recolored monochrome Simple Icons logo.
    """
    from archgen.registry.online import (
        color_logo_candidates,
        fetch_color_logo,
        load_downloaded_metadata,
        save_downloaded_metadata,
    )

    assets_dir = assets_dir or _ASSETS_DIR
    meta = load_downloaded_metadata(assets_dir)
    colored: list[tuple[str, str, str]] = []
    for key, entry in sorted(meta.items()):
        if not isinstance(entry, dict) or not re.fullmatch(r"tech\.[a-z0-9][a-z0-9-]*", key):
            continue
        slug = str(entry.get("simple_icons_slug") or key.split(".", 1)[1])
        title = str(entry.get("display_name") or key)
        logo = fetch_color_logo(color_logo_candidates(slug, title))
        if logo is None:
            continue
        dest = assets_dir / "tech" / f"{key.split('.', 1)[1]}.svg"
        dest.write_bytes(logo.svg)
        updated = dict(entry)
        updated["source"] = logo.prefix
        updated["recolorable"] = False
        updated["color_logo_name"] = logo.name
        meta[key] = updated
        colored.append((key, logo.prefix, logo.name))
    save_downloaded_metadata(assets_dir, meta)
    return colored


def fetch_simple_icons(seed: dict | None = None) -> list[str]:
    import requests  # lazy import

    from archgen.registry.online import fetch_catalog, record_download

    seed = seed or TECH_SEED
    out_dir = _ASSETS_DIR / "tech"
    out_dir.mkdir(parents=True, exist_ok=True)

    # Pull the catalog once to attach each logo's official brand color +
    # provenance, so recolorable tech icons render in brand color (not gray).
    try:
        version, catalog = fetch_catalog()
        by_slug = {icon.slug: icon for icon in catalog}
    except Exception as exc:  # noqa: BLE001 - offline is fine; we just skip color metadata
        print(f"  ! could not load Simple Icons catalog ({exc}); tech icons stay uncolored")
        version, by_slug = "", {}

    fetched: list[str] = []
    for key, (slug, _name, cat) in seed.items():
        dest = out_dir / f"{key.split('.', 1)[1]}.svg"
        try:
            resp = requests.get(_CDN.format(slug=slug), timeout=15)
            resp.raise_for_status()
            dest.write_bytes(resp.content)
            catalog_icon = by_slug.get(slug)
            if catalog_icon is not None and version:
                record_download(_ASSETS_DIR, key, dest, catalog_icon, version, cat)
            fetched.append(key)
        except Exception as exc:  # noqa: BLE001 - offline is acceptable; placeholder covers it
            print(f"  ! could not fetch {key} ({slug}): {exc}; will use placeholder")
            _write_placeholder(dest, PROVIDER_HEX["tech"], slug[:4].upper())
            fetched.append(key)
    return fetched


def _placeholder_rel(key: str) -> str:
    provider = key.split(".", 1)[0]
    return f"{provider}/{key.split('.', 1)[1]}.svg"


def resolve_rel_file(key: str, assets_dir: Path | None = None) -> str | None:
    """Relative path to use for a key: real pack file if present, else placeholder."""
    assets_dir = assets_dir or _ASSETS_DIR
    real = OFFICIAL_FILE.get(key)
    if real and (assets_dir / real).is_file():
        return real
    ph = _placeholder_rel(key)
    return ph if (assets_dir / ph).is_file() else None


def seed_placeholders(seed: dict | None = None, assets_dir: Path | None = None) -> list[str]:
    """Write a placeholder SVG for any cloud key that has no real pack file yet."""
    seed = seed or CLOUD_SEED
    assets_dir = assets_dir or _ASSETS_DIR
    discovered = discover_official_icons(assets_dir)
    made: list[str] = []
    for key, (_name, _cat, glyph) in seed.items():
        discovered_key = OFFICIAL_KEY_ALIAS.get(key, key)
        if discovered_key in discovered:
            continue
        real = OFFICIAL_FILE.get(key)
        if real and (assets_dir / real).is_file():
            continue  # real official icon present — never overwrite it
        provider = key.split(".", 1)[0]
        dest = assets_dir / provider / f"{key.split('.', 1)[1]}.svg"
        dest.parent.mkdir(parents=True, exist_ok=True)
        _write_placeholder(dest, PROVIDER_HEX.get(provider, PROVIDER_HEX["generic"]), glyph)
        made.append(key)
    return made


def _slugify(value: str) -> str:
    """Turn a provider filename fragment into a stable registry-key component."""
    value = value.casefold().replace("&", " and ").replace("+", " plus ")
    return re.sub(r"^-|-$", "", re.sub(r"[^a-z0-9]+", "-", value))


def _canonical_service_key(provider: str, name: str) -> str:
    slug = _slugify(name)
    prefixes = {
        "aws": ("amazon-", "aws-"),
        "azure": ("azure-",),
        "gcp": ("google-",),
    }
    for prefix in prefixes.get(provider, ()):
        if slug.startswith(prefix):
            slug = slug[len(prefix):]
            break
    return f"{provider}.{slug}"


def _display_name(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("_", " ").replace("-", " ")).strip()


def _entry(rel: Path, provider: str, display_name: str, category: str) -> dict:
    return {
        "file": rel.as_posix(),
        "provider": provider,
        "display_name": display_name,
        "category": category,
        "recolorable": False,
        "source": f"official-{provider}",
    }


def discover_official_icons(assets_dir: Path | None = None) -> dict[str, dict]:
    """Discover service-level icons from installed official provider packs.

    AWS publishes the same service artwork at several sizes plus separate resource
    and group catalogs. Only 48px architecture-service SVGs are indexed. Azure's
    pack repeats some services across categories; sorted traversal and ``setdefault``
    make the chosen file deterministic. GCP's service pack is flat and unique.
    """
    assets_dir = assets_dir or _ASSETS_DIR
    icons: dict[str, dict] = {}

    aws_root = assets_dir / "aws"
    aws_paths = sorted(aws_root.glob("Architecture-Service-Icons_*/*/48/*.svg"))
    for path in aws_paths:
        stem = re.sub(r"^Arch_", "", path.stem)
        stem = re.sub(r"_48$", "", stem)
        key = _canonical_service_key("aws", stem)
        category_dir = path.parent.parent.name
        category = _slugify(re.sub(r"^Arch_", "", category_dir))
        rel = path.relative_to(assets_dir)
        icons.setdefault(key, _entry(rel, "aws", _display_name(stem), category))

    azure_root = assets_dir / "azure"
    azure_paths = sorted(path for path in azure_root.rglob("*.svg") if path.parent != azure_root)
    for path in azure_paths:
        stem = re.sub(r"^\d+-icon-service-", "", path.stem, flags=re.IGNORECASE)
        key = _canonical_service_key("azure", stem)
        category = _slugify(path.relative_to(azure_root).parts[0])
        rel = path.relative_to(assets_dir)
        icons.setdefault(key, _entry(rel, "azure", _display_name(stem), category))

    gcp_root = assets_dir / "gcp"
    for path in sorted(gcp_root.glob("*.svg")):
        key = _canonical_service_key("gcp", path.stem)
        rel = path.relative_to(assets_dir)
        icons.setdefault(key, _entry(rel, "gcp", _display_name(path.stem), "service"))

    return icons


def build_manifest(
    assets_dir: Path | None = None,
    manifest_path: Path | None = None,
) -> dict:
    assets_dir = assets_dir or _ASSETS_DIR
    manifest_path = manifest_path or _MANIFEST
    icons = discover_official_icons(assets_dir)

    for key, (name, cat, _glyph) in CLOUD_SEED.items():
        provider = key.split(".", 1)[0]
        official = icons.get(OFFICIAL_KEY_ALIAS.get(key, key))
        if official is not None:
            icons[key] = {
                **official,
                "display_name": name,
                "category": cat,
            }
            continue
        rel = resolve_rel_file(key, assets_dir)
        if rel:
            icons[key] = {
                "file": rel,
                "provider": provider,
                "display_name": name,
                "category": cat,
                "recolorable": False,  # placeholders/brand icons: do not recolor
                "source": f"official-{provider}" if OFFICIAL_FILE.get(key) == rel else "placeholder",
            }

    for key, (_slug, name, cat) in TECH_SEED.items():
        rel = f"tech/{key.split('.', 1)[1]}.svg"
        if (assets_dir / rel).is_file():
            icons[key] = {
                "file": rel,
                "provider": "tech",
                "display_name": name,
                "category": cat,
                "recolorable": True,  # simple-icons artwork is CC0 -> safe to tint
                "source": "simple-icons",
            }

    for key, (_name, display, cat) in GENERIC_SEED.items():
        rel = f"generic/{key.split('.', 1)[1]}.svg"
        if (assets_dir / rel).is_file():
            icons[key] = {
                "file": rel,
                "provider": "generic",
                "display_name": display,
                "category": cat,
                # Multi-color flat icons: preserve their own colors (never tint).
                "recolorable": False,
                "source": "flat-color-icons",
                "license": {"name": "MIT", "holder": "Icons8"},
            }

    # Explicit ``archgen icons add`` downloads are persisted separately from the
    # curated seed so rebuilding the manifest never forgets locally added brands.
    assets_root = assets_dir.resolve()
    for key, meta in load_downloaded_metadata(assets_dir).items():
        if (
            not isinstance(meta, dict)
            or meta.get("source") not in ("simple-icons", "logos", "devicon")
            or not re.fullmatch(r"tech\.[a-z0-9][a-z0-9-]*", key)
        ):
            continue
        rel = Path(str(meta.get("file", "")))
        candidate = (assets_dir / rel).resolve()
        if not rel.parts or not candidate.is_relative_to(assets_root) or not candidate.is_file():
            continue
        source = str(meta.get("source"))
        icons[key] = {
            "file": candidate.relative_to(assets_root).as_posix(),
            "provider": "tech",
            "display_name": str(meta.get("display_name", key)),
            "category": str(meta.get("category", "brand")),
            # Color logos (logos/devicon) keep their own colors; only the
            # monochrome Simple Icons fallback is recolored to the brand hex.
            "recolorable": source == "simple-icons",
            "brand_hex": str(meta.get("brand_hex", "")),
            "source": source,
            "color_logo_name": str(meta.get("color_logo_name", "")),
            "simple_icons_slug": str(meta.get("simple_icons_slug", "")),
            "simple_icons_version": str(meta.get("simple_icons_version", "")),
            "source_url": str(meta.get("source_url", "")),
            "guidelines_url": str(meta.get("guidelines_url", "")),
            "license": meta.get("license", {}),
        }

    official_counts = {
        provider: sum(
            1 for meta in icons.values()
            if meta["provider"] == provider and meta.get("source") == f"official-{provider}"
        )
        for provider in ("aws", "azure", "gcp")
    }
    manifest = {
        "version": 2,
        "note": "Official AWS/Azure/GCP service packs are auto-discovered when present; "
                "curated compatibility keys fall back to generated placeholders; "
                "additional Simple Icons downloads are cached with provenance.",
        "official_counts": official_counts,
        "icons": dict(sorted(icons.items())),
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")
    return manifest


def _write_placeholder(dest: Path, hex_color: str, glyph: str) -> None:
    """A clean rounded-square placeholder with a short glyph — good enough to
    prove layout/rendering; visually distinct per provider color."""
    glyph = glyph[:4]
    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="100" height="100">
  <rect x="6" y="6" width="88" height="88" rx="14" ry="14" fill="{hex_color}"/>
  <text x="50" y="50" font-family="Arial, Helvetica, sans-serif" font-size="26"
        font-weight="700" fill="#FFFFFF" text-anchor="middle" dominant-baseline="central">{_xml(glyph)}</text>
</svg>
"""
    dest.write_text(svg, encoding="utf-8")


def _xml(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--no-download",
        action="store_true",
        help="rebuild from local assets without downloading simple-icons",
    )
    args = parser.parse_args(argv)

    print("Seeding placeholder cloud/generic icons ...")
    made = seed_placeholders()
    print(f"  created {len(made)} placeholders")
    if not args.no_download:
        print("Fetching CC0 tech logos from simple-icons (jsDelivr) ...")
        fetched = fetch_simple_icons()
        print(f"  fetched {len(fetched)} tech logos")
        print("Fetching color generic icons from Flat Color Icons (MIT) ...")
        generic = fetch_flat_color_icons()
        print(f"  fetched {len(generic)} generic color icons")
        print("Upgrading tech logos to full color (Gil Barbara logos -> Devicon) ...")
        colored = apply_color_logos()
        print(f"  colorized {len(colored)} tech logos")
    print("Building manifest ...")
    manifest = build_manifest()
    counts = manifest["official_counts"]
    print(
        f"  official services: AWS {counts['aws']}, Azure {counts['azure']}, "
        f"GCP {counts['gcp']}"
    )
    print(f"  manifest: {len(manifest['icons'])} icons -> {_MANIFEST}")


if __name__ == "__main__":
    main()
