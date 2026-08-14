"""Registry + licensing guardrail tests."""

from __future__ import annotations

import json
import subprocess

import pytest
from PIL import Image, ImageChops

from archgen.registry import emf
from archgen.registry.fetch import build_manifest, discover_official_icons
from archgen.registry.online import (
    CatalogIcon,
    OnlineIconError,
    download_simple_icon,
    load_downloaded_metadata,
    search_catalog,
)
from archgen.registry.resolver import IconResolver


def test_only_simple_icons_are_recolorable():
    """Brand icons (Azure/AWS/GCP) must never be flagged recolorable — the CC BY-ND /
    Azure terms forbid altering them. Only CC0 simple-icons (tech.*) may be tinted."""
    r = IconResolver()
    for key in r.keys():
        entry = r.get(key)
        if entry.recolorable:
            assert entry.provider == "tech", f"{key} is recolorable but not a CC0 tech icon"


def test_all_manifest_icons_resolve():
    r = IconResolver()
    assert all(r.exists(k) for k in r.keys())


def test_svg_aspect_defaults_square_and_reads_viewbox(tmp_path):
    sq = tmp_path / "sq.svg"
    sq.write_text('<svg viewBox="0 0 24 24"></svg>')
    assert abs(emf.svg_aspect(sq) - 1.0) < 1e-6
    wide = tmp_path / "wide.svg"
    wide.write_text('<svg viewBox="0 0 48 24"></svg>')
    assert abs(emf.svg_aspect(wide) - 2.0) < 1e-6


def _fake_svg(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('<svg viewBox="0 0 48 48"></svg>')


def test_discovers_only_service_level_official_icons(tmp_path):
    assets = tmp_path / "icons"
    aws_48 = assets / "aws/Architecture-Service-Icons_01012027/Arch_Compute/48/Arch_Amazon-EC2_48.svg"
    aws_32 = assets / "aws/Architecture-Service-Icons_01012027/Arch_Compute/32/Arch_Amazon-EC2_32.svg"
    aws_resource = assets / "aws/Resource-Icons_01012027/Res_Compute/Res_Amazon-EC2_Instance_48.svg"
    azure = assets / "azure/compute/10023-icon-service-Kubernetes-Services.svg"
    azure_duplicate = assets / "azure/containers/10023-icon-service-Kubernetes-Services.svg"
    gcp = assets / "gcp/Cloud-Run.svg"
    for path in (aws_48, aws_32, aws_resource, azure, azure_duplicate, gcp):
        _fake_svg(path)

    icons = discover_official_icons(assets)

    assert set(icons) == {"aws.ec2", "azure.kubernetes-services", "gcp.cloud-run"}
    assert icons["aws.ec2"]["file"] == aws_48.relative_to(assets).as_posix()
    assert icons["azure.kubernetes-services"]["file"] == azure.relative_to(assets).as_posix()
    assert all(not meta["recolorable"] for meta in icons.values())
    assert all(meta["source"].startswith("official-") for meta in icons.values())


def test_build_manifest_preserves_aliases_across_pack_versions(tmp_path):
    assets = tmp_path / "icons"
    aws = assets / "aws/Architecture-Service-Icons_01012027/Arch_Compute/48/Arch_Amazon-EC2_48.svg"
    azure = assets / "azure/compute/10023-icon-service-Kubernetes-Services.svg"
    gcp = assets / "gcp/Cloud-Run.svg"
    for path in (aws, azure, gcp):
        _fake_svg(path)
    manifest_path = tmp_path / "manifest.json"

    manifest = build_manifest(assets, manifest_path)

    assert manifest["version"] == 2
    assert manifest["icons"]["aws.ec2"]["file"] == aws.relative_to(assets).as_posix()
    assert manifest["icons"]["azure.aks"]["file"] == azure.relative_to(assets).as_posix()
    assert manifest["icons"]["gcp.run"]["file"] == gcp.relative_to(assets).as_posix()
    assert manifest_path.is_file()


class _Response:
    def __init__(self, content: bytes):
        self.content = content

    def raise_for_status(self):
        return None


class _Session:
    def __init__(self, svg: bytes | None = None):
        self.svg = svg or b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z"/></svg>'
        self.urls: list[str] = []

    def get(self, url, **_kwargs):
        self.urls.append(url)
        if url.endswith("/latest"):
            return _Response(json.dumps({"version": "16.21.0"}).encode())
        if url.endswith("/data/simple-icons.json"):
            catalog = {
                "icons": [
                    {
                        "title": "Next.js",
                        "slug": "nextdotjs",
                        "hex": "000000",
                        "source": "https://nextjs.org/",
                        "guidelines": "https://vercel.com/geist/brands",
                        "license": {"type": "MIT", "url": "https://example.test/license"},
                    },
                    {"title": "Nextcloud", "slug": "nextcloud", "hex": "0082C9"},
                ]
            }
            return _Response(json.dumps(catalog).encode())
        if url.endswith("/icons/nextdotjs.svg"):
            return _Response(self.svg)
        raise AssertionError(f"unexpected URL: {url}")


def test_online_catalog_search_prefers_exact_and_prefix_matches():
    catalog = [
        CatalogIcon("Nextcloud", "nextcloud"),
        CatalogIcon("Next.js", "nextdotjs"),
        CatalogIcon("NestJS", "nestjs"),
    ]
    assert search_catalog("Next.js", catalog)[0].slug == "nextdotjs"
    assert search_catalog("next", catalog)[0].slug in {"nextcloud", "nextdotjs"}


def test_downloaded_simple_icon_is_cached_and_added_to_manifest(tmp_path):
    assets = tmp_path / "icons"
    manifest_path = tmp_path / "manifest.json"
    result = download_simple_icon(
        "nextdotjs",
        key="tech.nextjs",
        category="frontend",
        assets_dir=assets,
        session=_Session(),
    )

    assert result.path.is_file()
    assert result.package_version == "16.21.0"
    metadata = load_downloaded_metadata(assets)["tech.nextjs"]
    assert metadata["simple_icons_slug"] == "nextdotjs"
    assert metadata["license"]["type"] == "MIT"

    manifest = build_manifest(assets, manifest_path)
    entry = manifest["icons"]["tech.nextjs"]
    assert entry["file"] == "tech/nextjs.svg"
    assert entry["source"] == "simple-icons"
    assert entry["simple_icons_version"] == "16.21.0"

    resolver = IconResolver(manifest_path, assets)
    assert resolver.exists("tech.nextjs")


def test_online_download_rejects_active_svg(tmp_path):
    unsafe = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><script>alert(1)</script></svg>'
    with pytest.raises(OnlineIconError, match="active or embedded"):
        download_simple_icon("nextdotjs", assets_dir=tmp_path / "icons", session=_Session(unsafe))
    assert not (tmp_path / "icons/tech/nextdotjs.svg").exists()


def test_online_download_does_not_claim_unverified_existing_file(tmp_path):
    assets = tmp_path / "icons"
    existing = assets / "tech/nextjs.svg"
    _fake_svg(existing)
    with pytest.raises(OnlineIconError, match="without matching Simple Icons provenance"):
        download_simple_icon(
            "nextdotjs",
            key="tech.nextjs",
            assets_dir=assets,
            session=_Session(),
        )
    assert load_downloaded_metadata(assets) == {}


def test_converter_falls_back_when_vector_engine_rejects_svg(tmp_path, monkeypatch):
    svg = tmp_path / "official.svg"
    _fake_svg(svg)
    cache = tmp_path / "cache"
    monkeypatch.setattr(emf, "_CACHE_DIR", cache)
    monkeypatch.setattr(emf, "_candidate_engines", lambda _prefer: ["soffice", "cairosvg"])

    def reject(_svg, _out):
        raise subprocess.CalledProcessError(77, ["soffice"], stderr="conversion rejected")

    def rasterize(_svg_bytes, out):
        out.write_bytes(b"png")

    monkeypatch.setattr(emf, "_soffice", reject)
    monkeypatch.setattr(emf, "_cairosvg", rasterize)

    out = emf.to_picture(svg)

    assert out.suffix == ".png"
    assert out.read_bytes() == b"png"


def test_concurrent_conversion_of_same_icon_is_safe(tmp_path, monkeypatch):
    """Regression for the emf.py temp-file collision: two callers converting the
    SAME icon at the same time must never write/unlink a shared filename. Each
    conversion now runs in a private temp dir and publishes with `os.replace`,
    so both racers should land on the identical, complete cached file."""
    import threading

    svg = tmp_path / "official.svg"
    _fake_svg(svg)
    cache = tmp_path / "cache"
    monkeypatch.setattr(emf, "_CACHE_DIR", cache)
    monkeypatch.setattr(emf, "_candidate_engines", lambda _prefer: ["cairosvg"])

    def slow_rasterize(_svg_bytes, out):
        import time

        time.sleep(0.02)  # widen the race window
        out.write_bytes(b"png-bytes")

    monkeypatch.setattr(emf, "_cairosvg", slow_rasterize)

    results: list[Path] = []
    errors: list[Exception] = []

    def worker():
        try:
            results.append(emf.to_picture(svg))
        except Exception as exc:  # noqa: BLE001
            errors.append(exc)

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert not errors
    assert len(results) == 8
    assert len(set(results)) == 1  # every thread resolved to the same cache path
    out = results[0]
    assert out.is_file()
    assert out.read_bytes() == b"png-bytes"
    # No leftover per-attempt temp dirs — each is rmtree'd in `finally`.
    assert not list(cache.glob("_tmp_*"))


def test_quicklook_thumbnail_canvas_is_trimmed_and_refit(tmp_path):
    path = tmp_path / "thumbnail.png"
    image = Image.new("RGBA", (512, 512), (255, 255, 255, 255))
    image.paste((140, 79, 255, 255), (0, 0, 64, 64))
    image.save(path)

    emf._trim_quicklook_canvas(path)

    result = Image.open(path).convert("RGB")
    diff = ImageChops.difference(result, Image.new("RGB", result.size, "white"))
    assert diff.getbbox() == (0, 0, 512, 512)
    assert result.getpixel((256, 256)) == (140, 79, 255)
