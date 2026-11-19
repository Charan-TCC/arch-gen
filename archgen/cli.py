"""archgen command-line interface.

  archgen render <spec.yaml> -o out.pptx     # spec -> editable pptx (+ layout report)
  archgen render ... --auto-direction --strict # pick best TB/LR; exit 3 on layout errors
  archgen lint   <spec.yaml>                  # validate a spec, list problems
  archgen prompt "<description>"              # print the agent prompt to feed Codex/Claude
  archgen generate "<description>" -o out.pptx # LLM -> spec -> pptx (needs ANTHROPIC_API_KEY)
  archgen icons                               # list valid icon keys
  archgen icons search "Next.js"              # search trusted online catalog
  archgen icons add nextdotjs --key tech.nextjs # cache SVG + rebuild manifest
  archgen schema                              # dump the spec JSON Schema
  archgen doctor                              # report layout/converter capabilities
  archgen qa out.pptx                         # alignment / text-box / arrow-overlap QA
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="archgen", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_render = sub.add_parser("render", help="render a spec file to pptx")
    p_render.add_argument("spec")
    p_render.add_argument("-o", "--out", default="out.pptx")
    p_render.add_argument("--engine", choices=["auto", "graphviz", "simple"], default="auto")
    p_render.add_argument(
        "--auto-direction",
        action="store_true",
        help="lay out both TB and LR and keep the better-fitting one (flow diagrams)",
    )
    p_render.add_argument("--page", choices=["16:9", "free"], default=None,
                          help="override/set the spec's page format")
    p_render.add_argument("--theme", default=None,
                          help="style preset (styles/*.json), overrides the spec's theme")
    p_render.add_argument(
        "--strict",
        action="store_true",
        help="exit non-zero when the layout report contains an error-level finding",
    )

    p_lint = sub.add_parser("lint", help="validate a spec file")
    p_lint.add_argument("spec")

    p_prompt = sub.add_parser("prompt", help="print the agent system+user prompt for a description")
    p_prompt.add_argument("description")

    p_gen = sub.add_parser("generate", help="LLM description -> spec -> pptx (needs ANTHROPIC_API_KEY)")
    p_gen.add_argument("description")
    p_gen.add_argument("-o", "--out", default="out.pptx")
    p_gen.add_argument("--model", default="claude-sonnet-5-5")
    p_gen.add_argument("--page", choices=["16:9", "free"], default=None,
                       help="page format (the user's decision; skips asking)")
    p_gen.add_argument("--theme", default=None, help="style preset (styles/*.json)")
    p_gen.add_argument("--save-spec", default=None, help="also write the generated spec YAML here")

    p_icons = sub.add_parser("icons", help="list, search, or add icon assets")
    icon_sub = p_icons.add_subparsers(dest="icons_cmd")
    p_icon_search = icon_sub.add_parser("search", help="search the online Simple Icons catalog")
    p_icon_search.add_argument("query")
    p_icon_search.add_argument("--limit", type=int, default=10)
    p_icon_add = icon_sub.add_parser("add", help="download an exact Simple Icons title or slug")
    p_icon_add.add_argument("query", help="exact catalog title or slug; run `archgen icons search` first")
    p_icon_add.add_argument("--key", default=None, help="registry key (default: tech.<simple-icons-slug>)")
    p_icon_add.add_argument("--category", default="brand")
    p_icon_add.add_argument("--force", action="store_true", help="replace an existing cached SVG")
    sub.add_parser("schema", help="dump the spec JSON Schema")
    sub.add_parser("doctor", help="report capabilities")
    p_qa = sub.add_parser("qa", help="geometry QA on any .pptx: alignment, text-box size, arrow overlaps")
    p_qa.add_argument("pptx", nargs="+")

    args = parser.parse_args(argv)

    from archgen.spec.io import SpecLoadError

    try:
        if args.cmd == "render":
            return _render(args)
        if args.cmd == "lint":
            return _lint(args)
        if args.cmd == "prompt":
            return _prompt(args)
        if args.cmd == "generate":
            return _generate(args)
        if args.cmd == "icons":
            return _icons(args)
        if args.cmd == "schema":
            return _schema()
        if args.cmd == "qa":
            from archgen.shared import geomcheck
            return geomcheck.main(args.pptx)
        if args.cmd == "doctor":
            return _doctor()
    except SpecLoadError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 1


def _render(args) -> int:
    from archgen.layout.diagnostics import has_errors
    from archgen.render.renderer import Renderer
    from archgen.spec.io import load_spec
    from archgen.spec.lint import lint

    diagram = load_spec(args.spec)
    upd = {}
    if getattr(args, "page", None):
        upd["page"] = args.page
    if getattr(args, "theme", None):
        upd["theme"] = args.theme
    if upd:
        diagram = diagram.model_copy(update=upd)
    errors = lint(diagram)
    if errors:
        print("spec invalid:\n  - " + "\n  - ".join(errors), file=sys.stderr)
        return 2
    renderer = Renderer()
    out = renderer.render(
        diagram, args.out, engine=args.engine, auto_direction=args.auto_direction
    )
    print(f"rendered {out}")
    from archgen.shared import geomcheck
    geo = geomcheck.check_pptx(out)
    if geo:
        print(f"geometry report: {sum(g.severity == 'error' for g in geo)} error(s), "
              f"{sum(g.severity == 'warn' for g in geo)} warning(s)")
        for g in geo:
            print(f"  {g}")
        if args.strict and any(g.severity == "error" for g in geo):
            return 3
    # The layout report is the cheap feedback loop: fix everything it flags
    # BEFORE rasterizing a PNG to look at (see CLAUDE.md §4).
    if renderer.diagnostics:
        print("layout report:")
        for f in renderer.diagnostics:
            print(f"  [{f.severity}] {f.message}")
        if args.strict and has_errors(renderer.diagnostics):
            return 3
    elif diagram.kind.value != "network":
        print("layout report: clean")
    return 0


def _lint(args) -> int:
    from archgen.spec.io import load_spec
    from archgen.spec.lint import lint

    errors = lint(load_spec(args.spec))
    if errors:
        print("INVALID:\n  - " + "\n  - ".join(errors))
        return 2
    print("OK")
    return 0


def _prompt(args) -> int:
    from archgen.agent.prompt import build_system_prompt, build_user_prompt

    print("===== SYSTEM =====")
    print(build_system_prompt())
    print("\n===== USER =====")
    print(build_user_prompt(args.description))
    return 0


def _generate(args) -> int:
    from archgen.agent.generate import anthropic_complete, generate_spec
    from archgen.render.renderer import render_spec

    import os

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print(
            "ANTHROPIC_API_KEY not set. Either set it, or run `archgen prompt \"...\"`, "
            "paste the prompt into your agent (Codex/Claude), save its YAML, and run "
            "`archgen render spec.yaml`.",
            file=sys.stderr,
        )
        return 2
    diagram = generate_spec(args.description, anthropic_complete(args.model), page=args.page)
    if args.theme:
        diagram = diagram.model_copy(update={"theme": args.theme})
    if args.save_spec:
        import yaml

        Path(args.save_spec).write_text(yaml.safe_dump(diagram.model_dump(mode="json"), sort_keys=False))
        print(f"spec -> {args.save_spec}")
    out = render_spec(diagram, args.out)
    print(f"rendered {out}")
    return 0


def _icons(args=None) -> int:
    action = getattr(args, "icons_cmd", None)
    if action in {"search", "add"}:
        from archgen.registry.online import (
            OnlineIconError,
            download_simple_icon,
            fetch_catalog,
            search_catalog,
        )

        try:
            if action == "search":
                version, catalog = fetch_catalog()
                matches = search_catalog(args.query, catalog, limit=args.limit)
                if not matches:
                    print(f"no Simple Icons matches for {args.query!r}")
                    return 1
                print(f"Simple Icons {version} — {len(matches)} match(es)")
                for icon in matches:
                    license_text = icon.license_display
                    print(f"{icon.slug:28} {icon.title}  [license: {license_text}]")
                return 0

            result = download_simple_icon(
                args.query,
                key=args.key,
                category=args.category,
                force=args.force,
            )
            from archgen.registry.fetch import build_manifest

            build_manifest()
            verb = "kept existing" if result.already_present and not args.force else "downloaded"
            print(f"{verb} {result.icon.title} -> {result.key}")
            print(f"asset    : {result.path}")
            print(f"package  : simple-icons {result.package_version}")
            print(f"license  : {result.icon.license_display}")
            if result.icon.guidelines_url:
                print(f"guidelines: {result.icon.guidelines_url}")
            print("manifest : rebuilt; the icon is now reusable offline")
            print("notice   : review the brand's trademark and usage rules before publishing")
            return 0
        except OnlineIconError as exc:
            print(f"icon operation failed: {exc}", file=sys.stderr)
            return 2

    from archgen.registry.resolver import IconResolver

    r = IconResolver()
    for k in r.keys():
        name = r.get(k).display_name if r.exists(k) else "(missing file)"
        print(f"{k:22} {name}")
    return 0


def _schema() -> int:
    from archgen.spec.schema import json_schema_str

    print(json_schema_str())
    return 0


def _doctor() -> int:
    from archgen.registry import emf
    from archgen.registry.resolver import IconResolver

    dot = shutil.which("dot")
    print("layout engine :", "graphviz (dot)" if dot else "simple (pure-Python fallback; `brew install graphviz` for tighter layouts)")
    print("icon converter:", emf.available_engine(), _engine_note(emf.available_engine()))
    r = IconResolver()
    present = sum(1 for k in r.keys() if r.exists(k))
    print(f"icons         : {present}/{len(r.keys())} resolvable")
    return 0


def _engine_note(engine: str) -> str:
    return {
        "inkscape": "(EMF vector — best fidelity/editability)",
        "soffice": "(EMF vector via LibreOffice)",
        "qlmanage": "(PNG via macOS QuickLook — zero-install; install Inkscape for vector EMF)",
        "cairosvg": "(PNG via libcairo)",
    }.get(engine, "")


if __name__ == "__main__":
    raise SystemExit(main())
