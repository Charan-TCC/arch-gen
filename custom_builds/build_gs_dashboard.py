"""Bespoke build of GS Dashboard – Technical Architecture.

All icons are resolved strictly from the archgen icon registry (manifest.json) via IconResolver:
  - azure.devops: Azure SPN & Azure DevOps
  - azure.users: Users (Browser)
  - azure.static-apps: Azure Static Web Apps
  - tech.react: React Frontend (SPA)
  - azure.app-services: Azure App Service
  - tech.python: Python Backend API
  - tech.jwt: Auth Middleware (JWT)
  - azure.sql-database: Azure SQL DB
  - azure.key-vaults: Azure Key Vault
  - azure.application-insights: Application Insights
  - tech.git: Azure Repos
  - generic.pipeline: Azure Pipelines
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

from archgen.registry.resolver import IconResolver
from archgen.registry import emf

R = IconResolver()

# --- Colors ---
BG_PAGE = "FFFFFF"
NAVY_TITLE = "1A365D"       # Dark navy pill
INK_DARK = "0F172A"         # Primary dark headings
INK_MUTED = "475569"        # Gray captions / descriptions
INK_WHITE = "FFFFFF"

# Flows & Accents
BLUE_FLOW = "1D6EE3"        # Application flow / primary blue
PURPLE_FLOW = "8B5CF6"      # Authentication flow / purple
PURPLE_DARK = "5B21B6"      # Dark purple headings
GREEN_ACCENT = "10B981"     # Emerald / Static web apps
GREEN_DARK = "047857"       # Static web apps text
CORAL_ACCENT = "EF4444"     # Azure DevOps / CI/CD coral/red
CORAL_DARK = "DC2626"       # Red headings
ORANGE_SQL = "EA580C"       # SQL DB border
PURPLE_VAULT = "7C3AED"     # Key vault border
TEAL_INSIGHTS = "0D9488"    # App Insights border


def C(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#"))


def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))


def rrect(slide, x, y, w, h, fill=None, line=None, line_pt=1.0, radius=0.06, dashed=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_pt)
        if dashed:
            ln = shp.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    _no_shadow(shp)
    return shp


def circle(slide, cx, cy, r, fill, line=None, line_pt=1.0):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(cx - r), Inches(cy - r), Inches(2 * r), Inches(2 * r)
    )
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_pt)
    _no_shadow(shp)
    return shp


def place_icon(slide, key: str, cx_in: float, cy_in: float, size_in: float):
    """Resolve icon strictly from registry manifest and place centered at (cx_in, cy_in)."""
    e = R.get(key)
    recolor = f"#{e.brand_hex}" if (e.recolorable and e.brand_hex) else None
    png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)
    asp = emf.svg_aspect(e.path)

    if asp >= 1:
        w = size_in
        h = size_in / asp
    else:
        h = size_in
        w = size_in * asp

    slide.shapes.add_picture(
        str(png),
        Inches(cx_in - w / 2),
        Inches(cy_in - h / 2),
        width=Inches(w),
        height=Inches(h),
    )


def text_box(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(2500)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for item in para:
            if len(item) == 4:
                txt, size, bold, color = item
                italic = False
            else:
                txt, size, bold, italic, color = item
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Segoe UI"
            r.font.color.rgb = C(color)
    return box


def draw_arrow(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: str = BLUE_FLOW,
    pt: float = 1.3,
    dashed: bool = False,
    head_end: bool = False,
    tail_end: bool = True,
):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if tail_end:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if head_end:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    cn.shadow.inherit = False
    return cn


def draw_polyline_arrow(
    slide,
    points: list[tuple[float, float]],
    color: str = BLUE_FLOW,
    pt: float = 1.3,
    dashed: bool = False,
    tail_arrow: bool = True,
):
    """Draw orthogonal straight segments with arrowhead at the final point."""
    if len(points) < 2:
        return
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        is_last = (i == len(points) - 2)
        draw_arrow(
            slide,
            x1,
            y1,
            x2,
            y2,
            color=color,
            pt=pt,
            dashed=dashed,
            head_end=False,
            tail_end=(is_last and tail_arrow),
        )


def draw_sequence_badge(slide, cx: float, cy: float, number_or_text: str, radius: float = 0.13):
    """Blue circle with white text for step numbers 1, 2, 3, 4, n."""
    circle(slide, cx, cy, radius, fill=BLUE_FLOW, line=None)
    box = slide.shapes.add_textbox(
        Inches(cx - radius), Inches(cy - radius - 0.02), Inches(2 * radius), Inches(2 * radius + 0.04)
    )
    tf = box.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(number_or_text)
    r.font.size = Pt(8.5)
    r.font.bold = True
    r.font.name = "Segoe UI"
    r.font.color.rgb = C(INK_WHITE)


def build_deck(out_path: str | Path = "out/gs_dashboard.pptx") -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # ---------------------------------------------------------------- TITLE BANNER
    TITLE_BOX = {"x": 0.50, "y": 0.35, "w": 5.40, "h": 0.60}
    rrect(slide, TITLE_BOX["x"], TITLE_BOX["y"], TITLE_BOX["w"], TITLE_BOX["h"], fill=NAVY_TITLE, line=None, radius=0.25)
    text_box(
        slide,
        TITLE_BOX["x"] + 0.22,
        TITLE_BOX["y"] + 0.10,
        TITLE_BOX["w"] - 0.44,
        0.40,
        [[("GS Dashboard – Technical Architecture", 16.0, True, INK_WHITE)]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # ---------------------------------------------------------------- TOP AUTH BOX
    # Azure SPN · Microsoft Entra ID
    AUTH_BOX = {"x": 3.25, "y": 1.15, "w": 3.50, "h": 0.78}
    rrect(slide, AUTH_BOX["x"], AUTH_BOX["y"], AUTH_BOX["w"], AUTH_BOX["h"], fill="FFFFFF", line=PURPLE_FLOW, line_pt=1.2, radius=0.08)
    place_icon(slide, "azure.devops", AUTH_BOX["x"] + 0.44, AUTH_BOX["y"] + 0.39, 0.46)
    text_box(
        slide,
        AUTH_BOX["x"] + 0.78,
        AUTH_BOX["y"] + 0.14,
        AUTH_BOX["w"] - 0.85,
        0.52,
        [
            [("Azure SPN · Microsoft Entra ID", 11.0, True, INK_DARK)],
            [("SSO Login", 9.2, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # ---------------------------------------------------------------- MIDDLE TIER
    # 1. Users (Browser) Card
    USERS_BOX = {"x": 0.50, "y": 2.30, "w": 1.45, "h": 2.40}
    rrect(slide, USERS_BOX["x"], USERS_BOX["y"], USERS_BOX["w"], USERS_BOX["h"], fill="FFFFFF", line="60A5FA", line_pt=1.2, radius=0.08)
    place_icon(slide, "azure.users", USERS_BOX["x"] + USERS_BOX["w"] / 2, USERS_BOX["y"] + 0.85, 0.85)
    text_box(
        slide,
        USERS_BOX["x"] + 0.05,
        USERS_BOX["y"] + 1.68,
        USERS_BOX["w"] - 0.10,
        0.45,
        [[("Users (Browser)", 10.5, True, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 2. Azure Static Web Apps Container
    SWA_BOX = {"x": 2.35, "y": 2.25, "w": 2.50, "h": 2.45}
    rrect(slide, SWA_BOX["x"], SWA_BOX["y"], SWA_BOX["w"], SWA_BOX["h"], fill="FFFFFF", line=GREEN_ACCENT, line_pt=1.4, radius=0.08)
    # Header
    place_icon(slide, "azure.static-apps", SWA_BOX["x"] + 0.25, SWA_BOX["y"] + 0.18, 0.26)
    text_box(
        slide,
        SWA_BOX["x"] + 0.44,
        SWA_BOX["y"] + 0.08,
        SWA_BOX["w"] - 0.50,
        0.24,
        [[("Azure Static Web Apps", 9.0, True, GREEN_DARK)]],
        align=PP_ALIGN.LEFT,
    )
    # Inside React Card
    REACT_CARD = {"x": 2.52, "y": 2.68, "w": 2.14, "h": 1.15}
    rrect(slide, REACT_CARD["x"], REACT_CARD["y"], REACT_CARD["w"], REACT_CARD["h"], fill="FFFFFF", line=GREEN_ACCENT, line_pt=0.9, radius=0.06)
    place_icon(slide, "tech.react", REACT_CARD["x"] + 0.45, REACT_CARD["y"] + 0.57, 0.56)
    text_box(
        slide,
        REACT_CARD["x"] + 0.86,
        REACT_CARD["y"] + 0.15,
        REACT_CARD["w"] - 0.90,
        0.85,
        [
            [("React", 11.5, True, INK_DARK)],
            [("Frontend (SPA)", 9.2, False, INK_DARK)],
            [("MSAL sign-in", 8.2, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )
    # Subtitle at bottom of SWA container
    text_box(
        slide,
        SWA_BOX["x"] + 0.10,
        SWA_BOX["y"] + 2.05,
        SWA_BOX["w"] - 0.20,
        0.28,
        [[("Static web hosting", 9.5, False, GREEN_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 3. Azure App Service Container
    APP_BOX = {"x": 5.30, "y": 2.25, "w": 2.50, "h": 2.45}
    rrect(slide, APP_BOX["x"], APP_BOX["y"], APP_BOX["w"], APP_BOX["h"], fill="FFFFFF", line=BLUE_FLOW, line_pt=1.4, radius=0.08)
    # Header
    place_icon(slide, "azure.app-services", APP_BOX["x"] + 0.25, APP_BOX["y"] + 0.18, 0.26)
    text_box(
        slide,
        APP_BOX["x"] + 0.44,
        APP_BOX["y"] + 0.08,
        APP_BOX["w"] - 0.50,
        0.30,
        [
            [
                ("Azure App Service ", 9.0, True, BLUE_FLOW),
                ("- Managed Identity enabled", 7.8, False, INK_MUTED),
            ]
        ],
        align=PP_ALIGN.LEFT,
    )
    # Python Backend API Card
    PY_CARD = {"x": 5.48, "y": 2.68, "w": 2.14, "h": 0.82}
    rrect(slide, PY_CARD["x"], PY_CARD["y"], PY_CARD["w"], PY_CARD["h"], fill="FFFFFF", line=GREEN_ACCENT, line_pt=0.9, radius=0.06)
    place_icon(slide, "tech.python", PY_CARD["x"] + 0.36, PY_CARD["y"] + 0.41, 0.48)
    text_box(
        slide,
        PY_CARD["x"] + 0.68,
        PY_CARD["y"] + 0.15,
        PY_CARD["w"] - 0.72,
        0.55,
        [
            [("Python", 11.0, True, INK_DARK)],
            [("Backend API", 9.0, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # Auth Middleware Card (JWT)
    JWT_CARD = {"x": 5.48, "y": 3.65, "w": 2.14, "h": 0.82}
    rrect(slide, JWT_CARD["x"], JWT_CARD["y"], JWT_CARD["w"], JWT_CARD["h"], fill="FFFFFF", line=PURPLE_FLOW, line_pt=1.0, radius=0.06, dashed=True)
    place_icon(slide, "tech.jwt", JWT_CARD["x"] + 0.36, JWT_CARD["y"] + 0.41, 0.46)
    text_box(
        slide,
        JWT_CARD["x"] + 0.68,
        JWT_CARD["y"] + 0.15,
        JWT_CARD["w"] - 0.72,
        0.55,
        [
            [("Auth Middleware", 10.2, True, PURPLE_DARK)],
            [("Validate token (JWT)", 8.5, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # 4. Right Service Stack
    STACK_X = 8.50
    STACK_W = 1.85
    STACK_H = 0.68

    # Card 1: Azure SQL DB
    SQL_Y = 2.30
    rrect(slide, STACK_X, SQL_Y, STACK_W, STACK_H, fill="FFFFFF", line=ORANGE_SQL, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.sql-database", STACK_X + 0.34, SQL_Y + 0.34, 0.44)
    text_box(
        slide,
        STACK_X + 0.62,
        SQL_Y + 0.11,
        STACK_W - 0.66,
        0.50,
        [
            [("Azure SQL DB", 9.6, True, INK_DARK)],
            [("Dashboard data", 8.2, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # Card 2: Azure Key Vault
    VAULT_Y = 3.16
    rrect(slide, STACK_X, VAULT_Y, STACK_W, STACK_H, fill="FFFFFF", line=PURPLE_VAULT, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.key-vaults", STACK_X + 0.34, VAULT_Y + 0.34, 0.44)
    text_box(
        slide,
        STACK_X + 0.62,
        VAULT_Y + 0.11,
        STACK_W - 0.66,
        0.50,
        [
            [("Azure Key Vault", 9.6, True, INK_DARK)],
            [("App secrets", 8.2, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # Card 3: Application Insights
    AI_Y = 4.02
    rrect(slide, STACK_X, AI_Y, STACK_W, STACK_H, fill="FFFFFF", line=TEAL_INSIGHTS, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.application-insights", STACK_X + 0.34, AI_Y + 0.34, 0.44)
    text_box(
        slide,
        STACK_X + 0.62,
        AI_Y + 0.11,
        STACK_W - 0.66,
        0.50,
        [
            [("Application Insights", 9.6, True, INK_DARK)],
            [("Monitoring · logs", 8.2, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # App Service to Right Stack Multi-branch Bus Connector
    BUS_X_START = APP_BOX["x"] + APP_BOX["w"]  # 7.80
    BUS_MID_Y = 3.50
    SPINE_X = 8.25

    # App service output line to spine
    draw_arrow(slide, BUS_X_START, BUS_MID_Y, SPINE_X, BUS_MID_Y, color=BLUE_FLOW, pt=1.3, tail_end=False)
    # Bus glyph (three short horizontal blue bars) at 7.96
    GLYPH_X1, GLYPH_X2 = 7.94, 8.08
    draw_arrow(slide, GLYPH_X1, BUS_MID_Y - 0.05, GLYPH_X2, BUS_MID_Y - 0.05, color=BLUE_FLOW, pt=1.2, tail_end=False)
    draw_arrow(slide, GLYPH_X1, BUS_MID_Y, GLYPH_X2, BUS_MID_Y, color=BLUE_FLOW, pt=1.2, tail_end=False)
    draw_arrow(slide, GLYPH_X1, BUS_MID_Y + 0.05, GLYPH_X2, BUS_MID_Y + 0.05, color=BLUE_FLOW, pt=1.2, tail_end=False)

    # Vertical spine
    draw_arrow(slide, SPINE_X, SQL_Y + STACK_H / 2, SPINE_X, AI_Y + STACK_H / 2, color=BLUE_FLOW, pt=1.3, tail_end=False)
    # 3 branches pointing into cards
    draw_arrow(slide, SPINE_X, SQL_Y + STACK_H / 2, STACK_X, SQL_Y + STACK_H / 2, color=BLUE_FLOW, pt=1.3, tail_end=True)
    draw_arrow(slide, SPINE_X, VAULT_Y + STACK_H / 2, STACK_X, VAULT_Y + STACK_H / 2, color=BLUE_FLOW, pt=1.3, tail_end=True)
    draw_arrow(slide, SPINE_X, AI_Y + STACK_H / 2, STACK_X, AI_Y + STACK_H / 2, color=BLUE_FLOW, pt=1.3, tail_end=True)

    # ---------------------------------------------------------------- BOTTOM BAND: AZURE DEVOPS · CI/CD
    DEVOPS_BOX = {"x": 0.50, "y": 4.95, "w": 9.85, "h": 1.55}
    rrect(slide, DEVOPS_BOX["x"], DEVOPS_BOX["y"], DEVOPS_BOX["w"], DEVOPS_BOX["h"], fill="FFFFFF", line=CORAL_ACCENT, line_pt=1.3, radius=0.08)
    # Header
    place_icon(slide, "azure.devops", DEVOPS_BOX["x"] + 0.25, DEVOPS_BOX["y"] + 0.18, 0.26)
    text_box(
        slide,
        DEVOPS_BOX["x"] + 0.44,
        DEVOPS_BOX["y"] + 0.08,
        DEVOPS_BOX["w"] - 0.50,
        0.24,
        [[("Azure DevOps · CI/CD", 9.8, True, CORAL_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # Left Inside Card: Azure Repos
    REPOS_CARD = {"x": 0.80, "y": 5.38, "w": 2.20, "h": 0.90}
    rrect(slide, REPOS_CARD["x"], REPOS_CARD["y"], REPOS_CARD["w"], REPOS_CARD["h"], fill="FFFFFF", line=CORAL_ACCENT, line_pt=0.9, radius=0.06)
    place_icon(slide, "tech.git", REPOS_CARD["x"] + 0.38, REPOS_CARD["y"] + 0.45, 0.48)
    text_box(
        slide,
        REPOS_CARD["x"] + 0.70,
        REPOS_CARD["y"] + 0.10,
        REPOS_CARD["w"] - 0.75,
        0.75,
        [
            [("Azure Repos", 10.2, True, INK_DARK)],
            [("Source code", 8.5, False, INK_MUTED)],
            [("React + Node.js", 8.5, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # Center Inside Card: Azure Pipelines
    PIPELINES_CARD = {"x": 4.15, "y": 5.38, "w": 2.55, "h": 0.90}
    rrect(slide, PIPELINES_CARD["x"], PIPELINES_CARD["y"], PIPELINES_CARD["w"], PIPELINES_CARD["h"], fill="FFFFFF", line=BLUE_FLOW, line_pt=0.9, radius=0.06)
    place_icon(slide, "generic.pipeline", PIPELINES_CARD["x"] + 0.40, PIPELINES_CARD["y"] + 0.45, 0.48)
    text_box(
        slide,
        PIPELINES_CARD["x"] + 0.76,
        PIPELINES_CARD["y"] + 0.10,
        PIPELINES_CARD["w"] - 0.80,
        0.75,
        [
            [("Azure Pipelines", 10.2, True, INK_DARK)],
            [("Build → Test → Deploy", 8.5, False, INK_MUTED)],
            [("Frontend & backend", 8.5, False, INK_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # Right Inside Stack: 2 Deployment Cards
    DEP_W = 2.45
    DEP_H = 0.46
    DEP_X = 7.65

    # Top: Static Web App Deployment
    SWAD_Y = 5.34
    rrect(slide, DEP_X, SWAD_Y, DEP_W, DEP_H, fill="FFFFFF", line=CORAL_ACCENT, line_pt=0.9, radius=0.06)
    text_box(
        slide,
        DEP_X + 0.05,
        SWAD_Y + 0.04,
        DEP_W - 0.10,
        0.38,
        [
            [("Static Web App Deployment", 8.8, True, CORAL_DARK)],
            [("(using deployment token)", 7.6, False, INK_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # Bottom: App Service Deployment
    ASDD_Y = 5.92
    rrect(slide, DEP_X, ASDD_Y, DEP_W, DEP_H, fill="FFFFFF", line=CORAL_ACCENT, line_pt=0.9, radius=0.06)
    text_box(
        slide,
        DEP_X + 0.05,
        ASDD_Y + 0.04,
        DEP_W - 0.10,
        0.38,
        [
            [("App Service Deployment", 8.8, True, CORAL_DARK)],
            [("(using Azure service connection)", 7.6, False, INK_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # Connectors within CI/CD:
    # 1. Repos -> Pipelines
    draw_arrow(slide, REPOS_CARD["x"] + REPOS_CARD["w"], 5.83, PIPELINES_CARD["x"], 5.83, color=BLUE_FLOW, pt=1.2, tail_end=True)
    text_box(slide, 3.00, 5.61, 1.15, 0.22, [[("trigger on push", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.CENTER)

    # 2. Pipelines -> Deployment cards (branching)
    draw_arrow(slide, PIPELINES_CARD["x"] + PIPELINES_CARD["w"], 5.75, DEP_X, SWAD_Y + DEP_H / 2, color=BLUE_FLOW, pt=1.2, tail_end=True)
    text_box(slide, 6.80, 5.40, 0.75, 0.20, [[("Frontend", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.CENTER)

    draw_arrow(slide, PIPELINES_CARD["x"] + PIPELINES_CARD["w"], 5.92, DEP_X, ASDD_Y + DEP_H / 2, color=BLUE_FLOW, pt=1.2, tail_end=True)
    text_box(slide, 6.80, 5.98, 0.75, 0.20, [[("Backend", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.CENTER)

    # ---------------------------------------------------------------- DEPLOYMENT ARROWS (CI/CD -> RUNTIME)
    # 1. Deploy frontend (from Pipelines up and left into Static Web Apps)
    draw_polyline_arrow(
        slide,
        [(4.85, DEVOPS_BOX["y"]), (4.85, 4.80), (3.60, 4.80), (3.60, SWA_BOX["y"] + SWA_BOX["h"])],
        color=BLUE_FLOW,
        pt=1.2,
        dashed=True,
        tail_arrow=True,
    )
    text_box(slide, 3.70, 4.62, 1.15, 0.18, [[("Deploy frontend", 8.0, False, BLUE_FLOW)]], align=PP_ALIGN.LEFT)

    # 2. Deploy backend (straight up into App Service)
    draw_arrow(
        slide,
        6.55,
        DEVOPS_BOX["y"],
        6.55,
        APP_BOX["y"] + APP_BOX["h"],
        color=BLUE_FLOW,
        pt=1.2,
        dashed=True,
        tail_end=True,
    )
    text_box(slide, 6.62, 4.68, 1.10, 0.20, [[("Deploy backend", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.LEFT)

    # ---------------------------------------------------------------- RUNTIME & AUTHENTICATION FLOWS (WITH BADGES)
    # Sequence 1: Login (Users -> React)
    LOGIN_Y = REACT_CARD["y"] + REACT_CARD["h"] / 2  # 3.255
    draw_arrow(
        slide,
        USERS_BOX["x"] + USERS_BOX["w"],
        LOGIN_Y,
        SWA_BOX["x"],
        LOGIN_Y,
        color=BLUE_FLOW,
        pt=1.3,
        tail_end=True,
    )
    draw_sequence_badge(slide, (USERS_BOX["x"] + USERS_BOX["w"] + SWA_BOX["x"]) / 2, LOGIN_Y, "1")
    text_box(slide, (USERS_BOX["x"] + USERS_BOX["w"] + SWA_BOX["x"]) / 2 - 0.35, LOGIN_Y + 0.16, 0.70, 0.20, [[("Login", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.CENTER)

    # Sequence 2: Authenticate (Users -> Entra ID)
    AUTH_CONN_Y = AUTH_BOX["y"] + AUTH_BOX["h"] / 2  # 1.54
    draw_polyline_arrow(
        slide,
        [(USERS_BOX["x"] + 0.72, USERS_BOX["y"]), (USERS_BOX["x"] + 0.72, AUTH_CONN_Y), (AUTH_BOX["x"], AUTH_CONN_Y)],
        color=PURPLE_FLOW,
        pt=1.3,
        dashed=True,
        tail_arrow=True,
    )
    draw_sequence_badge(slide, USERS_BOX["x"] + 1.25, AUTH_CONN_Y, "2")
    text_box(slide, USERS_BOX["x"] + 1.42, AUTH_CONN_Y - 0.22, 1.10, 0.22, [[("Authenticate", 8.5, False, PURPLE_FLOW)]], align=PP_ALIGN.LEFT)

    # Access Token (Entra ID -> React)
    TOKEN_X = 4.05
    draw_arrow(
        slide,
        TOKEN_X,
        AUTH_BOX["y"] + AUTH_BOX["h"],
        TOKEN_X,
        REACT_CARD["y"],
        color=PURPLE_FLOW,
        pt=1.3,
        dashed=True,
        tail_end=True,
    )
    text_box(slide, TOKEN_X + 0.08, 2.38, 1.00, 0.20, [[("Access token", 8.2, False, PURPLE_FLOW)]], align=PP_ALIGN.LEFT)

    # Sequence 3: API calls (React -> Python)
    draw_arrow(
        slide,
        SWA_BOX["x"] + SWA_BOX["w"],
        LOGIN_Y,
        APP_BOX["x"],
        LOGIN_Y,
        color=BLUE_FLOW,
        pt=1.3,
        tail_end=True,
    )
    draw_sequence_badge(slide, (SWA_BOX["x"] + SWA_BOX["w"] + APP_BOX["x"]) / 2, LOGIN_Y, "3")
    text_box(slide, (SWA_BOX["x"] + SWA_BOX["w"] + APP_BOX["x"]) / 2 - 0.35, LOGIN_Y + 0.16, 0.70, 0.20, [[("API calls", 8.2, False, BLUE_FLOW)]], align=PP_ALIGN.CENTER)

    # Sequence 4: Validate token (Entra ID <-> App Service)
    VAL_X = 6.05
    draw_arrow(
        slide,
        VAL_X,
        AUTH_BOX["y"] + AUTH_BOX["h"],
        VAL_X,
        APP_BOX["y"],
        color=PURPLE_FLOW,
        pt=1.3,
        dashed=True,
        head_end=True,
        tail_end=True,
    )
    draw_sequence_badge(slide, VAL_X + 0.18, (AUTH_BOX["y"] + AUTH_BOX["h"] + APP_BOX["y"]) / 2, "4")
    text_box(slide, VAL_X + 0.35, (AUTH_BOX["y"] + AUTH_BOX["h"] + APP_BOX["y"]) / 2 - 0.10, 1.10, 0.22, [[("Validate token", 8.2, False, PURPLE_FLOW)]], align=PP_ALIGN.LEFT)

    # ---------------------------------------------------------------- LEGEND
    LEG_Y = 6.80
    text_box(slide, 0.50, LEG_Y - 0.02, 0.65, 0.25, [[("Legend", 9.2, True, INK_DARK)]], align=PP_ALIGN.LEFT)

    # 1. Application flow (runtime)
    draw_arrow(slide, 1.15, LEG_Y + 0.10, 1.55, LEG_Y + 0.10, color=BLUE_FLOW, pt=1.3, tail_end=True)
    text_box(slide, 1.62, LEG_Y - 0.02, 1.80, 0.25, [[("Application flow (runtime)", 8.2, False, INK_DARK)]], align=PP_ALIGN.LEFT)

    # 2. Authentication flow (SSO / token)
    draw_arrow(slide, 3.55, LEG_Y + 0.10, 3.95, LEG_Y + 0.10, color=PURPLE_FLOW, pt=1.3, dashed=True, tail_end=True)
    text_box(slide, 4.02, LEG_Y - 0.02, 2.15, 0.25, [[("Authentication flow (SSO / token)", 8.2, False, INK_DARK)]], align=PP_ALIGN.LEFT)

    # 3. CI/CD deployment flow
    draw_arrow(slide, 6.35, LEG_Y + 0.10, 6.75, LEG_Y + 0.10, color=BLUE_FLOW, pt=1.3, dashed=True, tail_end=True)
    text_box(slide, 6.82, LEG_Y - 0.02, 1.65, 0.25, [[("CI/CD deployment flow", 8.2, False, INK_DARK)]], align=PP_ALIGN.LEFT)

    # 4. Request sequence
    draw_sequence_badge(slide, 8.70, LEG_Y + 0.10, "n", radius=0.11)
    text_box(slide, 8.90, LEG_Y - 0.02, 1.45, 0.25, [[("Request sequence", 8.2, False, INK_DARK)]], align=PP_ALIGN.LEFT)

    # Ensure output directory exists and save presentation
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"Rendered -> {out}")


if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else "out/gs_dashboard.pptx"
    build_deck(out_file)
