"""Bespoke build generating the complete architecture presentation deck.

Deck Structure:
- Slide 1: Functional / Workflow Architecture (Enterprise Data Platform & Web Application Architecture)
  - Full pipeline: Ingestion, DP360 / DQ360 Governance Checks, ML model, Google enrichments,
    Databricks tables (ml_result, feedback, agentic), SQL Server, Web App screens with browser icons,
    AI Validation on-demand trigger, high-visibility Auto-logs bus, and horizontal bottom legend.
- Slide 2: Technical Architecture (Clean, Minimal, High-Level)
  - Focused strictly on the core system components without clutter:
    1. Frontend: React application hosted on Azure Static Web Apps, with SSO enabled.
    2. Backend: Hono (Node.js) API hosted on Azure App Service, secured with bearer token authentication.
    3. Database: SQL Server, accessible only from the backend.
    4. Data platform: The backend connects to Azure Databricks to trigger jobs and run queries.
    5. Deployment: Azure DevOps CI/CD pipelines.
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

# ==========================================
# PALETTE (Executive, Restrained, Clean)
# ==========================================
BG_PAGE = "FFFFFF"

# Typography
TEXT_DARK = "0F172A"       # Primary titles / card headings
TEXT_MUTED = "475569"      # Secondary descriptions
TEXT_LIGHT = "64748B"      # Captions & tags
TEXT_WHITE = "FFFFFF"

# Layer Zones
ZONE_FE_BG = "FAFAFE"      # Soft slate-lavender
ZONE_FE_LINE = "818CF8"    # Soft indigo border (dashed)

ZONE_BE_BG = "F8FAFC"      # Soft slate-blue
ZONE_BE_LINE = "60A5FA"    # Soft blue border (dashed)

ZONE_GOV_BG = "F8FAFC"     # Soft neutral/slate-gray
ZONE_GOV_LINE = "94A3B8"   # Subtle slate border

# Component Cards
CARD_BG = "FFFFFF"
CARD_LINE = "1E293B"       # Crisp dark border matching minimal style
CARD_LINE_LIGHT = "CBD5E1" # Slate 300 border for Slide 1

# Specialized Containers
SQL_BG = "FFFBEB"          # Warm champagne/amber
SQL_LINE = "F59E0B"        # Amber border

DB_TABLE_BG = "F5F3FF"     # Soft purple tint
DB_TABLE_LINE = "8B5CF6"   # Purple border

AI_VAL_BG = "FEF2F2"       # Soft red/coral tint
AI_VAL_LINE = "F87171"

# Slide 2 Soft Container Fills (Minimal executive style)
SWA_BG = "DCFCE7"          # Soft light green
SWA_LINE = "86EFAC"

APP_BG = "DBEAFE"          # Soft light blue
APP_LINE = "93C5FD"

STORAGE_BG = "FEF3C7"      # Soft light yellow / amber
STORAGE_LINE = "FCD34D"

# Connectors
COLOR_SCHEDULED = "1E293B" # Dark slate for scheduled runs & main pipeline
COLOR_DEMAND = "EA580C"    # Warm orange/coral for user on demand & feedback
COLOR_LOGS = "0284C7"      # High-contrast vibrant ocean blue for auto logs
ARROW_COLOR = "1E293B"     # Clean dark slate for technical architecture arrows

def C(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#"))

def _no_shadow(shp):
    """Enforce zero drop shadow for clean enterprise look."""
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
    """Place icon centered at (cx_in, cy_in)."""
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
    pic = slide.shapes.add_picture(
        str(png),
        Inches(cx_in - w / 2),
        Inches(cy_in - h / 2),
        width=Inches(w),
        height=Inches(h),
    )
    _no_shadow(pic)
    return pic

def text_box(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(3000)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for item in para:
            if len(item) == 5:
                txt, size, bold, italic, color = item
            elif len(item) == 4:
                txt, size, bold, color = item
                italic = False
            else:
                txt, size, bold = item
                italic = False
                color = TEXT_DARK
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Segoe UI"
            r.font.color.rgb = C(color)
    _no_shadow(box)
    return box

def draw_arrow(slide, x1, y1, x2, y2, color=COLOR_SCHEDULED, pt=1.2, dashed=False, dash_style="dash", double_ended=False):
    """Straight arrow from (x1, y1) to (x2, y2)."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash_style}))
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if double_ended:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    _no_shadow(cn)
    return cn

def draw_line(slide, x1, y1, x2, y2, color=COLOR_SCHEDULED, pt=1.2, dashed=False, dash_style="dash"):
    """Plain line without arrowhead."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash_style}))
    _no_shadow(cn)
    return cn

def draw_elbow_arrow(slide, points: list[tuple[float, float]], color=COLOR_SCHEDULED, pt=1.2, dashed=False, dash_style="dash"):
    """Multi-segment orthogonal arrow ending with an arrowhead."""
    if len(points) < 2:
        return
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        is_last = (i == len(points) - 2)
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        cn.line.color.rgb = C(color)
        cn.line.width = Pt(pt)
        ln = cn.line._get_or_add_ln()
        if dashed:
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash_style}))
        if is_last:
            ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
        _no_shadow(cn)

def draw_badge(slide, num: str, cx: float, cy: float, fill="FFFFFF", border=COLOR_SCHEDULED, text_color=COLOR_SCHEDULED):
    """Draw circular numbered step badge."""
    circle(slide, cx, cy, 0.12, fill=fill, line=border, line_pt=1.0)
    text_box(
        slide,
        cx - 0.12,
        cy - 0.10,
        0.24,
        0.20,
        [[[num, 8, True, False, text_color]]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )


# =========================================================================
# SLIDE 1: FUNCTIONAL & WORKFLOW ARCHITECTURE (DATA PLATFORM & WEB APP)
# =========================================================================
def render_slide_1(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Slide Background
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C(BG_PAGE)

    # Header & Title Block
    text_box(
        slide,
        0.50,
        0.20,
        12.333,
        0.42,
        [
            [("Enterprise Data Platform & Web Application Architecture", 14.5, True, False, TEXT_DARK)],
            [("Automated ML predictions, Agentic AI fallback, automated data governance, and interactive web application workflows", 9, False, True, TEXT_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
    )

    # TOP-LEFT: Data Governance & Quality Container
    GOV_X = 0.50
    GOV_Y = 0.74
    GOV_W = 3.10
    GOV_H = 2.92

    rrect(slide, GOV_X, GOV_Y, GOV_W, GOV_H, fill=ZONE_GOV_BG, line=ZONE_GOV_LINE, line_pt=1.0, radius=0.04)
    hdr_cy = GOV_Y + 0.18
    place_icon(slide, "azure.activity-log", GOV_X + 0.22, hdr_cy, 0.22)
    text_box(
        slide,
        GOV_X + 0.38,
        GOV_Y + 0.06,
        2.60,
        0.24,
        [[("Data Profiling & Quality Governance", 9.2, True, False, TEXT_DARK)]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Card 1: Data Profiling check (DP360)
    DP_X = GOV_X + 0.18
    DP_Y = GOV_Y + 0.36
    DP_W = GOV_W - 0.36
    DP_H = 0.96
    rrect(slide, DP_X, DP_Y, DP_W, DP_H, fill="FFFFFF", line=CARD_LINE_LIGHT, line_pt=0.9, radius=0.04)
    dp_cy = DP_Y + DP_H / 2
    place_icon(slide, "azure.log-analytics-workspaces", DP_X + 0.30, dp_cy, 0.36)
    text_box(
        slide,
        DP_X + 0.56,
        DP_Y + 0.10,
        DP_W - 0.62,
        DP_H - 0.20,
        [
            [("Data Profiling check", 9.5, True, False, TEXT_DARK)],
            [("(DP360)", 8.5, True, False, "334155")],
            [("Automated schema profiling & distribution statistics", 7.2, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Vertical arrow between DP360 and DQ360
    draw_arrow(slide, DP_X + DP_W / 2, DP_Y + DP_H, DP_X + DP_W / 2, DP_Y + DP_H + 0.28, color=COLOR_SCHEDULED, pt=1.1)

    # Card 2: Data Quality check (DQ360)
    DQ_X = DP_X
    DQ_Y = DP_Y + DP_H + 0.28
    DQ_W = DP_W
    DQ_H = 0.96
    rrect(slide, DQ_X, DQ_Y, DQ_W, DQ_H, fill="FFFFFF", line=CARD_LINE_LIGHT, line_pt=0.9, radius=0.04)
    dq_cy = DQ_Y + DQ_H / 2
    place_icon(slide, "azure.policy", DQ_X + 0.30, dq_cy, 0.36)
    text_box(
        slide,
        DQ_X + 0.56,
        DQ_Y + 0.10,
        DQ_W - 0.62,
        DP_H - 0.20,
        [
            [("Data Quality check", 9.5, True, False, TEXT_DARK)],
            [("(DQ360)", 8.5, True, False, "334155")],
            [("Integrity rules, null check & anomaly validation", 7.2, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # TOP-RIGHT: Web App Layer Container (Frontend)
    FE_X = 3.80
    FE_Y = 0.74
    FE_W = 9.03
    FE_H = 2.92

    rrect(slide, FE_X, FE_Y, FE_W, FE_H, fill=ZONE_FE_BG, line=ZONE_FE_LINE, line_pt=1.0, radius=0.04, dashed=True)
    fe_hdr_cy = FE_Y + 0.18
    place_icon(slide, "azure.browser", FE_X + 0.20, fe_hdr_cy, 0.22)
    text_box(
        slide,
        FE_X + 0.36,
        FE_Y + 0.06,
        3.00,
        0.24,
        [[("Web app layer (Frontend Layer)", 9.2, True, False, "4338CA")]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Top Row: 4 Screen Cards
    SCREEN_Y = FE_Y + 0.32
    SCREEN_H = 0.90
    SCREEN_W = 1.68
    SCREEN_GAP = 0.15

    screens_data = [
        {
            "id": "home",
            "title": "Home screen:",
            "sub": "KPIs and charts",
            "icon": "azure.dashboard",
            "badge": "KPIs & Charts",
        },
        {
            "id": "mapping",
            "title": "Mapping screen:",
            "sub": "Show the data with user actions (approve, reject, flagged)",
            "icon": "azure.workflow",
            "badge": "User Actions",
        },
        {
            "id": "qa",
            "title": "QA Validation screen:",
            "sub": "Show Feedback data of Mapping for validation",
            "icon": "azure.compliance",
            "badge": "Validation QA",
        },
        {
            "id": "health",
            "title": "Health insights screen:",
            "sub": "Will show the customer health report",
            "icon": "azure.service-health",
            "badge": "Health Report",
        },
    ]

    screen_centers = []

    for i, s in enumerate(screens_data):
        sx = FE_X + 0.16 + i * (SCREEN_W + SCREEN_GAP)
        screen_centers.append(sx + SCREEN_W / 2)
        rrect(slide, sx, SCREEN_Y, SCREEN_W, SCREEN_H, fill="FFFFFF", line=CARD_LINE_LIGHT, line_pt=0.9, radius=0.04)
        rrect(slide, sx, SCREEN_Y, SCREEN_W, 0.18, fill="F8FAFC", line="E2E8F0", line_pt=0.5, radius=0.02)
        for dot_idx, dot_col in enumerate(["EF4444", "F59E0B", "10B981"]):
            circle(slide, sx + 0.08 + dot_idx * 0.065, SCREEN_Y + 0.09, 0.022, fill=dot_col, line=None)
        text_box(
            slide,
            sx + 0.30,
            SCREEN_Y + 0.01,
            SCREEN_W - 0.34,
            0.16,
            [[[s["badge"], 6.5, True, False, "64748B"]]],
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        body_cy = SCREEN_Y + 0.18 + (SCREEN_H - 0.18) / 2
        place_icon(slide, s["icon"], sx + 0.18, body_cy, 0.26)
        text_box(
            slide,
            sx + 0.34,
            SCREEN_Y + 0.20,
            SCREEN_W - 0.38,
            SCREEN_H - 0.22,
            [
                [(s["title"], 7.8, True, False, TEXT_DARK)],
                [(s["sub"], 6.5, False, False, TEXT_MUTED)],
            ],
            align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    # SQL Server Container
    SQL_X = FE_X + 0.16
    SQL_Y = SCREEN_Y + SCREEN_H + 0.28
    SQL_W = 7.18
    SQL_H = 1.16

    rrect(slide, SQL_X, SQL_Y, SQL_W, SQL_H, fill=SQL_BG, line=SQL_LINE, line_pt=1.0, radius=0.04)

    sql_hdr_cy = SQL_Y + 0.17
    place_icon(slide, "azure.sql-server", SQL_X + 0.20, sql_hdr_cy, 0.22)
    text_box(
        slide,
        SQL_X + 0.36,
        SQL_Y + 0.05,
        2.50,
        0.24,
        [[("SQL Server", 9.5, True, False, "92400E")]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    comp_row1 = [
        ("Data tables", "azure.table", SQL_X + 0.16, SQL_Y + 0.34, 2.10),
        ("feedback tables", "azure.table", SQL_X + 2.42, SQL_Y + 0.34, 2.30),
        ("Audit tables", "azure.table", SQL_X + 4.88, SQL_Y + 0.34, 2.14),
    ]
    for c_title, c_icon, cx, cy, cw in comp_row1:
        ch = 0.30
        rrect(slide, cx, cy, cw, ch, fill="FFFFFF", line="FDE68A", line_pt=0.8, radius=0.03)
        item_cy = cy + ch / 2
        place_icon(slide, c_icon, cx + 0.16, item_cy, 0.16)
        text_box(slide, cx + 0.28, cy, cw - 0.32, ch, [[(c_title, 8, True, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    comp_row2 = [
        ("Transformed views", "azure.table", SQL_X + 1.05, SQL_Y + 0.72, 2.45),
        ("stored procedures", "azure.workflow", SQL_X + 3.75, SQL_Y + 0.72, 2.45),
    ]
    for c_title, c_icon, cx, cy, cw in comp_row2:
        ch = 0.30
        rrect(slide, cx, cy, cw, ch, fill="FFFFFF", line="FDE68A", line_pt=0.8, radius=0.03)
        item_cy = cy + ch / 2
        place_icon(slide, c_icon, cx + 0.16, item_cy, 0.16)
        text_box(slide, cx + 0.28, cy, cw - 0.32, ch, [[(c_title, 8, True, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    for sc_x in screen_centers:
        draw_arrow(slide, sc_x - 0.07, SQL_Y, sc_x - 0.07, SCREEN_Y + SCREEN_H, color=COLOR_SCHEDULED, pt=1.0)
        draw_arrow(slide, sc_x + 0.07, SCREEN_Y + SCREEN_H, sc_x + 0.07, SQL_Y, color=COLOR_DEMAND, pt=1.0, dashed=True)

    # AI Validation on
    AI_TOGGLE_X = SQL_X + SQL_W + 0.24
    AI_TOGGLE_Y = SQL_Y + 0.25
    AI_TOGGLE_W = 1.25
    AI_TOGGLE_H = 0.65

    rrect(slide, AI_TOGGLE_X, AI_TOGGLE_Y, AI_TOGGLE_W, AI_TOGGLE_H, fill=AI_VAL_BG, line=AI_VAL_LINE, line_pt=1.0, radius=0.04)
    circle(slide, AI_TOGGLE_X + 0.18, AI_TOGGLE_Y + 0.32, 0.075, fill="EF4444", line=None)
    text_box(
        slide,
        AI_TOGGLE_X + 0.28,
        AI_TOGGLE_Y + 0.10,
        AI_TOGGLE_W - 0.32,
        0.45,
        [
            [("AI Validation on", 8.2, True, False, "B91C1C")],
            [("(User on demand)", 6.8, False, True, COLOR_DEMAND)],
        ],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # BOTTOM CONTAINER: Backend Layer
    BE_X = 0.50
    BE_Y = 4.10
    BE_W = 12.333
    BE_H = 2.58

    rrect(slide, BE_X, BE_Y, BE_W, BE_H, fill=ZONE_BE_BG, line=ZONE_BE_LINE, line_pt=1.0, radius=0.04, dashed=True)
    place_icon(slide, "azure.databricks", BE_X + 0.20, BE_Y + 0.16, 0.20)
    text_box(
        slide,
        BE_X + 0.35,
        BE_Y + 0.05,
        3.00,
        0.22,
        [[("Backend layer", 9.2, True, False, "1D4ED8")]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    Y_PIPE = BE_Y + 0.40
    CARD_H = 0.88

    # 1. Customer data
    CUST_X = BE_X + 0.65
    CUST_W = 1.20
    rrect(slide, CUST_X, Y_PIPE, CUST_W, CARD_H, fill=CARD_BG, line=CARD_LINE_LIGHT, line_pt=1.0, radius=0.04)
    place_icon(slide, "azure.users", CUST_X + CUST_W / 2, Y_PIPE + 0.28, 0.28)
    text_box(slide, CUST_X + 0.05, Y_PIPE + 0.48, CUST_W - 0.10, 0.34, [[("Customer data", 8.5, True, False, TEXT_DARK)]], align=PP_ALIGN.CENTER)

    DDH_X = CUST_X + CUST_W + 0.45
    draw_arrow(slide, CUST_X + CUST_W, Y_PIPE + CARD_H / 2, DDH_X, Y_PIPE + CARD_H / 2, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "1", (CUST_X + CUST_W + DDH_X) / 2, Y_PIPE + CARD_H / 2)

    # 2. DDH
    DDH_W = 1.40
    rrect(slide, DDH_X, Y_PIPE, DDH_W, CARD_H, fill=CARD_BG, line=CARD_LINE_LIGHT, line_pt=1.0, radius=0.04)
    place_icon(slide, "aws.datazone", DDH_X + DDH_W / 2, Y_PIPE + 0.26, 0.28)
    text_box(
        slide,
        DDH_X + 0.05,
        Y_PIPE + 0.44,
        DDH_W - 0.10,
        0.40,
        [
            [("DDH", 9, True, False, TEXT_DARK)],
            [("(any source, platform)", 7, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    ML_X = DDH_X + DDH_W + 0.45
    draw_arrow(slide, DDH_X + DDH_W, Y_PIPE + CARD_H / 2, ML_X, Y_PIPE + CARD_H / 2, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "2", (DDH_X + DDH_W + ML_X) / 2, Y_PIPE + CARD_H / 2)

    # 3. ML model
    ML_W = 1.80
    rrect(slide, ML_X, Y_PIPE, ML_W, CARD_H, fill=CARD_BG, line=CARD_LINE_LIGHT, line_pt=1.0, radius=0.04)
    place_icon(slide, "azure.machine-learning", ML_X + ML_W / 2, Y_PIPE + 0.26, 0.30)
    text_box(
        slide,
        ML_X + 0.05,
        Y_PIPE + 0.44,
        ML_W - 0.10,
        0.40,
        [
            [("ML model to predict", 8.8, True, False, TEXT_DARK)],
            [("the suggestions from the data", 7.2, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    GE_X = ML_X
    GE_Y = Y_PIPE + CARD_H + 0.35
    GE_W = ML_W
    GE_H = 0.78

    draw_arrow(slide, ML_X + ML_W / 2, Y_PIPE + CARD_H, GE_X + GE_W / 2, GE_Y, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "3", ML_X + ML_W / 2, (Y_PIPE + CARD_H + GE_Y) / 2)

    rrect(slide, GE_X, GE_Y, GE_W, GE_H, fill="FFFFFF", line=CARD_LINE_LIGHT, line_pt=1.0, radius=0.04)
    ge_cy = GE_Y + GE_H / 2
    place_icon(slide, "gcp.maps-platform", GE_X + 0.20, ge_cy, 0.28)
    text_box(
        slide,
        GE_X + 0.38,
        GE_Y + 0.06,
        GE_W - 0.42,
        GE_H - 0.12,
        [
            [("For un-match sources will use", 7.5, True, False, TEXT_DARK)],
            [("Google enrichments, if needed", 7.5, True, False, TEXT_DARK)],
            [("to find lat, long", 6.8, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    TBL_ML_X = ML_X + ML_W + 0.45
    TBL_ML_W = 1.70
    draw_arrow(slide, ML_X + ML_W, Y_PIPE + CARD_H / 2, TBL_ML_X, Y_PIPE + CARD_H / 2, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "4", (ML_X + ML_W + TBL_ML_X) / 2, Y_PIPE + CARD_H / 2)

    # 4. ml_result table
    rrect(slide, TBL_ML_X, Y_PIPE, TBL_ML_W, CARD_H, fill=DB_TABLE_BG, line=DB_TABLE_LINE, line_pt=1.1, radius=0.04)
    place_icon(slide, "azure.databricks", TBL_ML_X + 0.18, Y_PIPE + 0.18, 0.20)
    text_box(slide, TBL_ML_X + 0.32, Y_PIPE + 0.08, TBL_ML_W - 0.36, 0.20, [[("Databricks table", 7.2, True, False, "6D28D9")]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    text_box(
        slide,
        TBL_ML_X + 0.05,
        Y_PIPE + 0.32,
        TBL_ML_W - 0.10,
        0.50,
        [
            [("Store the information in", 7.5, False, False, TEXT_MUTED)],
            [("ml_result table", 9, True, False, TEXT_DARK)],
        ],
        align=PP_ALIGN.CENTER,
    )

    STEP5_X = TBL_ML_X + 0.45
    draw_arrow(slide, STEP5_X, Y_PIPE, STEP5_X, SQL_Y + SQL_H, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "5", STEP5_X, (Y_PIPE + SQL_Y + SQL_H) / 2)

    # Feedback table
    FB_X = TBL_ML_X
    FB_Y = GE_Y
    FB_W = TBL_ML_W
    FB_H = GE_H

    rrect(slide, FB_X, FB_Y, FB_W, FB_H, fill=DB_TABLE_BG, line=DB_TABLE_LINE, line_pt=1.1, radius=0.04)
    place_icon(slide, "azure.databricks", FB_X + 0.18, FB_Y + 0.18, 0.20)
    text_box(slide, FB_X + 0.32, FB_Y + 0.08, FB_W - 0.36, 0.20, [[("Databricks table", 7.2, True, False, "6D28D9")]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    text_box(
        slide,
        FB_X + 0.05,
        FB_Y + 0.30,
        FB_W - 0.10,
        0.44,
        [
            [("Store the information in", 7.5, False, False, TEXT_MUTED)],
            [("feedback table", 9, True, False, TEXT_DARK)],
        ],
        align=PP_ALIGN.CENTER,
    )

    CORRIDOR_X = TBL_ML_X + TBL_ML_W + 0.22
    draw_elbow_arrow(
        slide,
        [
            (SQL_X + 2.42 + 1.15, SQL_Y + SQL_H),
            (CORRIDOR_X, SQL_Y + SQL_H + 0.14),
            (CORRIDOR_X, FB_Y + FB_H / 2),
            (FB_X + FB_W, FB_Y + FB_H / 2),
        ],
        color=COLOR_DEMAND,
        pt=1.2,
        dashed=True,
    )
    draw_badge(slide, "6", CORRIDOR_X, SQL_Y + SQL_H + 0.35, fill="FFFFFF", border=COLOR_DEMAND, text_color=COLOR_DEMAND)
    draw_badge(slide, "9", CORRIDOR_X, FB_Y + 0.02, fill="FFFFFF", border=COLOR_DEMAND, text_color=COLOR_DEMAND)

    # AI + Gmaps
    AI_CARD_X = TBL_ML_X + TBL_ML_W + 0.70
    AI_CARD_W = 2.45

    draw_arrow(slide, TBL_ML_X + TBL_ML_W, Y_PIPE + CARD_H / 2, AI_CARD_X, Y_PIPE + CARD_H / 2, color=COLOR_SCHEDULED, pt=1.2)

    rrect(slide, AI_CARD_X, Y_PIPE, AI_CARD_W, CARD_H, fill=CARD_BG, line="A5B4FC", line_pt=1.0, radius=0.04)
    place_icon(slide, "azure.cognitive-services", AI_CARD_X + 0.28, Y_PIPE + 0.26, 0.28)
    place_icon(slide, "gcp.maps-platform", AI_CARD_X + 0.58, Y_PIPE + 0.26, 0.24)
    text_box(
        slide,
        AI_CARD_X + 0.05,
        Y_PIPE + 0.40,
        AI_CARD_W - 0.10,
        0.44,
        [
            [("AI + Gmaps to predict", 8.5, True, False, "4338CA")],
            [("the suggestion for the records for which ml was not able to predict", 7, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    STEP7_X = AI_CARD_X + AI_CARD_W / 2
    draw_elbow_arrow(
        slide,
        [
            (AI_TOGGLE_X + AI_TOGGLE_W / 2, AI_TOGGLE_Y + AI_TOGGLE_H),
            (AI_TOGGLE_X + AI_TOGGLE_W / 2, Y_PIPE - 0.16),
            (STEP7_X, Y_PIPE - 0.16),
            (STEP7_X, Y_PIPE),
        ],
        color=COLOR_DEMAND,
        pt=1.2,
        dashed=True,
    )
    draw_badge(slide, "7", AI_TOGGLE_X + AI_TOGGLE_W / 2, (AI_TOGGLE_Y + AI_TOGGLE_H + Y_PIPE - 0.16) / 2, fill="FFFFFF", border=COLOR_DEMAND, text_color=COLOR_DEMAND)

    # Agentic table
    TBL_AGT_X = AI_CARD_X
    TBL_AGT_Y = GE_Y
    TBL_AGT_W = AI_CARD_W
    TBL_AGT_H = GE_H

    draw_arrow(slide, AI_CARD_X + AI_CARD_W / 2, Y_PIPE + CARD_H, TBL_AGT_X + TBL_AGT_W / 2, TBL_AGT_Y, color=COLOR_SCHEDULED, pt=1.2)
    draw_badge(slide, "8", AI_CARD_X + AI_CARD_W / 2, (Y_PIPE + CARD_H + TBL_AGT_Y) / 2)

    rrect(slide, TBL_AGT_X, TBL_AGT_Y, TBL_AGT_W, TBL_AGT_H, fill=DB_TABLE_BG, line=DB_TABLE_LINE, line_pt=1.1, radius=0.04)
    place_icon(slide, "azure.databricks", TBL_AGT_X + 0.18, TBL_AGT_Y + 0.18, 0.20)
    text_box(slide, TBL_AGT_X + 0.32, TBL_AGT_Y + 0.08, TBL_AGT_W - 0.36, 0.20, [[("Databricks table", 7.2, True, False, "6D28D9")]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    text_box(
        slide,
        TBL_AGT_X + 0.05,
        TBL_AGT_Y + 0.30,
        TBL_AGT_W - 0.10,
        0.44,
        [
            [("Store the information in", 7.5, False, False, TEXT_MUTED)],
            [("agentic table", 9, True, False, TEXT_DARK)],
        ],
        align=PP_ALIGN.CENTER,
    )

    RETURN_X = TBL_AGT_X + TBL_AGT_W + 0.32
    draw_elbow_arrow(
        slide,
        [
            (TBL_AGT_X + TBL_AGT_W, TBL_AGT_Y + TBL_AGT_H / 2),
            (RETURN_X, TBL_AGT_Y + TBL_AGT_H / 2),
            (RETURN_X, SQL_Y + SQL_H * 0.82),
            (SQL_X + SQL_W, SQL_Y + SQL_H * 0.82),
        ],
        color=COLOR_SCHEDULED,
        pt=1.2,
    )

    # Prominent Auto-Logs Bus
    LOG_BUS_Y = BE_Y + BE_H - 0.20
    LOG_TRUNK_X = BE_X + 0.32

    draw_line(slide, LOG_TRUNK_X, LOG_BUS_Y, TBL_AGT_X + TBL_AGT_W / 2, LOG_BUS_Y, color=COLOR_LOGS, pt=1.6, dashed=True, dash_style="dash")
    draw_arrow(slide, DDH_X + 0.25, Y_PIPE + CARD_H, DDH_X + 0.25, LOG_BUS_Y, color=COLOR_LOGS, pt=1.5, dashed=True, dash_style="dash")

    draw_elbow_arrow(
        slide,
        [
            (TBL_ML_X + 0.20, Y_PIPE + CARD_H),
            (TBL_ML_X + 0.20, FB_Y - 0.12),
            (ML_X + ML_W + 0.18, FB_Y - 0.12),
            (ML_X + ML_W + 0.18, LOG_BUS_Y),
        ],
        color=COLOR_LOGS,
        pt=1.5,
        dashed=True,
        dash_style="dash",
    )

    draw_arrow(slide, FB_X + FB_W / 2, FB_Y + FB_H, FB_X + FB_W / 2, LOG_BUS_Y, color=COLOR_LOGS, pt=1.5, dashed=True, dash_style="dash")
    draw_arrow(slide, TBL_AGT_X + TBL_AGT_W / 2, TBL_AGT_Y + TBL_AGT_H, TBL_AGT_X + TBL_AGT_W / 2, LOG_BUS_Y, color=COLOR_LOGS, pt=1.5, dashed=True, dash_style="dash")

    draw_elbow_arrow(
        slide,
        [
            (LOG_TRUNK_X, LOG_BUS_Y),
            (LOG_TRUNK_X, DQ_Y + DQ_H + 0.25),
            (DQ_X + 0.40, DQ_Y + DQ_H + 0.25),
            (DQ_X + 0.40, DQ_Y + DQ_H),
        ],
        color=COLOR_LOGS,
        pt=1.8,
        dashed=True,
        dash_style="dash",
    )

    rrect(slide, LOG_TRUNK_X - 0.20, (GOV_Y + GOV_H + BE_Y) / 2 - 0.11, 1.30, 0.22, fill="FFFFFF", line=COLOR_LOGS, line_pt=1.0, radius=0.03)
    place_icon(slide, "azure.activity-log", LOG_TRUNK_X - 0.10, (GOV_Y + GOV_H + BE_Y) / 2, 0.14)
    text_box(
        slide,
        LOG_TRUNK_X + 0.02,
        (GOV_Y + GOV_H + BE_Y) / 2 - 0.11,
        1.05,
        0.22,
        [[("Auto-logs stream", 7.5, True, False, COLOR_LOGS)]],
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    rrect(slide, CUST_X + CUST_W + 0.08, LOG_BUS_Y - 0.10, 1.80, 0.20, fill="FFFFFF", line=COLOR_LOGS, line_pt=0.8, radius=0.02)
    text_box(
        slide,
        CUST_X + CUST_W + 0.08,
        LOG_BUS_Y - 0.10,
        1.80,
        0.20,
        [[("Auto-logs bus (DDH & Delta tables)", 6.8, True, False, COLOR_LOGS)]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # HORIZONTAL LEGEND
    LEG_X = 0.50
    LEG_Y = 6.85
    LEG_W = 12.333
    LEG_H = 0.44

    rrect(slide, LEG_X, LEG_Y, LEG_W, LEG_H, fill="FFFFFF", line=CARD_LINE_LIGHT, line_pt=0.9, radius=0.03)
    leg_cy = LEG_Y + LEG_H / 2
    text_box(slide, LEG_X + 0.15, LEG_Y, 0.90, LEG_H, [[("Legend:", 8.5, True, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x1 = LEG_X + 1.05
    draw_arrow(slide, x1, leg_cy, x1 + 0.35, leg_cy, color=COLOR_SCHEDULED, pt=1.2)
    text_box(slide, x1 + 0.40, LEG_Y, 1.45, LEG_H, [[("Scheduled job runs", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x2 = x1 + 1.95
    draw_arrow(slide, x2, leg_cy, x2 + 0.35, leg_cy, color=COLOR_DEMAND, pt=1.2, dashed=True)
    text_box(slide, x2 + 0.40, LEG_Y, 1.30, LEG_H, [[("User on demand", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x3 = x2 + 1.75
    draw_arrow(slide, x3, leg_cy, x3 + 0.35, leg_cy, color=COLOR_LOGS, pt=1.5, dashed=True, dash_style="dash")
    text_box(slide, x3 + 0.40, LEG_Y, 0.95, LEG_H, [[("Auto logs", 7.5, True, False, COLOR_LOGS)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x4 = x3 + 1.40
    rrect(slide, x4, leg_cy - 0.10, 0.28, 0.20, fill=DB_TABLE_BG, line=DB_TABLE_LINE, line_pt=0.8, radius=0.02)
    place_icon(slide, "azure.databricks", x4 + 0.14, leg_cy, 0.14)
    text_box(slide, x4 + 0.34, LEG_Y, 1.40, LEG_H, [[("Databricks tables", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x5 = x4 + 1.78
    rrect(slide, x5, leg_cy - 0.09, 0.28, 0.18, fill=ZONE_FE_BG, line=ZONE_FE_LINE, line_pt=0.8, radius=0.02, dashed=True)
    text_box(slide, x5 + 0.34, LEG_Y, 1.20, LEG_H, [[("Frontend layer", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x6 = x5 + 1.58
    rrect(slide, x6, leg_cy - 0.09, 0.28, 0.18, fill=ZONE_BE_BG, line=ZONE_BE_LINE, line_pt=0.8, radius=0.02, dashed=True)
    text_box(slide, x6 + 0.34, LEG_Y, 1.20, LEG_H, [[("Backend layer", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    x7 = x6 + 1.58
    rrect(slide, x7, leg_cy - 0.09, 0.28, 0.18, fill=ZONE_GOV_BG, line=ZONE_GOV_LINE, line_pt=0.8, radius=0.02)
    text_box(slide, x7 + 0.34, LEG_Y, 1.40, LEG_H, [[("Governance / Checks", 7.5, False, False, TEXT_DARK)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================
# SLIDE 2: TECHNICAL ARCHITECTURE (CLEAN, MINIMAL, EXECUTIVE)
# Matching the exact clean style of build_edtih.py
# 5 core components:
# 1. Frontend: React app hosted on Azure Static Web Apps (SSO enabled)
# 2. Backend: Hono (Node.js) API hosted on Azure App Service (bearer token auth)
# 3. Database: SQL Server (accessible only from backend)
# 4. Data platform: Azure Databricks (trigger jobs & run queries)
# 5. Deployment: Azure DevOps CI/CD pipelines
# =========================================================================
def render_slide_2(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Slide Background
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C(BG_PAGE)

    # ----------------------------------------------------
    # HEADER
    # ----------------------------------------------------
    text_box(
        slide,
        0.65,
        0.42,
        12.0,
        0.38,
        [[("Technical Architecture", 18, True, False, TEXT_DARK)]],
        align=PP_ALIGN.LEFT,
    )
    text_box(
        slide,
        0.65,
        0.82,
        12.0,
        0.24,
        [[("Cloud infrastructure, application hosting, database isolation, data platform, and CI/CD pipelines", 9.2, False, True, TEXT_MUTED)]],
        align=PP_ALIGN.LEFT,
    )

    # Main Horizontal Axis
    Y_MAIN = 3.30

    # ----------------------------------------------------
    # 1. CONTAINERS (Soft Fills, Crisp Lines)
    # ----------------------------------------------------

    # A. Azure Static Web Apps Container (Soft Green)
    SWA_BOX = {"x": 2.30, "y": 1.75, "w": 2.70, "h": 3.00}
    rrect(slide, SWA_BOX["x"], SWA_BOX["y"], SWA_BOX["w"], SWA_BOX["h"], SWA_BG, SWA_LINE, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.static-apps", SWA_BOX["x"] + 0.16, SWA_BOX["y"] - 0.10, 0.26)
    text_box(
        slide, SWA_BOX["x"] + 0.34, SWA_BOX["y"] - 0.22, 2.20, 0.25,
        [[("Azure Static Web Apps", 9, True, False, TEXT_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # B. App Service Container (Soft Blue)
    APP_BOX = {"x": 5.90, "y": 1.75, "w": 2.70, "h": 3.00}
    rrect(slide, APP_BOX["x"], APP_BOX["y"], APP_BOX["w"], APP_BOX["h"], APP_BG, APP_LINE, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.app-services", APP_BOX["x"] + 0.16, APP_BOX["y"] - 0.10, 0.26)
    text_box(
        slide, APP_BOX["x"] + 0.34, APP_BOX["y"] - 0.22, 2.00, 0.25,
        [[("Azure App Service", 9, True, False, TEXT_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # C. Storage / Platform Container (Soft Yellow/Amber)
    STORAGE_BOX = {"x": 9.40, "y": 1.55, "w": 3.30, "h": 3.40}
    rrect(slide, STORAGE_BOX["x"], STORAGE_BOX["y"], STORAGE_BOX["w"], STORAGE_BOX["h"], STORAGE_BG, STORAGE_LINE, line_pt=1.0, radius=0.06)
    place_icon(slide, "azure.sql-server", STORAGE_BOX["x"] + 0.16, STORAGE_BOX["y"] - 0.10, 0.26)
    text_box(
        slide, STORAGE_BOX["x"] + 0.34, STORAGE_BOX["y"] - 0.22, 2.30, 0.25,
        [[("Database & Data Platform", 9.5, True, False, TEXT_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # ----------------------------------------------------
    # 2. CARDS
    # ----------------------------------------------------

    # 1. User
    USER_CX = 0.95
    USER_CY = Y_MAIN
    place_icon(slide, "generic.user", USER_CX, USER_CY - 0.14, 0.48)
    text_box(
        slide, USER_CX - 0.65, USER_CY + 0.24, 1.30, 0.45,
        [
            [("End User", 9.5, True, False, TEXT_DARK)],
            [("Web Browser", 7.8, False, True, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # 2. Frontend Card: React (inside Static Web Apps container)
    FE_CARD = {"x": 2.50, "y": 2.10, "w": 2.30, "h": 2.35}
    rrect(slide, FE_CARD["x"], FE_CARD["y"], FE_CARD["w"], FE_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.05)
    place_icon(slide, "tech.react", FE_CARD["x"] + FE_CARD["w"] / 2, FE_CARD["y"] + 0.52, 0.52)
    text_box(
        slide, FE_CARD["x"] + 0.08, FE_CARD["y"] + 0.90, FE_CARD["w"] - 0.16, 0.68,
        [
            [("Frontend", 8.5, True, False, TEXT_MUTED)],
            [("React Application", 10.5, True, False, "0369A1")],
            [("Single Page App (SPA)", 7.8, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )
    # SSO Pill Badge inside React card
    sso_w = 1.95
    sso_h = 0.34
    sso_x = FE_CARD["x"] + (FE_CARD["w"] - sso_w) / 2
    sso_y = FE_CARD["y"] + FE_CARD["h"] - sso_h - 0.14
    rrect(slide, sso_x, sso_y, sso_w, sso_h, fill="EFF6FF", line="BFDBFE", line_pt=0.8, radius=0.04)
    place_icon(slide, "azure.token-service", sso_x + 0.18, sso_y + sso_h / 2, 0.20)
    text_box(
        slide, sso_x + 0.32, sso_y, sso_w - 0.36, sso_h,
        [[("SSO Enabled", 8, True, False, "1D4ED8")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 3. Backend Card: Hono (inside App Service container)
    BE_CARD = {"x": 6.10, "y": 2.10, "w": 2.30, "h": 2.35}
    rrect(slide, BE_CARD["x"], BE_CARD["y"], BE_CARD["w"], BE_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.05)
    place_icon(slide, "tech.hono", BE_CARD["x"] + BE_CARD["w"] / 2, BE_CARD["y"] + 0.52, 0.50)
    text_box(
        slide, BE_CARD["x"] + 0.08, BE_CARD["y"] + 0.90, BE_CARD["w"] - 0.16, 0.68,
        [
            [("Backend", 8.5, True, False, TEXT_MUTED)],
            [("Hono (Node.js) API", 10.5, True, False, "EA580C")],
            [("REST API Runtime", 7.8, False, False, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )
    # Bearer Token Auth Pill Badge inside Hono card
    auth_w = 2.05
    auth_h = 0.34
    auth_x = BE_CARD["x"] + (BE_CARD["w"] - auth_w) / 2
    auth_y = BE_CARD["y"] + BE_CARD["h"] - auth_h - 0.14
    rrect(slide, auth_x, auth_y, auth_w, auth_h, fill="FAF5FF", line="E9D5FF", line_pt=0.8, radius=0.04)
    place_icon(slide, "azure.security", auth_x + 0.18, auth_y + auth_h / 2, 0.20)
    text_box(
        slide, auth_x + 0.32, auth_y, auth_w - 0.36, auth_h,
        [[("Bearer Token Auth", 8, True, False, "7E22CE")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 4. Database Card: SQL Server (inside Storage container, Top)
    SQL_CARD = {"x": 9.60, "y": 1.85, "w": 2.90, "h": 1.35}
    rrect(slide, SQL_CARD["x"], SQL_CARD["y"], SQL_CARD["w"], SQL_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.05)
    place_icon(slide, "azure.sql-server", SQL_CARD["x"] + 0.42, SQL_CARD["y"] + 0.45, 0.46)
    text_box(
        slide, SQL_CARD["x"] + 0.82, SQL_CARD["y"] + 0.16, SQL_CARD["w"] - 0.90, 0.58,
        [
            [("Database", 8, True, False, TEXT_MUTED)],
            [("SQL Server", 10.5, True, False, "92400E")],
        ],
        align=PP_ALIGN.LEFT,
    )
    # Accessible only from backend security pill
    sec_w = 2.62
    sec_h = 0.32
    sec_x = SQL_CARD["x"] + (SQL_CARD["w"] - sec_w) / 2
    sec_y = SQL_CARD["y"] + SQL_CARD["h"] - sec_h - 0.12
    rrect(slide, sec_x, sec_y, sec_w, sec_h, fill="FEF2F2", line="FECACA", line_pt=0.8, radius=0.03)
    text_box(
        slide, sec_x + 0.08, sec_y, sec_w - 0.16, sec_h,
        [[("Accessible only from backend", 7.5, True, False, "DC2626")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 5. Data Platform Card: Azure Databricks (inside Storage container, Bottom)
    DBR_CARD = {"x": 9.60, "y": 3.40, "w": 2.90, "h": 1.35}
    rrect(slide, DBR_CARD["x"], DBR_CARD["y"], DBR_CARD["w"], DBR_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.05)
    place_icon(slide, "azure.databricks", DBR_CARD["x"] + 0.42, DBR_CARD["y"] + 0.45, 0.46)
    text_box(
        slide, DBR_CARD["x"] + 0.82, DBR_CARD["y"] + 0.16, DBR_CARD["w"] - 0.90, 0.58,
        [
            [("Data Platform", 8, True, False, TEXT_MUTED)],
            [("Azure Databricks", 10.5, True, False, "6D28D9")],
        ],
        align=PP_ALIGN.LEFT,
    )
    # Trigger jobs and run queries pill
    dbr_w = 2.62
    dbr_h = 0.32
    dbr_x = DBR_CARD["x"] + (DBR_CARD["w"] - dbr_w) / 2
    dbr_y = DBR_CARD["y"] + DBR_CARD["h"] - dbr_h - 0.12
    rrect(slide, dbr_x, dbr_y, dbr_w, dbr_h, fill="F5F3FF", line="DDD6FE", line_pt=0.8, radius=0.03)
    text_box(
        slide, dbr_x + 0.08, dbr_y, dbr_w - 0.16, dbr_h,
        [[("Trigger jobs & run queries", 7.5, True, False, "6D28D9")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 6. Deployment Card: Azure DevOps (Bottom-Center)
    DEVOPS_CARD = {"x": 4.60, "y": 5.40, "w": 4.00, "h": 1.35}
    rrect(slide, DEVOPS_CARD["x"], DEVOPS_CARD["y"], DEVOPS_CARD["w"], DEVOPS_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.05)
    place_icon(slide, "azure.devops", DEVOPS_CARD["x"] + DEVOPS_CARD["w"] / 2, DEVOPS_CARD["y"] + 0.42, 0.44)
    text_box(
        slide, DEVOPS_CARD["x"] + 0.08, DEVOPS_CARD["y"] + 0.72, DEVOPS_CARD["w"] - 0.16, 0.55,
        [
            [("Azure DevOps", 10.5, True, False, "0284C7")],
            [("CI/CD Deployment Pipelines", 8.2, False, True, TEXT_MUTED)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # ----------------------------------------------------
    # 3. CONNECTORS (Crisp, Clean, Non-overlapping)
    # ----------------------------------------------------

    # A. User -> Frontend (Straight Horizontal)
    draw_arrow(slide, USER_CX + 0.32, Y_MAIN, FE_CARD["x"], Y_MAIN, color=ARROW_COLOR, pt=1.2)

    # B. Frontend (React) <---> Backend (Hono)
    draw_arrow(
        slide,
        FE_CARD["x"] + FE_CARD["w"],
        Y_MAIN,
        BE_CARD["x"],
        Y_MAIN,
        color=ARROW_COLOR,
        pt=1.3,
        double_ended=True,
    )
    # Clean pill badge on the connector
    rrect(slide, (FE_CARD["x"] + FE_CARD["w"] + BE_CARD["x"]) / 2 - 0.60, Y_MAIN - 0.16, 1.20, 0.32, fill="FFFFFF", line=CARD_LINE, line_pt=0.8, radius=0.03)
    text_box(
        slide,
        (FE_CARD["x"] + FE_CARD["w"] + BE_CARD["x"]) / 2 - 0.60,
        Y_MAIN - 0.16,
        1.20,
        0.32,
        [[("REST API\n(Bearer Token)", 7, True, False, "334155")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # C. Backend (Hono) -> SQL Server & Databricks (Forked connector)
    fork_x = 8.95
    # Stem from Backend right edge
    draw_line(slide, BE_CARD["x"] + BE_CARD["w"], Y_MAIN, fork_x, Y_MAIN, color=ARROW_COLOR, pt=1.2)

    # Branch 1 (Up to SQL Server)
    sql_cy = SQL_CARD["y"] + SQL_CARD["h"] / 2
    draw_elbow_arrow(
        slide,
        [
            (fork_x, Y_MAIN),
            (fork_x, sql_cy),
            (SQL_CARD["x"], sql_cy),
        ],
        color=ARROW_COLOR,
        pt=1.2,
    )

    # Branch 2 (Down to Azure Databricks)
    dbr_cy = DBR_CARD["y"] + DBR_CARD["h"] / 2
    draw_elbow_arrow(
        slide,
        [
            (fork_x, Y_MAIN),
            (fork_x, dbr_cy),
            (DBR_CARD["x"], dbr_cy),
        ],
        color=ARROW_COLOR,
        pt=1.2,
    )

    # D. Azure DevOps -> Frontend & Backend (Dashed CI/CD Pipelines)
    devops_top_y = DEVOPS_CARD["y"]
    devops_guide_y = 5.08
    fe_cx = FE_CARD["x"] + FE_CARD["w"] / 2   # 3.65
    be_cx = BE_CARD["x"] + BE_CARD["w"] / 2   # 7.25

    # Branch 1: CI/CD to Azure Static Web Apps (Elbow up & left)
    draw_elbow_arrow(
        slide,
        [
            (DEVOPS_CARD["x"] + 0.40, devops_top_y),
            (DEVOPS_CARD["x"] + 0.40, devops_guide_y),
            (fe_cx, devops_guide_y),
            (fe_cx, SWA_BOX["y"] + SWA_BOX["h"]),
        ],
        color="0284C7",
        pt=1.2,
        dashed=True,
    )
    # Branch 1 Pill
    rrect(slide, fe_cx - 0.55, 4.96, 1.10, 0.24, fill="FFFFFF", line="93C5FD", line_pt=0.8, radius=0.03)
    text_box(
        slide, fe_cx - 0.55, 4.96, 1.10, 0.24,
        [[("CI/CD Deploy", 7.2, True, False, "0369A1")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Branch 2: CI/CD to Azure App Service (Clean straight vertical line)
    draw_arrow(
        slide,
        be_cx,
        devops_top_y,
        be_cx,
        APP_BOX["y"] + APP_BOX["h"],
        color="0284C7",
        pt=1.2,
        dashed=True,
    )
    # Branch 2 Pill
    rrect(slide, be_cx - 0.55, 4.96, 1.10, 0.24, fill="FFFFFF", line="93C5FD", line_pt=0.8, radius=0.03)
    text_box(
        slide, be_cx - 0.55, 4.96, 1.10, 0.24,
        [[("CI/CD Deploy", 7.2, True, False, "0369A1")]],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )



# =========================================================================
# MAIN PRESENTATION BUILDER
# =========================================================================
def build_deck(out_path: str | Path = "out/recreated_architecture.pptx") -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    print("Rendering Slide 1: Functional Data Platform & Web App Architecture...")
    render_slide_1(prs)

    print("Rendering Slide 2: Technical Cloud & Infrastructure Architecture...")
    render_slide_2(prs)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    import os; os.makedirs("out", exist_ok=True)
    from archgen.shared.geomcheck import fix_presentation
    fix_presentation(prs)  # hug text boxes, snap near-miss alignment
    prs.save(str(out_path))
    print(f"Full presentation deck saved to: {out_path}")

    # Standalone slide decks for preview / rendering
    prs_s1 = Presentation()
    prs_s1.slide_width = Inches(13.333)
    prs_s1.slide_height = Inches(7.5)
    render_slide_1(prs_s1)
    prs_s1.save("out/slide1_functional.pptx")

    prs_s2 = Presentation()
    prs_s2.slide_width = Inches(13.333)
    prs_s2.slide_height = Inches(7.5)
    render_slide_2(prs_s2)
    prs_s2.save("out/slide2_technical.pptx")
    print("Standalone slide decks saved to out/slide1_functional.pptx and out/slide2_technical.pptx")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "out/recreated_architecture.pptx"
    build_deck(target)
