"""Proposed GCP architecture deck - All 4 slides in ONE presentation.

Slides:
1. DT + DSO · Common architecture
2. Physical Covers / Proposed architecture
3. Virtual Tank / Proposed architecture
4. Impact Analysis / Proposed GCP architecture

Design & Execution Rules:
- All 4 slides contained in ONE .pptx file (no separate PPT files)
- Completely flat connectors and shapes with NO SHADOWS (strips <p:style> & effectLst)
- Clean, exact-snapped orthogonal connector routing (snapping to shape edges)
- Bottom footnote sentences omitted across all slides per user request
- Consistent soft pastel theme, Segoe UI typography, and grouped shapes
- Existing project files remain completely untouched
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from archgen.registry import emf
from archgen.registry.resolver import IconResolver

ROOT = Path(__file__).resolve().parents[2]  # repo root (bespoke/<name>/build.py)
ICONS = ROOT / "assets" / "icons"
R = IconResolver()
FONT = "Segoe UI"

# ---- Design System / Theme Constants ----------------------------------------
PAGE = "FBFCFE"
INK = "1E293B"
MUTED = "64748B"
ARROW = "5F86C9"
ARROW2 = "94A3B8"
ARROW_DASH = "94A3B8"
CARD_LINE = "DCE3EE"
CARD_BG = "FFFFFF"
SUBHEAD = "2F4F8F"
SUCCESS_GREEN = "16A34A"

THEMES = {
    "users": ("EEF3FB", "C9D6EC", "2F4F8F"),
    "azure": ("EAF4F8", "BCD9E6", "1F6F8B"),
    "gcp": ("F7F9FC", "D3DBE8", "2F4F8F"),
    "run": ("F3F4FB", "B9BFE3", "4B4F9E"),
    "mods": ("EEF3FB", "C9D6EC", "2F4F8F"),
    "dbx": ("F5F2FC", "D5CBEF", "5B45A3"),
    "grey": ("F1F5F9", "DCE3EE", "475569"),
}

BRAND = {
    "databricks.svg": "#FF3621",
    "Azure-Devops.svg": "#0078D7",
}

TECH = {
    "react": ICONS / "tech" / "React.svg",
    "fastapi": ICONS / "tech" / "FastAPI.svg",
    "gcloud": ICONS / "tech" / "Google-Cloud.svg",
    "run": ICONS / "gcp" / "Cloud-Run.svg",
    "sql": ICONS / "gcp" / "Cloud-SQL.svg",
    "storage": ICONS / "gcp" / "Cloud-Storage.svg",
    "secret": ICONS / "gcp" / "Secret-Manager.svg",
    "azure": ICONS / "tech" / "Azure.svg",
    "azure.devops": ICONS / "tech" / "Azure-Devops.svg",
    "dbx": ICONS / "tech" / "databricks.svg",
    "python": ICONS / "tech" / "Python.svg",
    "entra": ICONS / "azure" / "identity" / "02251-icon-service-Entra-Privleged-Identity-Management.svg",
}


def C(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def flat(shp):
    """Strip style references and shadow effects completely for 100% flat rendering."""
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    eff = shp._element.spPr.find(qn("a:effectLst"))
    if eff is not None:
        shp._element.spPr.remove(eff)
    shp._element.spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))


def key_path(key: str) -> Path:
    if key in TECH:
        return TECH[key]
    entry = R.get(key)
    return entry.path if entry else (ICONS / "generic" / f"{key.split('.')[-1]}.svg")


def ico(k: str) -> Path:
    return key_path(k)


# ---- Drawing Primitives -----------------------------------------------------
def rrect(sh, x, y, w, h, fill, line, pt=0.9, radius=0.05, dashed=False):
    s = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.adjustments[0] = min(radius, 0.5 * 0.3 / max(min(w, h), 0.01) * 2) if radius < 0.2 else radius
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = C(fill)
    else:
        s.fill.background()
    s.line.color.rgb = C(line)
    s.line.width = Pt(pt)
    if dashed:
        ln = s.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    flat(s)
    return s


def text(sh, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """paras: list of list of (text, size, bold, color, mono)."""
    b = sh.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = b.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(18000)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for item in para:
            t, size, bold, color, mono = item
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.name = FONT
            r.font.color.rgb = C(color)
    return b


def pic(sh, path: Path, cx, cy, size):
    png = emf.to_picture(path, recolor_hex=BRAND.get(path.name), prefer_vector=False)
    asp = emf.svg_aspect(path)
    w, h = (size, size / asp) if asp >= 1 else (size * asp, size)
    sh.add_picture(str(png), Inches(cx - w / 2), Inches(cy - h / 2), width=Inches(w), height=Inches(h))


# ---- Component Builders ----------------------------------------------------
def container(slide, kind, x, y, w, h, title, icon=None, dashed=False) -> dict:
    fill, line, lab_col = THEMES[kind]
    rrect(slide.shapes, x, y, w, h, fill, line, pt=1.0, radius=0.035, dashed=dashed)
    tx = x + 0.16
    if icon:
        pic(slide.shapes, ico(icon), x + 0.28, y + 0.20, 0.24)
        tx = x + 0.46
    text(slide.shapes, tx, y + 0.07, w - (tx - x) - 0.1, 0.26, [[(title, 10.5, True, lab_col, False)]], anchor=MSO_ANCHOR.MIDDLE)
    return dict(l=x, r=x + w, t=y, b=y + h, cx=x + w / 2, cy=y + h / 2, w=w, h=h)


def card(slide, x, y, w, h, title, lines, icon=None, mono=False, fill=CARD_BG, title_pt=9.5, radius=0.07) -> dict:
    """Card shape with icon, title, and body lines. Returns boundary anchors."""
    g = slide.shapes.add_group_shape()
    rrect(g.shapes, x, y, w, h, fill, CARD_LINE, radius=radius)
    tx = x + 0.12
    if icon:
        pic(g.shapes, ico(icon), x + 0.26, y + 0.21, 0.23)
        tx = x + 0.44
    if title:
        text(g.shapes, tx, y + 0.08, w - (tx - x) - 0.08, 0.26, [[(title, title_pt, True, INK, False)]],
             anchor=MSO_ANCHOR.MIDDLE)
    paras = []
    for line in lines:
        if isinstance(line, list):
            paras.append(line)
        else:
            paras.append([(line, 7.5, False, MUTED, mono)])
    if paras:
        text(g.shapes, x + 0.12, y + 0.36, w - 0.22, max(0.2, h - 0.40), paras)
    return dict(l=x, r=x + w, t=y, b=y + h, cx=x + w / 2, cy=y + h / 2, w=w, h=h)


def node(slide, cx, y, icon, title, sub, tile=0.62, w=2.6) -> dict:
    """Grouped icon tile with title and subtitle underneath."""
    g = slide.shapes.add_group_shape()
    rrect(g.shapes, cx - tile / 2, y, tile, tile, CARD_BG, CARD_LINE, radius=0.18)
    pic(g.shapes, ico(icon), cx, y + tile / 2, tile * 0.56)
    paras = [[(title, 10, True, INK, False)]]
    for s in sub:
        paras.append([(s, 7.5, False, MUTED, True)])
    th = 0.2 + 0.15 * len(sub) + 0.05
    text(g.shapes, cx - w / 2, y + tile + 0.06, w, th, paras, align=PP_ALIGN.CENTER)
    return dict(l=cx - tile / 2, r=cx + tile / 2, t=y, b=y + tile, cx=cx, cy=y + tile / 2,
                bot=y + tile + th)


# ---- Clean Flat Connectors (NO SHADOWS) -------------------------------------
def seg(slide, x1, y1, x2, y2, color=ARROW, pt=1.2, end=True, start=False, dash=False):
    """Draw a clean flat connector line with NO shadows."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dash:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if start:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if end:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    flat(cn)  # Strips <p:style> and empty effectLst removes drop shadow
    return cn


def path(slide, pts: List[Tuple[float, float]], color=ARROW, pt=1.2, end=True, start=False, dash=False):
    """Draw an orthogonal or multi-segment path cleanly connecting vertices with NO shadows."""
    n = len(pts) - 1
    for i in range(n):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        seg(slide, x1, y1, x2, y2, color=color, pt=pt,
            end=end and i == n - 1, start=start and i == 0, dash=dash)


def lab(slide, x, y, s, w=1.6, color=ARROW, align=PP_ALIGN.CENTER):
    """Label along or above a connector path."""
    text(slide.shapes, x, y, w, 0.18, [[(s, 7.5, False, color, True)]], align=align)


def header(slide, title: str, subtitle: str):
    """Accent pill + title and subtitle."""
    rrect(slide.shapes, 0.35, 0.34, 0.09, 0.58, "9DB7F0", "9DB7F0", pt=0.1, radius=0.5)
    text(slide.shapes, 0.55, 0.26, 11.5, 0.42, [[(title, 21, True, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    text(slide.shapes, 0.55, 0.68, 12.0, 0.26, [[(subtitle, 10.0, False, MUTED, False)]], anchor=MSO_ANCHOR.MIDDLE)


def new_slide(prs: Presentation, title: str, subtitle: str):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = C(PAGE)
    header(s, title, subtitle)
    return s


# ============================================================================
# SLIDE 1: DT + DSO · Common architecture (Image 1 recreation)
# ============================================================================
def slide_dt_dso_common(prs: Presentation):
    s = new_slide(
        prs,
        "DT + DSO · Common architecture",
        "One delivered module · shared identity, application, data and delivery services · Phase 0 proposal",
    )

    # 1. Left column: Users
    u_box = container(s, "users", 0.35, 1.05, 2.55, 2.15, "Users · one workspace", "generic.user")
    b1 = card(s, 0.50, 1.48, 2.25, 0.76, "Buyer", ["Upload · review · run scenarios"])
    b2 = card(s, 0.50, 2.32, 2.25, 0.76, "Administrator / Support", ["Access · mappings · investigation"])

    # 2. Left column: Azure · identity
    az_box = container(s, "azure", 0.35, 3.32, 2.55, 1.82, "Azure · identity", "azure")
    ent = card(s, 0.50, 3.76, 2.25, 1.26, "Microsoft Entra ID", [
        "SSO · app registrations /",
        "SPNs",
        "AD groups · roles"
    ], icon="entra", title_pt=9.0)

    # 3. Left column: Azure DevOps
    do_box = container(s, "azure", 0.35, 5.26, 2.55, 1.84, "Azure DevOps", "azure.devops")
    cicd = card(s, 0.50, 5.70, 2.25, 1.28, "Common CI/CD", [
        [("Azure Repos\n", 7.5, True, INK, False), ("React · FastAPI · DT / DSO code", 7.0, False, MUTED, False)],
        [("Azure Pipelines\n", 7.5, True, INK, False), ("Build · checks · package", 7.0, False, MUTED, False)],
        [("Release stages\n", 7.5, True, INK, False), ("Deploy app + Databricks jobs\nEnvironment gates / rollback¹", 7.0, False, MUTED, False)],
    ], title_pt=9.0)

    # 4. Right column: UPALM · shared application & processing
    gcp_box = container(s, "gcp", 3.15, 1.05, 9.83, 6.05, "UPALM · shared application & processing", "gcloud")

    # App Row: React & FastAPI
    react = card(
        s, 3.38, 1.62, 3.25, 1.35, "React",
        [
            "Cloud Run · shared frontend",
            "Uploads · mapping inputs · review",
            "Scenario settings · comparison · export",
            [("SSO / PKCE", 7.0, False, ARROW2, True)],
        ],
        icon="react", title_pt=11.0
    )
    api = card(
        s, 9.42, 1.62, 3.25, 1.35, "FastAPI",
        [
            "Cloud Run · shared API",
            "Authorization · validation · job control",
            "Progress · readiness · results",
        ],
        icon="fastapi", title_pt=11.0
    )

    # Middle Row: Databricks · DT + DSO execution
    dbx_box = container(s, "dbx", 3.38, 3.15, 9.35, 2.75, "Databricks · DT + DSO execution", "dbx")

    # Sub-row 1: DT · Agentic solution
    text(s.shapes, 3.58, 3.42, 4.0, 0.22, [[("DT · Agentic solution", 9.5, True, SUBHEAD, False)]])
    dt_xs = [3.58, 5.88, 8.18, 10.48]
    dt_w, dt_h = 2.05, 0.88
    dt_cards = [
        card(s, dt_xs[0], 3.75, dt_w, dt_h, "Evaluate documents", [
            "Excel · CSV · PDF · Images¹", "Readability / structure checks"
        ], icon="generic.document", title_pt=8.5),
        card(s, dt_xs[1], 3.75, dt_w, dt_h, "Extract fields", [
            "Multimodal agent · native parsing", "Plan rows · values · source refs"
        ], icon="generic.search", title_pt=8.5),
        card(s, dt_xs[2], 3.75, dt_w, dt_h, "Calculate + map", [
            "Python · rules / FX / maps", "EXW → CIF → DDP"
        ], icon="generic.analytics", title_pt=8.5),
        card(s, dt_xs[3], 3.75, dt_w, dt_h, "Validate Costbook", [
            "Routes / fields / prices / QoQ",
            [("Costbook ready ", 7.5, True, SUCCESS_GREEN, False), ("or Needs review", 7.5, False, MUTED, False)]
        ], icon="generic.security", title_pt=8.5),
    ]

    # Sub-row 2: DSO · Prepare data → create scenario
    text(s.shapes, 3.58, 4.80, 5.0, 0.22, [[("DSO · Prepare data → create scenario", 9.5, True, SUBHEAD, False)]])
    dso_cards = [
        card(s, dt_xs[0], 5.12, dt_w, dt_h, "Prepare inputs", [
            "Costbook + demand / references", "Enrich prices and route data"
        ], icon="generic.document", title_pt=8.5),
        card(s, dt_xs[1], 5.12, dt_w, dt_h, "Validate + DP ID", [
            "Persist validated input snapshot",
            [("Ready to create scenario", 7.5, True, SUCCESS_GREEN, False)]
        ], icon="generic.security", title_pt=8.5),
        card(s, dt_xs[2], 5.12, dt_w, dt_h, "Solve scenario", [
            "Buyer rules · guided presets", "Pyomo / IPOPT baseline²"
        ], icon="generic.analytics", title_pt=8.5),
        card(s, dt_xs[3], 5.12, dt_w, dt_h, "Save results", [
            "OPT ID · solver status · costs", "Allocations · comparison · export"
        ], icon="generic.storage", title_pt=8.5),
    ]

    # Bottom Row: Data Services (Storage & SQL)
    st = card(s, 3.38, 6.05, 4.45, 0.95, "Cloud Storage", [
        "Uploads · masters · Costbooks · DP snapshots",
        "Scenario datasets · exports · source evidence"
    ], icon="storage", title_pt=9.5)
    sq = card(s, 8.28, 6.05, 4.45, 0.95, "Cloud SQL · PostgreSQL", [
        "Job / run IDs · readiness · DP / OPT metadata",
        "Scenario settings · access · file references"
    ], icon="sql", title_pt=9.5)

    # ---- Clean Flat Connectors (NO SHADOWS) ---------------------------------
    # 1. Buyer to React
    seg(s, b1["r"], b1["cy"], react["l"], b1["cy"], color=ARROW, pt=1.2)

    # 2. Entra ID to React (SSO / PKCE)
    path(s, [
        (ent["r"], ent["cy"]),
        (3.02, ent["cy"]),
        (3.02, react["b"] - 0.20),
        (react["l"], react["b"] - 0.20)
    ], color=ARROW2, pt=1.1)

    # 3. React to FastAPI (API + token)
    seg(s, react["r"], react["t"] + 0.38, api["l"], api["t"] + 0.38, color=ARROW, pt=1.2)
    lab(s, (react["r"] + api["l"]) / 2 - 1.0, react["t"] + 0.18, "API + token", 2.0)

    # 4. FastAPI to React (poll status)
    seg(s, api["l"], react["b"] - 0.38, react["r"], react["b"] - 0.38, color=ARROW, pt=1.1, dash=True)
    lab(s, (react["r"] + api["l"]) / 2 - 1.5, react["b"] - 0.32, "poll status / readiness / results", 3.0)

    # 5. FastAPI to Databricks (launch job / check status)
    seg(s, api["cx"] - 0.55, api["b"], api["cx"] - 0.55, dbx_box["t"], color=ARROW, pt=1.2)
    lab(s, api["cx"] - 1.70, 2.72, "launch job / check status", 2.3, align=PP_ALIGN.LEFT)

    # 6. DT Row sequential connectors
    for i in range(3):
        seg(s, dt_cards[i]["r"], dt_cards[i]["cy"], dt_cards[i + 1]["l"], dt_cards[i + 1]["cy"], color=ARROW, pt=1.1)

    # 7. DT targeted retry / review loop
    path(s, [
        (dt_cards[3]["cx"], dt_cards[3]["t"]),
        (dt_cards[3]["cx"], 3.58),
        (dt_cards[1]["cx"], 3.58),
        (dt_cards[1]["cx"], dt_cards[1]["t"])
    ], color=ARROW2, pt=1.0, dash=True)
    lab(s, (dt_cards[1]["cx"] + dt_cards[3]["cx"]) / 2 - 1.2, 3.42, "targeted retry / review", 2.4, color=ARROW2)

    # 8. DT Card 4 -> DSO Card 1 (Validated Costbook -> DSO preparation)
    path(s, [
        (dt_cards[3]["cx"], dt_cards[3]["b"]),
        (dt_cards[3]["cx"], 4.70),
        (dso_cards[0]["l"] - 0.18, 4.70),
        (dso_cards[0]["l"] - 0.18, dso_cards[0]["cy"]),
        (dso_cards[0]["l"], dso_cards[0]["cy"])
    ], color=ARROW, pt=1.2)
    lab(s, (dso_cards[0]["l"] + dt_cards[3]["cx"]) / 2 - 1.7, 4.56, "Validated Costbook → DSO preparation", 3.4)

    # 9. DSO Row sequential connectors
    seg(s, dso_cards[0]["r"], dso_cards[0]["cy"], dso_cards[1]["l"], dso_cards[1]["cy"], color=ARROW, pt=1.1)
    seg(s, dso_cards[1]["r"], dso_cards[1]["cy"], dso_cards[2]["l"], dso_cards[2]["cy"], color=ARROW, pt=1.1)
    lab(s, (dso_cards[1]["r"] + dso_cards[2]["l"]) / 2 - 0.7, dso_cards[1]["t"] - 0.22, "Create scenario", 1.4)
    seg(s, dso_cards[2]["r"], dso_cards[2]["cy"], dso_cards[3]["l"], dso_cards[3]["cy"], color=ARROW, pt=1.1)

    # 10. Databricks to Cloud Storage (read / write files)
    seg(s, st["cx"] - 0.25, dbx_box["b"], st["cx"] - 0.25, st["t"], color=ARROW, pt=1.2)
    lab(s, st["cx"] - 0.15, (dbx_box["b"] + st["t"]) / 2 - 0.10, "read / write files", 1.6, align=PP_ALIGN.LEFT)

    # 11. FastAPI to Cloud SQL · PostgreSQL
    path(s, [
        (api["r"], api["cy"] + 0.15),
        (12.82, api["cy"] + 0.15),
        (12.82, sq["cy"]),
        (sq["r"], sq["cy"])
    ], color=ARROW, pt=1.2)

    # 12. Azure DevOps Continuous Deployment Loop (Dashed)
    # Top bus: deploy React + API
    path(s, [
        (cicd["r"], cicd["t"] + 0.20),
        (2.88, cicd["t"] + 0.20),
        (2.88, 1.42),
        (12.98, 1.42),
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    # Taps into React and FastAPI tops
    seg(s, react["cx"], 1.42, react["cx"], react["t"], color=ARROW_DASH, pt=1.0, dash=True)
    seg(s, api["cx"], 1.42, api["cx"], api["t"], color=ARROW_DASH, pt=1.0, dash=True)
    lab(s, 5.9, 1.28, "deploy React + API", 2.2, color=ARROW2)

    # Bottom bus: deploy DT / DSO jobs + configuration
    path(s, [
        (cicd["r"], cicd["b"] - 0.20),
        (2.88, cicd["b"] - 0.20),
        (2.88, 7.15),
        (12.98, 7.15),
        (12.98, 1.42)  # joins with top perimeter
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    lab(s, 6.8, 7.00, "deploy DT / DSO jobs + configuration", 3.5, color=ARROW2)


# ============================================================================
# SLIDE 2: Physical Covers / Proposed architecture (Image 2 recreation)
# ============================================================================
def slide_physical_covers(prs: Presentation):
    s = new_slide(
        prs,
        "Physical Covers / Proposed architecture",
        "React + FastAPI on Cloud Run · Azure identity + Azure DevOps delivery · Phase 0 proposal",
    )

    # Left Column: User roles · proposed
    u_box = container(s, "users", 0.35, 1.05, 2.85, 2.90, "User roles · proposed", "generic.user")
    rw, rh = 1.22, 0.72
    b1 = card(s, 0.48, 1.48, rw, rh, "Buyer", ["Own purchases", "& coverage"], title_pt=8.5)
    b2 = card(s, 1.82, 1.48, rw + 0.1, rh, "Procurement manager", ["Team purchases", "& coverage"], title_pt=8.0)
    b3 = card(s, 0.48, 2.28, rw, rh, "Director", ["Aggregated", "reporting"], title_pt=8.5)
    b4 = card(s, 1.82, 2.28, rw + 0.1, rh, "Administrator", ["Demand / reference", "uploads"], title_pt=8.0)
    text(s.shapes, 0.48, 3.65, 2.6, 0.26, [[("Final role/action matrix to be confirmed.", 7.5, False, MUTED, False)]])

    # Left Column: Azure · enterprise identity
    az_box = container(s, "azure", 0.35, 4.05, 2.85, 1.55, "Azure · enterprise identity", "azure")
    ent = node(s, 1.77, 4.38, "entra", "Microsoft Entra ID", [
        "SSO · app registrations / SPNs · AD groups",
        "API roles and business data scope enforced server-side"
    ], tile=0.48, w=2.7)

    # Left Column: Azure DevOps
    do_box = container(s, "azure", 0.35, 5.70, 2.85, 1.52, "Azure DevOps", "azure.devops")
    cicd = card(s, 0.48, 6.08, 2.58, 1.02, "Common CI/CD", [
        [("Azure Repos: ", 7.0, True, INK, False), ("React · FastAPI · SQL / configuration", 7.0, False, MUTED, False)],
        [("Azure Pipelines: ", 7.0, True, INK, False), ("Build · checks · package", 7.0, False, MUTED, False)],
        [("Release stages: ", 7.0, True, INK, False), ("Deploy frontend + API · Environment gates", 7.0, False, MUTED, False)],
    ], title_pt=8.5)

    # Right Container: Google Cloud · application & data
    gcp_box = container(s, "gcp", 3.42, 1.05, 9.56, 6.17, "Google Cloud · application & data", "gcloud")

    # Cloud Run FE & API containers
    fe_box = container(s, "run", 3.68, 1.56, 3.40, 1.62, "Cloud Run · FE", "run", dashed=True)
    rx = node(s, 5.38, 1.94, "react", "React", ["UI assets / browser MSAL"], tile=0.55, w=3.0)

    api_box = container(s, "run", 9.18, 1.56, 3.55, 1.62, "Cloud Run · API", "run", dashed=True)
    fa = node(s, 10.95, 1.94, "fastapi", "FastAPI", ["validation / role checks"], tile=0.55, w=3.0)

    # Lower Services
    st = node(s, 5.10, 4.45, "storage", "Cloud Storage", ["private uploads / exports"], tile=0.52, w=2.8)
    sq = node(s, 8.20, 4.45, "sql", "Cloud SQL", ["PostgreSQL / FIFO / prices"], tile=0.52, w=2.8)
    sm = node(s, 11.20, 4.45, "secret", "Secret Manager", ["server-side secrets"], tile=0.52, w=2.8)

    # Bottom notes inside GCP container
    note_paras = [
        [("Reuse: ", 8.0, True, INK, False), ("existing API / SQL calculations after verification; migrate Streamlit business logic server-side.", 8.0, False, MUTED, False)],
        [("Data: ", 8.0, True, INK, False), ("manual files; preserve Impact Analysis snapshots. GCP services use workload identities.", 8.0, False, MUTED, False)]
    ]
    text(s.shapes, 3.68, 5.80, 9.0, 0.40, note_paras)

    # Connections (Clean Flat, NO SHADOWS)
    # User to React
    seg(s, u_box["r"], rx["cy"], rx["l"], rx["cy"], color=ARROW, pt=1.2)
    lab(s, u_box["r"] + 0.15, rx["cy"] - 0.20, "load UI", 0.9, align=PP_ALIGN.LEFT)

    # Entra ID to React
    path(s, [
        (ent["r"], ent["cy"]),
        (3.30, ent["cy"]),
        (3.30, 3.45),
        (rx["cx"], 3.45),
        (rx["cx"], rx["bot"])
    ], color=ARROW2, pt=1.1)
    lab(s, 3.45, 3.28, "SSO / PKCE", 1.0, color=ARROW2, align=PP_ALIGN.LEFT)

    # Entra ID to FastAPI
    path(s, [
        (ent["r"], ent["cy"] + 0.18),
        (3.20, ent["cy"] + 0.18),
        (3.20, 5.75),
        (12.85, 5.75),
        (12.85, fa["cy"] + 0.20),
        (fa["r"] + 0.20, fa["cy"] + 0.20)
    ], color=ARROW2, pt=1.1)
    lab(s, 9.40, 5.58, "issuer / signing keys", 3.2, color=ARROW2, align=PP_ALIGN.RIGHT)

    # React to FastAPI
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"], color=ARROW, pt=1.2)
    lab(s, 6.70, fa["cy"] - 0.20, "API + token", 3.5)

    # FastAPI to Storage
    path(s, [
        (9.40, 3.17),
        (9.40, 3.85),
        (st["cx"], 3.85),
        (st["cx"], st["t"])
    ], color=ARROW2, pt=1.1)
    lab(s, 6.50, 3.68, "files", 1.0, color=ARROW2)

    # FastAPI to SQL
    path(s, [
        (9.80, 3.17),
        (9.80, 4.05),
        (sq["cx"], 4.05),
        (sq["cx"], sq["t"])
    ], color=ARROW, pt=1.2)
    lab(s, 9.85, 3.45, "SQL", 0.6, align=PP_ALIGN.LEFT)

    # FastAPI to Secret Manager
    seg(s, sm["cx"], 3.17, sm["cx"], sm["t"], color=ARROW2, pt=1.1)
    lab(s, sm["cx"] + 0.08, 3.45, "read secrets", 1.2, color=ARROW2, align=PP_ALIGN.LEFT)

    # Azure DevOps deploy bus (routed cleanly above Cloud Run FE/API boxes)
    path(s, [
        (cicd["r"], cicd["cy"]),
        (3.10, cicd["cy"]),
        (3.10, 1.40),
        (12.85, 1.40),
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    seg(s, fe_box["cx"], 1.40, fe_box["cx"], fe_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    seg(s, api_box["cx"], 1.40, api_box["cx"], api_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    lab(s, 6.0, 1.25, "deploy frontend + API", 2.5, color=ARROW2)


# ============================================================================
# SLIDE 3: Virtual Tank / Proposed architecture (Image 3 recreation)
# ============================================================================
def slide_virtual_tank(prs: Presentation):
    s = new_slide(
        prs,
        "Virtual Tank / Proposed architecture",
        "React + FastAPI on separate Cloud Run services · Existing allocation job · Azure enterprise identity · Azure DevOps CI/CD · Phase 0 proposal",
    )

    # Left Column: User roles · VT personas
    u_box = container(s, "users", 0.35, 1.05, 2.95, 2.90, "User roles · VT personas", "generic.user")
    card(s, 0.48, 1.48, 1.28, 1.00, "Oil buyer", ["Purchase uploads", "Reservation requests"], title_pt=8.5)
    card(s, 1.88, 1.48, 1.30, 1.00, "Sales / demand planner", ["Demand / consumption", "Review / finalize"], title_pt=8.0)
    card(s, 0.48, 2.55, 2.70, 0.80, "Portfolio buyer", ["Read-only balances, coverage and prices"], title_pt=8.5)
    text(s.shapes, 0.48, 3.65, 2.7, 0.26, [[("Final role/action matrix to be confirmed.", 7.5, False, MUTED, False)]])

    # Left Column: Azure · enterprise identity
    az_box = container(s, "azure", 0.35, 4.05, 2.95, 1.55, "Azure · enterprise identity", "azure")
    ent = node(s, 1.82, 4.38, "entra", "Microsoft Entra ID", [
        "SSO · app registrations / SPNs · AD groups"
    ], tile=0.48, w=2.7)

    # Left Column: Azure DevOps
    do_box = container(s, "azure", 0.35, 5.70, 2.95, 1.52, "Azure DevOps", "azure.devops")
    cicd = card(s, 0.48, 6.08, 2.68, 1.02, "Common CI/CD", [
        [("Azure Repos: ", 7.0, True, INK, False), ("React · FastAPI · allocation code", 7.0, False, MUTED, False)],
        [("Azure Pipelines: ", 7.0, True, INK, False), ("Build · checks · package", 7.0, False, MUTED, False)],
        [("Release stages: ", 7.0, True, INK, False), ("Deploy frontend + API · Environment gates¹", 7.0, False, MUTED, False)],
    ], title_pt=8.5)

    # Right Container: Google Cloud · application & data
    gcp_box = container(s, "gcp", 3.50, 1.05, 9.48, 6.17, "Google Cloud · application & data", "gcloud")

    # Cloud Run FE & API containers
    fe_box = container(s, "run", 3.75, 1.56, 3.30, 1.55, "Cloud Run · FE", "run", dashed=True)
    rx = node(s, 5.40, 1.88, "react", "React", ["UI assets / browser MSAL"], tile=0.52, w=2.8)

    api_box = container(s, "run", 9.15, 1.56, 3.55, 1.55, "Cloud Run · API", "run", dashed=True)
    fa = node(s, 10.92, 1.88, "fastapi", "FastAPI", ["validation / roles / ledger APIs"], tile=0.52, w=3.1)

    # Middle: Allocation engine (Cloud Run job) & Cloud SQL
    job_box = container(s, "run", 3.75, 3.35, 3.30, 1.50, "Cloud Run job · if needed", "run", dashed=True)
    al = node(s, 5.40, 3.70, "generic.analytics", "Allocation engine", ["Python / PuLP / CBC"], tile=0.50, w=2.8)

    sq = node(s, 10.92, 3.70, "sql", "Cloud SQL", ["PostgreSQL / oil ledger / history"], tile=0.52, w=2.8)

    # Bottom Row: Storage, Databricks, Secret Manager
    st = node(s, 5.40, 5.15, "storage", "Cloud Storage", ["private uploads / exports"], tile=0.45, w=2.6)
    db = node(s, 8.80, 5.15, "dbx", "Databricks", ["Optional Python operations", "/ calculations"], tile=0.45, w=2.0)
    sm = node(s, 11.20, 5.15, "secret", "Secret Manager", ["server-side secrets · proposed"], tile=0.45, w=2.2)

    # Connections (Clean Flat, NO SHADOWS)
    # User to React
    seg(s, u_box["r"], rx["cy"], rx["l"], rx["cy"], color=ARROW, pt=1.2)
    lab(s, u_box["r"] + 0.10, rx["cy"] - 0.20, "load UI", 0.9, align=PP_ALIGN.LEFT)

    # Entra ID to React
    path(s, [
        (ent["r"], ent["cy"]),
        (3.38, ent["cy"]),
        (3.38, 3.15),
        (rx["cx"], 3.15),
        (rx["cx"], rx["bot"])
    ], color=ARROW2, pt=1.1)
    lab(s, 3.50, 2.98, "SSO / PKCE", 1.0, color=ARROW2, align=PP_ALIGN.LEFT)

    # Entra ID to FastAPI
    path(s, [
        (ent["r"], ent["cy"] + 0.18),
        (3.28, ent["cy"] + 0.18),
        (3.28, 6.90),
        (12.80, 6.90),
        (12.80, fa["cy"] + 0.20),
        (fa["r"] + 0.20, fa["cy"] + 0.20)
    ], color=ARROW2, pt=1.1)
    lab(s, 8.80, 4.65, "issuer / signing keys", 2.5, color=ARROW2)

    # React to FastAPI
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"], color=ARROW, pt=1.2)
    lab(s, 6.80, fa["cy"] - 0.20, "API + token", 3.0)

    # FastAPI to Allocation engine (launch / status)
    path(s, [
        (9.40, 3.07),
        (9.40, 3.25),
        (7.40, 3.25),
        (7.40, al["cy"]),
        (job_box["r"], al["cy"])
    ], color=ARROW, pt=1.2)
    lab(s, 7.50, 3.08, "launch / status", 1.5, align=PP_ALIGN.LEFT)

    # Allocation engine to Cloud SQL (read / write ledger)
    seg(s, job_box["r"], al["cy"] + 0.20, sq["l"] - 0.15, sq["cy"] + 0.20, color=ARROW, pt=1.2, start=True)
    lab(s, 7.40, al["cy"] + 0.05, "read / write ledger", 2.2)

    # FastAPI to Cloud SQL
    seg(s, fa["cx"], 3.07, fa["cx"], sq["t"], color=ARROW, pt=1.2)
    lab(s, fa["cx"] + 0.08, 3.25, "SQL", 0.6, align=PP_ALIGN.LEFT)

    # Databricks optional jobs to FastAPI
    path(s, [
        (db["cx"], db["t"]),
        (db["cx"], 2.80),
        (fa["l"] - 0.20, 2.80)
    ], color=ARROW2, pt=1.0, dash=True)
    lab(s, 7.55, 2.65, "optional jobs", 1.5, color=ARROW2, align=PP_ALIGN.RIGHT)

    # FastAPI to Storage
    path(s, [
        (9.60, 3.07),
        (9.60, 5.00),
        (st["cx"], 5.00),
        (st["cx"], st["t"])
    ], color=ARROW2, pt=1.1)
    lab(s, 6.70, 4.80, "files", 1.0, color=ARROW2)

    # FastAPI to Secret Manager
    path(s, [
        (fa["r"] + 0.05, 2.45),
        (12.70, 2.45),
        (12.70, 5.35),
        (sm["r"] + 0.15, 5.35)
    ], color=ARROW2, pt=1.1)
    lab(s, 11.85, 4.70, "secrets", 0.9, color=ARROW2)

    # CI/CD deployment bus
    path(s, [
        (cicd["r"], cicd["cy"]),
        (3.18, cicd["cy"]),
        (3.18, 1.40),
        (12.75, 1.40),
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    seg(s, fe_box["cx"], 1.40, fe_box["cx"], fe_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    seg(s, api_box["cx"], 1.40, api_box["cx"], api_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    lab(s, 6.0, 1.25, "deploy frontend + API", 2.5, color=ARROW2)

    path(s, [
        (cicd["r"], cicd["b"] - 0.15),
        (3.18, cicd["b"] - 0.15),
        (3.18, 7.15),
        (12.75, 7.15),
        (12.75, 1.40)
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    lab(s, 6.8, 7.00, "deploy allocation / optional jobs if retained", 3.8, color=ARROW2)


# ============================================================================
# SLIDE 4: Impact Analysis / Proposed GCP architecture (Image 4 recreation)
# ============================================================================
def slide_impact_analysis(prs: Presentation):
    s = new_slide(
        prs,
        "Impact Analysis / Proposed GCP architecture",
        "Compare planning cycles · Understand cost exposure · Simulate future prices · Azure DevOps CI/CD · Phase 0 proposal",
    )

    # Left Column: Users · proposed action mapping
    u_box = container(s, "users", 0.35, 1.05, 2.85, 2.90, "Users · proposed action mapping", "generic.user")
    card(s, 0.48, 1.48, 2.58, 0.70, "Base user", ["Leadership / finance: view financial impact"], icon="generic.user", title_pt=8.5)
    card(s, 0.48, 2.24, 2.58, 0.70, "Super user", ["Review baseline and run what-if simulations"], icon="generic.user", title_pt=8.5)
    card(s, 0.48, 3.00, 2.58, 0.80, "Admin / data owner", [
        "Maintain assumptions, formulas and access",
        "Final permissions to be confirmed"
    ], icon="generic.user", title_pt=8.5)

    # Left Column: Azure · enterprise identity
    az_box = container(s, "azure", 0.35, 4.05, 2.85, 1.55, "Azure · enterprise identity", "azure")
    ent = node(s, 1.77, 4.38, "entra", "Microsoft Entra ID", [
        "SSO · SPNs · AD groups"
    ], tile=0.48, w=2.7)

    # Left Column: Azure DevOps
    do_box = container(s, "azure", 0.35, 5.70, 2.85, 1.52, "Azure DevOps", "azure.devops")
    cicd = card(s, 0.48, 6.08, 2.58, 1.02, "Common CI/CD", [
        [("Azure Repos: ", 7.0, True, INK, False), ("React · FastAPI · Databricks code", 7.0, False, MUTED, False)],
        [("Azure Pipelines: ", 7.0, True, INK, False), ("Build · checks · package", 7.0, False, MUTED, False)],
        [("Release stages: ", 7.0, True, INK, False), ("Deploy app + calculation jobs\nEnvironment gates / rollback¹", 7.0, False, MUTED, False)],
    ], title_pt=8.5)

    # Right Container: Google Cloud · application & data
    gcp_box = container(s, "gcp", 3.42, 1.05, 9.56, 6.17, "Google Cloud · application & data", "gcloud")

    # Cloud Run frontend & backend containers
    fe_box = container(s, "run", 3.68, 1.56, 3.40, 1.62, "Cloud Run · frontend", "run", dashed=True)
    rx = node(s, 5.38, 1.94, "react", "React", ["Dashboards / inputs / simulation"], tile=0.55, w=3.1)

    api_box = container(s, "run", 9.18, 1.56, 3.55, 1.62, "Cloud Run · backend", "run", dashed=True)
    fa = node(s, 10.95, 1.94, "fastapi", "FastAPI", ["Validation / roles / job orchestration"], tile=0.55, w=3.2)

    # Databricks · Python calculation jobs
    dbx_box = container(s, "dbx", 3.68, 3.35, 9.05, 1.45, "Databricks · Python calculation jobs", "dbx")
    mods = [
        ("Validate inputs", ["Required values / units", "Formula & mapping checks"]),
        ("Derive prices", ["Feedstock assumptions", "→ commodity prices"]),
        ("Calculate impact", ["Covered + uncovered", "Current vs prior cycle"]),
        ("Simulate", ["Hypothetical prices", "Compare with baseline"])
    ]
    dbx_cards = []
    for i, (t, ls) in enumerate(mods):
        c = card(s, 3.82 + i * 2.22, 3.75, 2.12, 0.92, t, ls, title_pt=8.5)
        dbx_cards.append(c)

    # Data Row: Cloud SQL & Cloud Storage
    sq = node(s, 5.38, 5.15, "sql", "Cloud SQL · PostgreSQL", ["Snapshots · assumptions · formulas"], tile=0.50, w=3.2)
    st = node(s, 10.95, 5.15, "storage", "Cloud Storage", ["Private file archives / exports", "Proposed storage responsibility"], tile=0.50, w=3.2)

    # Dependency card: Physical Covers
    pc_dep = card(s, 3.68, 6.25, 9.05, 0.65, "Physical Covers · demand vs secured purchases", [
        "Supplies saved demand, purchased quantities and prices for comparison between planning cycles."
    ], icon="generic.document", title_pt=9.0)

    # Connections (Clean Flat, NO SHADOWS)
    # Users to React
    seg(s, u_box["r"], rx["cy"], rx["l"], rx["cy"], color=ARROW, pt=1.2)
    lab(s, u_box["r"] + 0.15, rx["cy"] - 0.20, "HTTPS", 0.8, align=PP_ALIGN.LEFT)

    # Entra ID to React
    path(s, [
        (ent["r"], ent["cy"]),
        (3.30, ent["cy"]),
        (3.30, 3.35),
        (rx["cx"], 3.35),
        (rx["cx"], rx["bot"])
    ], color=ARROW2, pt=1.1)
    lab(s, 3.45, 3.18, "SSO / PKCE", 1.0, color=ARROW2, align=PP_ALIGN.LEFT)

    # React to FastAPI
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"], color=ARROW, pt=1.2, start=True)
    lab(s, 6.70, fa["cy"] - 0.20, "API + access token", 3.5)

    # FastAPI to Databricks
    seg(s, fa["cx"] - 0.50, 3.17, fa["cx"] - 0.50, dbx_box["t"], color=ARROW, pt=1.2, start=True)
    lab(s, fa["cx"] - 0.40, 3.22, "jobs / status", 1.2, align=PP_ALIGN.LEFT)

    # Databricks to Cloud SQL
    seg(s, sq["cx"], dbx_box["b"], sq["cx"], sq["t"], color=ARROW, pt=1.2, start=True)
    lab(s, sq["cx"] + 0.08, 4.90, "read / write", 1.1, align=PP_ALIGN.LEFT)

    # Databricks to Cloud Storage
    seg(s, st["cx"], dbx_box["b"], st["cx"], st["t"], color=ARROW2, pt=1.2, start=True)
    lab(s, st["cx"] + 0.08, 4.90, "archive / export", 1.4, color=ARROW2, align=PP_ALIGN.LEFT)

    # Physical Covers to Cloud SQL (Arrow points up towards Cloud SQL)
    seg(s, sq["cx"], pc_dep["t"], sq["cx"], sq["bot"] + 0.05, color=ARROW, pt=1.2)
    lab(s, sq["cx"] + 0.08, 6.00, "cycle snapshot", 1.4, align=PP_ALIGN.LEFT)

    # Azure DevOps deploy bus (routed cleanly above Cloud Run FE/backend boxes)
    path(s, [
        (cicd["r"], cicd["cy"]),
        (3.10, cicd["cy"]),
        (3.10, 1.40),
        (12.85, 1.40),
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    seg(s, fe_box["cx"], 1.40, fe_box["cx"], fe_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    seg(s, api_box["cx"], 1.40, api_box["cx"], api_box["t"], color=ARROW_DASH, pt=1.0, dash=True)
    lab(s, 6.0, 1.25, "deploy frontend + API", 2.5, color=ARROW2)

    path(s, [
        (cicd["r"], cicd["b"] - 0.15),
        (3.10, cicd["b"] - 0.15),
        (3.10, 7.15),
        (12.85, 7.15),
        (12.85, 3.80)
    ], color=ARROW_DASH, pt=1.0, end=False, dash=True)
    lab(s, 6.8, 7.00, "deploy calculation jobs / configuration", 3.8, color=ARROW2)


# ============================================================================
# Main Entrypoint: Build all 4 slides into ONE presentation deck
# ============================================================================
def build(out_path: str = "out/proposed_architectures.pptx") -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    # 4 Slides in one PPT
    slide_dt_dso_common(prs)
    slide_physical_covers(prs)
    slide_virtual_tank(prs)
    slide_impact_analysis(prs)
    
    dest = Path(out_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    import os; os.makedirs("out", exist_ok=True)
    from archgen.shared.geomcheck import fix_presentation
    fix_presentation(prs)  # hug text boxes, snap near-miss alignment
    prs.save(str(dest))
    print(f"Saved complete 4-slide deck to: {dest}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "out/proposed_architectures.pptx"
    build(out)
