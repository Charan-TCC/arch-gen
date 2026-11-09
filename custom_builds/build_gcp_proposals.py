"""Four 'Proposed GCP architecture' slides in one deck, one soft pastel theme.

Slides: Data Transformer + DSO, Impact Analysis, Physical Covers, Virtual Tank.
Native shapes only (no shadows); every node is icon + label grouped into one shape.
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from archgen.registry import emf
from archgen.registry.resolver import IconResolver

ROOT = Path(__file__).resolve().parent.parent
ICONS = ROOT / "assets" / "icons"
R = IconResolver()
FONT = "Segoe UI"

# ---- theme (soft pastels, flat) -------------------------------------------
PAGE = "FBFCFE"
INK, MUTED = "1E293B", "64748B"
ARROW, ARROW2 = "5F86C9", "94A3B8"
CARD_LINE = "DCE3EE"
THEMES = {  # fill, line, label colour  (one blue-slate family + soft violet/indigo accents)
    "users": ("EEF3FB", "C9D6EC", "2F4F8F"),
    "azure": ("EAF4F8", "BCD9E6", "1F6F8B"),
    "gcp": ("F7F9FC", "D3DBE8", "2F4F8F"),
    "run": ("F3F4FB", "B9BFE3", "4B4F9E"),
    "mods": ("EEF3FB", "C9D6EC", "2F4F8F"),
    "dbx": ("F5F2FC", "D5CBEF", "5B45A3"),
    "grey": ("F1F5F9", "DCE3EE", "475569"),
}
SUBHEAD = "2F4F8F"


def C(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def flat(shp):
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    shp._element.spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))


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
        for t, size, bold, color, mono in para:
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


def key_path(key: str) -> Path:
    return R.get(key).path


BRAND = {"databricks.svg": "#FF3621"}  # single-path logo: apply official brand colour

TECH = {
    "react": ICONS / "tech" / "React.svg",
    "fastapi": ICONS / "tech" / "FastAPI.svg",
    "gcloud": ICONS / "tech" / "Google-Cloud.svg",
    "run": ICONS / "gcp" / "Cloud-Run.svg",
    "sql": ICONS / "gcp" / "Cloud-SQL.svg",
    "storage": ICONS / "gcp" / "Cloud-Storage.svg",
    "secret": ICONS / "gcp" / "Secret-Manager.svg",
    "azure": ICONS / "tech" / "Azure.svg",
    "dbx": ICONS / "tech" / "databricks.svg",
    "python": ICONS / "tech" / "Python.svg",
}


def ico(k: str) -> Path:
    return TECH[k] if k in TECH else key_path(k)


# ---- composite shapes ------------------------------------------------------
def container(slide, kind, x, y, w, h, title, icon=None, dashed=False):
    fill, line, lab = THEMES[kind]
    rrect(slide.shapes, x, y, w, h, fill, line, pt=1.0, radius=0.035, dashed=dashed)
    tx = x + 0.15
    if icon:
        pic(slide.shapes, ico(icon), x + 0.28, y + 0.24, 0.26)
        tx = x + 0.48
    text(slide.shapes, tx, y + 0.11, w - 0.6, 0.28, [[(title, 10.5, True, lab, False)]], anchor=MSO_ANCHOR.MIDDLE)


def node(slide, cx, y, icon, title, sub, tile=0.62, w=2.6):
    """Icon tile with title + mono subtitle underneath, one group. Returns anchors."""
    g = slide.shapes.add_group_shape()
    rrect(g.shapes, cx - tile / 2, y, tile, tile, "FFFFFF", CARD_LINE, radius=0.18)
    pic(g.shapes, ico(icon), cx, y + tile / 2, tile * 0.56)
    paras = [[(title, 10, True, INK, False)]]
    for s in sub:
        paras.append([(s, 7.5, False, MUTED, True)])
    text(g.shapes, cx - w / 2, y + tile + 0.06, w, 0.2 + 0.15 * len(sub) + 0.05, paras, align=PP_ALIGN.CENTER)
    return dict(l=cx - tile / 2, r=cx + tile / 2, t=y, b=y + tile, cx=cx, cy=y + tile / 2,
                bot=y + tile + 0.3 + 0.15 * len(sub))


def card(slide, x, y, w, h, title, lines, icon=None, mono=False, fill="FFFFFF", title_pt=9.5):
    """White card with optional icon left of the title; body text underneath."""
    g = slide.shapes.add_group_shape()
    rrect(g.shapes, x, y, w, h, fill, CARD_LINE, radius=0.07)
    tx = x + 0.12
    if icon:
        pic(g.shapes, ico(icon), x + 0.27, y + 0.22, 0.24)
        tx = x + 0.45
    text(g.shapes, tx, y + 0.09, w - (tx - x) - 0.08, 0.26, [[(title, title_pt, True, INK, False)]],
         anchor=MSO_ANCHOR.MIDDLE)
    paras = [[(ln, 7.5, False, MUTED, mono)] for ln in lines]
    text(g.shapes, x + 0.12, y + 0.40, w - 0.22, h - 0.44, paras)
    return dict(l=x, r=x + w, t=y, b=y + h, cx=x + w / 2, cy=y + h / 2)


def seg(slide, x1, y1, x2, y2, color=ARROW, pt=1.25, end=True, start=False, dash=False):
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
    return cn


def path(slide, pts, color=ARROW, pt=1.25, end=True, start=False, dash=False):
    n = len(pts) - 1
    for i in range(n):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        seg(slide, x1, y1, x2, y2, color, pt, end=end and i == n - 1, start=start and i == 0, dash=dash)


def lab(slide, x, y, s, w=1.6, color=ARROW, align=PP_ALIGN.CENTER):
    text(slide.shapes, x, y, w, 0.18, [[(s, 7.5, False, color, True)]], align=align)


def header(slide, title, subtitle):
    # soft accent bar + title block, identical on every slide
    rrect(slide.shapes, 0.35, 0.36, 0.09, 0.62, "9DB7F0", "9DB7F0", pt=0.1, radius=0.5)
    text(slide.shapes, 0.55, 0.28, 11.5, 0.45, [[(title, 22, True, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    text(slide.shapes, 0.55, 0.72, 12.0, 0.28, [[(subtitle, 10.5, False, MUTED, False)]], anchor=MSO_ANCHOR.MIDDLE)


def footer(slide, y, h, lines):
    rrect(slide.shapes, 0.35, y, 12.63, h, "F1F5F9", "E2E8F0", radius=0.12)
    paras = []
    for head, body in lines:
        para = []
        if head:
            para.append((head + " ", 8, True, INK, False))
        para.append((body, 8, False, MUTED, False))
        paras.append(para)
    text(slide.shapes, 0.5, y + 0.06, 12.3, h - 0.1, paras, anchor=MSO_ANCHOR.MIDDLE)


def new_slide(prs, title, subtitle):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = C(PAGE)
    header(s, title, subtitle)
    return s


ENTRA = "azure.entra-privleged-identity-management"
USERS = "generic.user"


# ============================================================ slides 1-2: DT and DSO (shared layout)
def _dt_like(prs, title, subtitle, users_title, users, react_sub, api_sub, back_label, row, loop, storage, sqlsub, foot):
    s = new_slide(prs, title, subtitle)
    container(s, "users", 0.35, 1.2, 2.55, 2.3, users_title, USERS)
    b1 = card(s, 0.5, 1.7, 2.25, 0.8, users[0][0], [users[0][1]])
    card(s, 0.5, 2.6, 2.25, 0.8, users[1][0], [users[1][1]])
    container(s, "azure", 0.35, 3.7, 2.55, 2.2, "Azure · identity", "azure")
    ent = node(s, 1.62, 4.25, ENTRA, "Microsoft Entra ID", ["SSO / PKCE · app registrations", "SPNs · AD groups · roles"], w=2.4)

    container(s, "gcp", 3.1, 1.2, 9.88, 5.6, "Google Cloud · application & data", "gcloud")
    react = card(s, 3.35, 1.7, 3.2, 1.0, "React", ["Cloud Run · frontend", react_sub], icon="react", title_pt=11)
    api = card(s, 9.5, 1.7, 3.2, 1.0, "FastAPI", ["Cloud Run · API", api_sub], icon="fastapi", title_pt=11)
    seg(s, b1["r"], 2.0, react["l"], 2.0)
    path(s, [(ent["r"] + 0.75, 4.55), (3.0, 4.55), (3.0, 2.4), (react["l"], 2.4)], color=ARROW2)
    seg(s, react["r"], 2.0, api["l"], 2.0)
    lab(s, 6.6, 1.78, "API + token", 2.9)
    seg(s, api["l"], 2.4, react["r"], 2.4, dash=True)
    lab(s, 6.6, 2.46, back_label, 2.9)
    seg(s, api["cx"], api["b"], api["cx"], 3.0)
    lab(s, api["cx"] + 0.1, 2.76, "launch job / check status", 2.0, align=PP_ALIGN.LEFT)

    container(s, "dbx", 3.35, 3.0, 9.4, 2.1, "Databricks · execution", "dbx")
    text(s.shapes, 3.55, 3.42, 6, 0.22, [[(row[0], 9.5, True, SUBHEAD, False)]])
    xs = [3.55, 5.9, 8.25, 10.6]
    for x, (t, ls, ic) in zip(xs, row[1]):
        card(s, x, 3.95, 2.05, 0.95, t, ls, icon=ic, title_pt=8.5)
    for i in range(3):
        seg(s, xs[i] + 2.05, 4.42, xs[i + 1], 4.42, pt=1.0)
    kind, src, dst, label = loop
    sx = xs[src] + 1.02
    dx = xs[dst] + 1.02 if dst is not None else sx
    if kind == "retry":
        path(s, [(sx, 3.95), (sx, 3.78), (dx, 3.78), (dx, 3.95)], color=ARROW2, pt=1.0, dash=True)
        lab(s, 7.3, 3.56, label, 2.4, ARROW2)
    else:  # fail: errors return to the FastAPI card
        path(s, [(sx, 3.95), (sx, 3.78), (10.0, 3.78), (10.0, api["b"])], color=ARROW2, pt=1.0, dash=True)
        lab(s, 7.1, 3.56, label, 3.4, ARROW2, PP_ALIGN.LEFT)

    st = node(s, 5.0, 5.55, "storage", "Cloud Storage", storage, tile=0.45, w=3.0)
    sm = node(s, 7.65, 5.55, "secret", "Secret Manager", ["server-side secrets"], tile=0.45, w=1.9)
    sq = node(s, 10.5, 5.55, "sql", "Cloud SQL · PostgreSQL", sqlsub, tile=0.45, w=3.4)
    seg(s, 5.0, 5.1, 5.0, 5.55)
    lab(s, 5.1, 5.2, "read / write files", 1.4, align=PP_ALIGN.LEFT)
    path(s, [(api["r"] + 0.05, 2.3), (12.92, 2.3), (12.92, 5.77), (sq["r"] + 0.1, 5.77)], color=ARROW2, pt=1.0)
    path(s, [(api["r"] + 0.05, 2.55), (12.82, 2.55), (12.82, 6.58), (sm["cx"], 6.58), (sm["cx"], 6.46)], color=ARROW2, pt=1.0)
    lab(s, 8.5, 6.6, "read secrets", 1.2, ARROW2, PP_ALIGN.LEFT)
    footer(s, 6.92, 0.5, foot)


def slide_dt(prs):
    _dt_like(
        prs, "Data Transformer | Phase 0 - Architecture",
        "React + FastAPI on Cloud Run · Databricks execution · Cloud Storage + Cloud SQL · Phase 0",
        "Users · DT",
        [("Buyer", "Upload · review · download Costbook"), ("KB editor / Support", "Rules · mappings · run investigation")],
        "Upload quotations · download Costbook", "DT jobs · status · Costbook download", "status / Costbook download",
        ("DT · Agentic solution", [
            ("Load quotation inputs", ["Python: Excel / CSV cells", "Images + supplier KB / masters"], "generic.document"),
            ("Multimodal agent", ["Interpret layout + images", "Plan rows · field → source refs"], "generic.search"),
            ("Calculate + map", ["Read cells · rules / FX / maps", "EXW → CIF → DDP"], "generic.analytics"),
            ("Generate Costbook", ["Validate routes / fields / prices", "Costbook ready · Download"], "generic.security")]),
        ("retry", 3, 1, "targeted retry / review"),
        ["Quotation files · KBs · masters · Costbooks", "Validation reports · source evidence"],
        ["DT job / run IDs · status · input versions", "Access · validation · file references"],
        [("Phase 0:", "Upload quotation files in DT → generate and download Costbook. No automatic connection to DSO."),
         ("", "Multimodal agent + deterministic pricing; validation and targeted review remain part of the proposed design.")])


def slide_dso(prs):
    _dt_like(
        prs, "DSO | Phase 0 - Architecture",
        "Downstream Optimizer · React + FastAPI on Cloud Run · Databricks execution · Cloud Storage + Cloud SQL · Phase 0",
        "Users · DSO",
        [("Buyer", "Upload · review · run scenarios"), ("Support / admin", "Input rules · access · run investigation")],
        "Upload DT Costbook + other files", "Process / validate · errors · optimizer jobs", "status / errors / results",
        ("DSO · Process files → validate → optimize", [
            ("Process uploaded files", ["DT Costbook + demand / references", "Enrich prices and route data"], "generic.document"),
            ("Validate + DP ID", ["Persist validated input snapshot", "PASS · ready to run optimizer"], "generic.security"),
            ("Solve scenario", ["DP ID + buyer-selected rules", "Optimize quantities / routes"], "generic.analytics"),
            ("Save results", ["OPT ID · solver status · costs", "Allocations · comparison · export"], "generic.storage")]),
        ("fail", 1, None, "FAIL → errors returned to React via API"),
        ["DT Costbooks · other inputs · DP snapshots", "Scenario datasets · reports · Excel exports"],
        ["Job / run IDs · readiness · DP / OPT metadata", "Scenario settings · access · file references"],
        [("Phase 0:", "Upload DT-generated Costbook + other files in DSO → process / validate → errors return to React on failure; run optimizer on success."),
         ("", "Buyer selects constraints; persist DP ID and OPT ID. Solver outcome is distinct from job completion. No direct DT ingestion.")])


# ============================================================ slide 3: Impact Analysis
def slide_impact(prs):
    s = new_slide(prs, "Impact Analysis | Phase 0 - Architecture",
                  "Compare planning cycles · Understand cost exposure · Simulate future prices · Phase 0 proposal")
    container(s, "users", 0.35, 1.2, 3.0, 3.5, "Users · proposed action mapping", USERS)
    for i, (t, d1, d2) in enumerate([
            ("Base user", "Leadership / finance:", "view financial impact"),
            ("Super user", "Review baseline and run", "what-if simulations"),
            ("Admin / data owner", "Maintain assumptions, formulas", "and access · permissions TBC")]):
        card(s, 0.5, 1.68 + i * 0.98, 2.7, 0.9, t, [d1, d2], icon=USERS)
    container(s, "azure", 0.35, 4.9, 3.0, 2.0, "Azure · enterprise identity", "azure")
    ent = node(s, 1.85, 5.4, ENTRA, "Microsoft Entra ID", ["SSO · SPNs · AD groups"], w=2.7)

    container(s, "gcp", 3.6, 1.2, 9.38, 5.7, "Google Cloud · application & data", "gcloud")
    container(s, "run", 3.9, 1.65, 3.3, 1.85, "Cloud Run · frontend", "run", dashed=True)
    rx = node(s, 5.55, 2.1, "react", "React", ["Dashboards / inputs / simulation"], w=3.2)
    container(s, "run", 9.3, 1.65, 3.45, 1.85, "Cloud Run · backend", "run", dashed=True)
    fa = node(s, 11.02, 2.1, "fastapi", "FastAPI", ["Validation / roles / job orchestration"], w=3.3)
    sm = node(s, 8.25, 2.62, "secret", "Secret Manager", ["read secrets"], tile=0.45, w=1.9)
    seg(s, sm["r"] + 0.02, sm["cy"], 9.3, sm["cy"], color=ARROW2)
    seg(s, 3.35, rx["cy"], rx["l"], rx["cy"])
    lab(s, 3.5, rx["cy"] - 0.2, "HTTPS", 0.7, align=PP_ALIGN.LEFT)
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"], start=True)
    lab(s, 6.7, fa["cy"] - 0.2, "API + access token", 3.5)
    path(s, [(ent["r"] + 0.7, ent["cy"]), (3.48, ent["cy"]), (3.48, 3.65), (rx["cx"], 3.65), (rx["cx"], 3.5)], color=ARROW2)
    lab(s, 3.8, 3.47, "SSO / PKCE", 1.0, ARROW2, PP_ALIGN.LEFT)

    container(s, "dbx", 3.9, 3.85, 8.85, 1.4, "Databricks · Python calculation jobs", "dbx")
    mods = [("Validate inputs", ["Required values / units", "Formula & mapping checks"]),
            ("Derive prices", ["Feedstock assumptions", "→ commodity prices"]),
            ("Calculate impact", ["Covered + uncovered", "Current vs prior cycle"]),
            ("Simulate", ["Hypothetical prices", "Compare with baseline"])]
    for i, (t, ls) in enumerate(mods):
        card(s, 4.05 + i * 2.18, 4.28, 2.05, 0.86, t, ls)
    seg(s, fa["cx"], 3.5, fa["cx"], 3.85, start=True)
    lab(s, fa["cx"] + 0.08, 3.56, "jobs / status", 1.2, align=PP_ALIGN.LEFT)

    sq = node(s, 5.55, 5.5, "sql", "Cloud SQL · PostgreSQL", ["Snapshots · assumptions · formulas"], tile=0.45, w=3.3)
    st = node(s, 11.02, 5.5, "storage", "Cloud Storage", ["Private file archives / exports"], tile=0.45, w=3.3)
    seg(s, 5.55, 5.25, 5.55, 5.5, start=True)
    lab(s, 5.65, 5.3, "read / write", 1.0, align=PP_ALIGN.LEFT)
    seg(s, 11.02, 5.25, 11.02, 5.5, color=ARROW2, start=True)
    lab(s, 11.12, 5.3, "archive / export", 1.3, ARROW2, PP_ALIGN.LEFT)

    card(s, 4.05, 6.46, 8.55, 0.38, "", [], icon="azure.browser")
    text(s.shapes, 4.55, 6.46, 8.0, 0.38, [[("Physical Covers · demand vs secured purchases   ", 9.5, True, INK, False),
                                             ("saved demand, purchased quantities and prices · cycle snapshot → Cloud SQL", 8, False, MUTED, False)]],
         anchor=MSO_ANCHOR.MIDDLE)


# ============================================================ slide 4: Physical Covers
def slide_physical(prs):
    s = new_slide(prs, "Physical Covers | Phase 0 - Architecture",
                  "React on Cloud Run · FastAPI on Cloud Run · Azure SSO (Entra ID) · Phase 0 proposal")
    container(s, "users", 0.35, 1.2, 3.0, 3.35, "User roles · proposed", USERS)
    roles = [("Buyer", "Own purchases & coverage"), ("Procurement manager", "Team purchases & coverage"),
             ("Director", "Aggregated reporting"), ("Administrator", "Demand & reference uploads")]
    for i, (t, d) in enumerate(roles):
        card(s, 0.5, 1.62 + i * 0.72, 2.7, 0.66, t, [d], icon=USERS)
    container(s, "azure", 0.35, 4.7, 3.0, 1.65, "Azure · identity", "azure")
    ent = node(s, 1.85, 5.1, ENTRA, "Microsoft Entra ID", ["SSO · enterprise policies"], tile=0.5, w=2.7)

    container(s, "gcp", 3.6, 1.2, 9.38, 5.15, "Google Cloud · application & data", "gcloud")
    container(s, "run", 3.9, 1.65, 3.3, 1.75, "Cloud Run · FE", "run", dashed=True)
    rx = node(s, 5.55, 2.05, "react", "React", ["UI assets / browser MSAL"], w=3.2)
    container(s, "run", 9.3, 1.65, 3.45, 1.75, "Cloud Run · API", "run", dashed=True)
    fa = node(s, 11.02, 2.05, "fastapi", "FastAPI", ["validation / role checks"], w=3.3)
    seg(s, 3.35, rx["cy"], rx["l"], rx["cy"])
    lab(s, 3.5, rx["cy"] - 0.2, "load UI", 0.7, align=PP_ALIGN.LEFT)
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"])
    lab(s, 6.7, fa["cy"] - 0.2, "API + token", 3.5)
    path(s, [(ent["r"] + 0.7, ent["cy"]), (3.43, ent["cy"]), (3.43, 3.65), (rx["cx"], 3.65), (rx["cx"], 3.4)], color=ARROW2)
    lab(s, 3.8, 3.47, "SSO / PKCE", 1.0, ARROW2, PP_ALIGN.LEFT)
    path(s, [(ent["r"] + 0.7, ent["cy"] + 0.14), (3.53, ent["cy"] + 0.14), (3.53, 5.95), (12.87, 5.95), (12.87, fa["cy"]),
             (12.75, fa["cy"])], color=ARROW2)
    lab(s, 9.4, 5.74, "issuer / signing keys", 3.4, ARROW2, PP_ALIGN.RIGHT)

    st = node(s, 5.0, 4.5, "storage", "Cloud Storage", ["private uploads / exports"], tile=0.5, w=2.6)
    sq = node(s, 8.0, 4.5, "sql", "Cloud SQL", ["PostgreSQL / FIFO / prices"], tile=0.5, w=2.6)
    sm = node(s, 11.0, 4.5, "secret", "Secret Manager", ["server-side secrets"], tile=0.5, w=2.6)
    path(s, [(9.5, 3.4), (9.5, 4.1), (st["cx"], 4.1), (st["cx"], st["t"])], color=ARROW2)
    lab(s, 6.3, 3.88, "files", 1.0, ARROW2)
    path(s, [(9.9, 3.4), (9.9, 4.3), (sq["cx"], 4.3), (sq["cx"], sq["t"])])
    lab(s, 9.95, 3.55, "SQL", 0.6, align=PP_ALIGN.LEFT)
    seg(s, sm["cx"], 3.4, sm["cx"], sm["t"], color=ARROW2)
    lab(s, 11.1, 3.55, "read secrets", 1.2, ARROW2, PP_ALIGN.LEFT)

    footer(s, 6.5, 0.75, [("Processing:", "React handles screens. FastAPI handles authoritative validation, transformations and Python helpers. "
                                           "PostgreSQL handles FIFO, coverage and prices."),
                          ("Scope:", "manual files; preserve Impact Analysis snapshots. Separate FE/API services; private business data. "
                                     "Ingress / domain routing deferred. Azure identity retained; no client secrets in the browser. "
                                     "Final role/action matrix to be confirmed.")])


# ============================================================ slide 5: Virtual Tank
def slide_tank(prs):
    s = new_slide(prs, "Virtual Tank | Phase 0 - Architecture",
                  "React + FastAPI on separate Cloud Run services · Existing allocation job · Azure enterprise identity · Phase 0 proposal")
    container(s, "users", 0.35, 1.15, 3.2, 3.3, "User roles · VT personas", USERS)
    card(s, 0.5, 1.6, 1.4, 1.2, "Oil buyer", ["Purchase uploads", "Reservation requests"], icon=USERS, title_pt=9)
    card(s, 2.0, 1.6, 1.4, 1.2, "Demand planner", ["Demand / consumption", "Review / finalize"], icon=USERS, title_pt=8.5)
    card(s, 0.5, 2.9, 2.9, 0.85, "Portfolio buyer", ["Read-only balances, coverage and prices"], icon=USERS)
    text(s.shapes, 0.5, 3.95, 2.9, 0.3, [[("Final role/action matrix to be confirmed.", 8, False, MUTED, False)]])
    container(s, "azure", 0.35, 4.65, 3.2, 1.7, "Azure · enterprise identity", "azure")
    ent = node(s, 1.95, 5.05, ENTRA, "Microsoft Entra ID", ["SSO · SPNs · AD groups"], tile=0.5, w=2.8)

    container(s, "gcp", 3.8, 1.15, 9.18, 5.2, "Google Cloud · application & data", "gcloud")
    container(s, "run", 4.2, 1.6, 2.9, 1.5, "Cloud Run · FE", "run", dashed=True)
    rx = node(s, 5.65, 2.0, "react", "React", ["UI assets / browser MSAL"], tile=0.5, w=2.8)
    container(s, "run", 9.6, 1.6, 3.1, 1.5, "Cloud Run · API", "run", dashed=True)
    fa = node(s, 11.15, 2.0, "fastapi", "FastAPI", ["validation / roles / ledger APIs"], tile=0.5, w=3.0)
    container(s, "run", 4.2, 3.55, 2.9, 1.45, "Cloud Run job · existing", "run", dashed=True)
    al = node(s, 5.65, 3.95, "generic.analytics", "Allocation engine", ["Python / PuLP / CBC"], tile=0.5, w=2.8)
    sq = node(s, 11.15, 3.95, "sql", "Cloud SQL", ["PostgreSQL / oil ledger / history"], tile=0.5, w=2.3)
    st = node(s, 5.65, 5.3, "storage", "Cloud Storage", ["private uploads / exports"], tile=0.4, w=2.6)
    db = node(s, 9.2, 5.3, "dbx", "Databricks", ["Optional Python operations", "/ calculations"], tile=0.4, w=1.8)
    sm = node(s, 11.4, 5.3, "secret", "Secret Manager", ["server-side secrets"], tile=0.4, w=1.9)

    seg(s, 3.55, rx["cy"], rx["l"], rx["cy"])
    lab(s, 3.6, rx["cy"] - 0.22, "load UI", 0.8, align=PP_ALIGN.LEFT)
    seg(s, rx["r"], fa["cy"], fa["l"], fa["cy"])
    lab(s, 7.2, fa["cy"] - 0.22, "API + token", 2.2)
    path(s, [(ent["r"] + 0.3, ent["cy"]), (3.68, ent["cy"]), (3.68, 3.38), (rx["cx"], 3.38), (rx["cx"], 3.1)], color=ARROW2)
    lab(s, 3.85, 3.17, "SSO / PKCE", 1.0, ARROW2, PP_ALIGN.LEFT)
    path(s, [(9.75, 3.1), (9.75, 3.38), (7.5, 3.38), (7.5, 4.2), (7.1, 4.2)])
    lab(s, 7.6, 3.17, "launch / status", 1.4, align=PP_ALIGN.LEFT)
    seg(s, 7.2, 4.4, sq["l"] - 0.1, 4.4, start=True)
    lab(s, 7.2, 4.18, "read / write ledger", 1.8)
    seg(s, fa["cx"], 3.1, fa["cx"], sq["t"])
    lab(s, fa["cx"] + 0.08, 3.4, "SQL", 0.6, align=PP_ALIGN.LEFT)
    path(s, [(9.95, 3.1), (9.95, 5.12), (st["cx"], 5.12), (st["cx"], st["t"])], color=ARROW2)
    lab(s, 6.9, 4.93, "files", 0.8, ARROW2)
    path(s, [(db["cx"], db["t"]), (db["cx"], 2.85), (9.6, 2.85)], color=ARROW2, dash=True)
    lab(s, 7.65, 2.7, "optional jobs", 1.45, ARROW2, PP_ALIGN.RIGHT)
    path(s, [(12.7, 2.4), (12.85, 2.4), (12.85, 5.5), (sm["r"] + 0.1, 5.5)], color=ARROW2)
    lab(s, 11.9, 5.1, "secrets", 0.9, ARROW2)

    footer(s, 6.55, 0.75, [("Processing:", "React renders screens; FastAPI handles validation, conversions and ledger workflows; "
                                           "the existing allocation job selects purchases for quantity / price."),
                           ("Design:", "manual Excel / CSV inputs; no new SAP / Anaplan integration. Ingress / domain routing deferred. "
                                       "No browser client secrets. Logical proposal, not deployed topology.")])


def build(out: str = "out/gcp_proposed_architectures.pptx") -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for fn in (slide_dt, slide_dso, slide_impact, slide_physical, slide_tank):
        fn(prs)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print("saved", out)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "out/gcp_proposed_architectures.pptx")
