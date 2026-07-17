"""Bespoke high-fidelity build of the Heimdall end-to-end architecture.

Six labelled layer bands (swimlanes) from source systems down to the Heimdall
web app, using archgen's real icon assets so branding stays authentic and the
output stays native/editable. See CLAUDE.md sec5.

    .venv/bin/python custom_builds/build_heimdall.py
"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

from archgen.registry.resolver import IconResolver
from archgen.registry import emf

R = IconResolver()

# --- Palette ---
INK = "0F172A"; GRAY = "64748B"; WHITE = "FFFFFF"
BAND_FILL = "F8FAFC"; BAND_LINE = "E2E8F0"
FLOW = "2563EB"          # main data-flow arrows
READ = "94A3B8"          # dashed read / query arrows
# Per-layer tab palette: (fill, line, text)
TAB = {
    "sources":  ("E2E8F0", "475569", "334155"),
    "ingest":   ("DBEAFE", "2563EB", "1E40AF"),
    "logs":     ("E0F2FE", "0284C7", "075985"),
    "data":     ("E0E7FF", "4F46E5", "3730A3"),
    "fair":     ("FEF3C7", "D97706", "92400E"),
    "app":      ("DCFCE7", "16A34A", "166534"),
}

def C(h): return RGBColor.from_string(h)

def icon(slide, key, cx, cy, size):
    e = R.get(key)
    recolor = f"#{e.brand_hex}" if e.recolorable and e.brand_hex else None
    png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)
    asp = emf.svg_aspect(e.path)
    if asp >= 1:
        w = size; h = int(size / asp)
    else:
        h = size; w = int(size * asp)
    slide.shapes.add_picture(str(png), int(cx - w / 2), int(cy - h / 2), width=w, height=h)

def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None: shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))

def rrect(slide, x, y, w, h, fill, line, line_pt=1.0, radius=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try: shp.adjustments[0] = radius
    except Exception: pass
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
    _no_shadow(shp)
    return shp

def text(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap; tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(9525); tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 0.95
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    return box

def _arrow(ln, end=True, start=False):
    if end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))

def conn(slide, x1, y1, x2, y2, color=FLOW, pt=1.75, dashed=False, arrow_end=True):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed: ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    _arrow(ln, end=arrow_end)
    cn.shadow.inherit = False
    return cn

def elbow(slide, pts, color=READ, pt=1.5, dashed=True):
    """Multi-segment orthogonal route; arrowhead on the final segment only."""
    for i in range(len(pts) - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        conn(slide, x1, y1, x2, y2, color=color, pt=pt, dashed=dashed, arrow_end=(i == len(pts) - 2))

def elabel(slide, x, y, txt, color=GRAY, w=1.1):
    text(slide, x - w / 2, y - 0.12, w, 0.24, [[(txt, 8.5, False, color)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- geometry
SW, SH = 13.333, 7.5
BX, BW = 0.35, 12.65        # band x, width
BH, GAP = 1.0, 0.08         # band height, gap
TOP0 = 0.78                 # first band top
C1, C2, C3, C4 = 3.55, 6.25, 8.95, 11.55   # node column centres
TAB_X, TAB_W = 0.45, 1.55

BANDS = ["sources", "ingest", "logs", "data", "fair", "app"]
BAND_TITLE = {
    "sources": "Data\nSources", "ingest": "Ingestion\nLayer",
    "logs": "Log\nIngestion", "data": "Data\nLayer",
    "fair": "FAIR\nAnalysis", "app": "Heimdall\nApp",
}
top = {b: TOP0 + i * (BH + GAP) for i, b in enumerate(BANDS)}

def cy(b): return top[b] + 0.30            # icon centre height (upper part of band)
def band_center(b): return top[b] + 0.50

prs = Presentation()
prs.slide_width = Inches(SW); prs.slide_height = Inches(SH)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Title
text(slide, 0.5, 0.14, 11.0, 0.5,
     [[("Heimdall — End-to-End Data & Governance Architecture", 22, True, INK)]],
     align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

# Bands + left tabs
for b in BANDS:
    rrect(slide, BX, top[b], BW, BH, BAND_FILL, BAND_LINE, 1.0, radius=0.04)
    f, l, t = TAB[b]
    rrect(slide, TAB_X, top[b] + 0.13, TAB_W, BH - 0.26, f, l, 1.0, radius=0.10)
    text(slide, TAB_X, top[b] + 0.13, TAB_W, BH - 0.26,
         [[(line, 11.5, True, t)] for line in BAND_TITLE[b].split("\n")],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- nodes
# (band, cx, icon_key, label)
NODES = {
    "github":       ("sources", C1, "tech.github",      "GitHub"),
    "harness":      ("sources", C2, "generic.pipeline", "Harness"),
    "aiagents":     ("sources", C3, "generic.workflow", "AI Agents"),
    "runtime":      ("sources", C4, "generic.service",  "Runtime"),

    "sandbox":      ("ingest", C1,   "generic.container", "Sandbox Container"),
    "codescan":     ("ingest", 5.15, "generic.security",  "Code Scan Engine"),

    "cslogs":       ("logs", C1, "generic.document",   "Code Scan Logs"),
    "hlogs":        ("logs", C2, "generic.document",   "Harness Logs"),
    "ailogs":       ("logs", C3, "generic.monitoring", "AI Agent Obs Logs"),
    "rtlogs":       ("logs", C4, "generic.document",   "Runtime Logs"),

    "lake":         ("data", C1, "generic.folder",    "Data Lake"),
    "databricks":   ("data", C2, "tech.databricks",   "Databricks Job"),
    "neo4j":        ("data", C3, "tech.neo4j",        "Neo4j Graph DB"),
    "sqldb":        ("data", C4, "generic.database",  "SQL Database"),

    "fairagent":    ("fair", C1, "generic.function", "FAIR Agent"),
    "trigger":      ("fair", C2, "generic.workflow", "Workflow Trigger"),
    "servicenow":   ("fair", C3, "generic.support",  "ServiceNow"),
    "email":        ("fair", C4, "generic.email",    "Email Alerts"),

    "users":        ("app", C1, "generic.user",     "Users"),
    "auth":         ("app", C2, "generic.identity", "Authentication"),
    "react":        ("app", C3, "tech.react",       "React Frontend"),
    "django":       ("app", C4, "tech.django",      "Django Backend"),
}

for nid, (b, cx, key, lbl) in NODES.items():
    icon(slide, key, Inches(cx), Inches(cy(b)), Inches(0.46))
    text(slide, cx - 1.25, top[b] + 0.55, 2.5, 0.4, [[(lbl, 9, True, INK)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

# ---------------------------------------------------------------- helpers for anchors
def N(nid): return NODES[nid]
def bot(nid):  b, cx, *_ = N(nid); return (cx, top[b] + 0.95)
def topa(nid): b, cx, *_ = N(nid); return (cx, top[b] + 0.03)
def L(nid):    b, cx, *_ = N(nid); return (cx - 0.75, cy(b))
def Rt(nid):   b, cx, *_ = N(nid); return (cx + 0.75, cy(b))

def down(src, dst, color=FLOW, dashed=False, label=None):
    x1, y1 = bot(src); x2, y2 = topa(dst)
    conn(slide, x1, y1, x2, y2, color=color, dashed=dashed)
    if label:
        elabel(slide, (x1 + x2) / 2, (y1 + y2) / 2, label, color=(FLOW if not dashed else GRAY))

def across(src, dst, color=FLOW, label=None):
    x1, y1 = Rt(src); x2, y2 = L(dst)
    conn(slide, x1, y1, x2, y2, color=color)
    if label:
        elabel(slide, (x1 + x2) / 2, y1 - 0.16, label, color=FLOW, w=0.9)

# ---------------------------------------------------------------- edges
# Sources -> Ingestion / Logs
down("github", "sandbox", label="clone")
across("sandbox", "codescan")
# code scan engine -> code scan logs (diagonal down-left)
x1, y1 = bot("codescan"); x2, y2 = topa("cslogs")
conn(slide, x1, y1, x2, y2, color=FLOW)
down("harness", "hlogs")
down("aiagents", "ailogs")
down("runtime", "rtlogs")

# Logs -> Data Lake (fan-in)
for lg in ("cslogs", "hlogs", "ailogs", "rtlogs"):
    x1, y1 = bot(lg); x2, y2 = topa("lake")
    conn(slide, x1, y1, x2, y2, color=FLOW)

# Data layer internal
across("lake", "databricks", label="transform")
across("databricks", "neo4j")
# databricks -> sqldb (skip neo4j: route slightly below the icon row)
x1, y1 = Rt("databricks"); x2, y2 = L("sqldb")
yb = cy("data") + 0.02
conn(slide, x1, yb, x2, yb, color=FLOW)

# Data -> FAIR (dashed reads to FAIR Agent)
for src in ("neo4j", "sqldb"):
    x1, y1 = bot(src); x2, y2 = topa("fairagent")
    conn(slide, x1, y1, x2, y2, color=READ, dashed=True)

# FAIR internal
across("fairagent", "trigger")
across("trigger", "servicenow", label="ticket")
# trigger -> email (skip servicenow: route below icon row)
x1, y1 = Rt("trigger"); x2, y2 = L("email")
yb = cy("fair") + 0.02
conn(slide, x1, yb, x2, yb, color=FLOW)
elabel(slide, (x1 + x2) / 2, yb - 0.16, "notify", color=FLOW, w=0.9)

# FAIR results -> Heimdall backend
x1, y1 = bot("fairagent"); x2, y2 = topa("django")
conn(slide, x1, y1, x2, y2, color=FLOW)
elabel(slide, C2, (y1 + y2) / 2, "FAIR results", color=FLOW, w=1.3)

# Heimdall app request chain
across("users", "auth", label="login")
across("auth", "react")
across("react", "django")

# Django reads Neo4j + SQL (dashed, routed down the right margin)
MX = 12.95
xn, yn = bot("neo4j"); xs, ys = bot("sqldb")
_, dtx = N("django")[0], N("django")[1]
dj_top = top["app"] + 0.03
elbow(slide, [(xs, ys), (MX, ys), (MX, dj_top - 0.18), (dtx + 0.35, dj_top - 0.18), (dtx + 0.35, dj_top)],
      color=READ, dashed=True)
elbow(slide, [(xn, yn), (MX - 0.25, yn + 0.12), (MX - 0.25, dj_top - 0.05), (dtx - 0.15, dj_top - 0.05), (dtx - 0.15, dj_top)],
      color=READ, dashed=True)
text(slide, MX - 2.4, dj_top - 0.55, 2.3, 0.22, [[("query Neo4j + SQL", 8.5, False, GRAY)]],
     align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- footer legend
fy = 7.2
text(slide, 0.5, fy, 8.5, 0.25,
     [[("Solid = data flow    ", 8.5, True, FLOW), ("Dashed = read / query    ", 8.5, True, GRAY),
       ("FAIR = Factor Analysis of Information Risk", 8.5, False, GRAY)]],
     align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

out = sys.argv[1] if len(sys.argv) > 1 else "out/heimdall.pptx"
import os
os.makedirs(os.path.dirname(out), exist_ok=True)
prs.save(out)
print("saved", out)
