"""Bespoke build of the Azure Modern Data Platform reference architecture.

A six-tier layered diagram (Source -> Ingest -> Storage -> Process -> Serve ->
Business) is a hand-designed composition the generic `flow` layout can't render
legibly: the 6-node source column fixes the graph height, so uniform scaling
shrinks every card until labels break mid-word. Here we place each card at an
absolute size, so the sources sit in a compact left rail while the services keep
room for an icon + label. Icons come from archgen's real registry, kept native.
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
INK = "1F2937"; GRAY = "6B7280"; WHITE = "FFFFFF"
AZ_BLUE = "0F6CBD"                     # Azure accent
AZ_ZONE_FILL = "EEF4FB"; AZ_ZONE_LINE = "CBDDF0"
GEN_ZONE_FILL = "F4F5F7"; GEN_ZONE_LINE = "E1E4EA"
CARD_FILL = "FFFFFF"; CARD_LINE = "DBE1EA"
FLOW = "5B7183"                        # connector color
FLOW_AI = "8A6FB0"                     # AI-enrichment connector

def C(h): return RGBColor.from_string(h)

def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))

def rrect(slide, x, y, w, h, fill, line, line_pt=1.0, radius=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
    _no_shadow(shp)
    return shp

def text(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(6350); tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    return box

def icon(slide, key, left_in, top_in, size_in):
    e = R.get(key)
    recolor = f"#{e.brand_hex}" if (e.recolorable and e.brand_hex) else None
    png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)
    asp = emf.svg_aspect(e.path)
    if asp >= 1:
        w = size_in; h = size_in / asp
    else:
        h = size_in; w = size_in * asp
    slide.shapes.add_picture(str(png), Inches(left_in + (size_in - w) / 2),
                             Inches(top_in + (size_in - h) / 2), width=Inches(w), height=Inches(h))

def connector(slide, x1, y1, x2, y2, color=FLOW, pt=1.5, dashed=False, elbow=True):
    kind = MSO_CONNECTOR.ELBOW if (elbow and x1 != x2 and y1 != y2) else MSO_CONNECTOR.STRAIGHT
    cn = slide.shapes.add_connector(kind, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    cn.shadow.inherit = False
    return cn

def edge_label(slide, cx, cy, txt, color=GRAY):
    w = 0.9
    box = slide.shapes.add_textbox(Inches(cx - w / 2), Inches(cy - 0.11), Inches(w), Inches(0.22))
    tf = box.text_frame; tf.word_wrap = False; tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = txt
    r.font.size = Pt(8); r.font.italic = True; r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    # opaque backing so the label reads over the connector
    r2 = None

# ---------------------------------------------------------------- geometry
CARD_W, CARD_H = 1.72, 0.78
# Tier zones (left edges) — six evenly spaced columns.
ZONE_W = 1.95
ZONE_X = {"source": 0.35, "ingest": 2.48, "storage": 4.61,
          "process": 6.74, "serve": 8.87, "business": 11.0}
ZONE_TOP, ZONE_H = 0.95, 5.75

# Per-node card top-Y (hand-tuned so connectors run through clear channels).
# Key routing constraints:
#   * eventhub -> stream bypasses Storage, so the Storage tier's TOP band is
#     kept empty for that line to pass through.
#   * synapse -> aas must cross the Process tier, so Process has a clear channel
#     (~y4.3) between Cognitive Services and Databricks; synapse and aas sit at
#     that same height so the edge stays level.
NODES = {
    # id:        (zone,       card_top_y, icon_key,                     label)
    "traffic":  ("source",   1.42, "generic.internet",            "Online Traffic"),
    "iot":      ("source",   2.28, "azure.iot-hub",               "IoT Devices"),
    "mobile":   ("source",   3.14, "generic.mobile",              "Mobile Apps"),
    "social":   ("source",   4.00, "generic.organization",        "Social Networks"),
    "web":      ("source",   4.88, "generic.service",             "Web Servers"),
    "oltp":     ("source",   5.74, "generic.database",            "OLTP Databases"),

    "eventhub": ("ingest",   2.71, "azure.event-hubs",            "Azure Event Hub"),
    "adf":      ("ingest",   5.31, "azure.data-factories",        "Azure Data Factory"),

    "synapse":  ("storage",  3.91, "azure.synapse-analytics",     "Synapse Analytics"),
    "lake":     ("storage",  5.21, "azure.data-lake-store-gen1",  "Azure Data Lake"),

    "stream":     ("process", 1.55, "azure.stream-analytics-jobs", "Stream Analytics"),
    "cognitive":  ("process", 2.91, "azure.cognitive-services",    "Cognitive Services"),
    "databricks": ("process", 4.91, "azure.databricks",            "Azure Databricks"),

    "pbi_rt": ("serve", 1.51, "azure.power-bi-embedded",  "Power BI — Realtime"),
    "pbi_bi": ("serve", 2.61, "azure.power-bi-embedded",  "Power BI — Reporting"),
    "aas":    ("serve", 3.91, "azure.analysis-services",  "Analysis Services"),
    "cosmos": ("serve", 5.21, "azure.cosmos-db",          "Cosmos DB"),

    "analytics": ("business", 2.05, "generic.analytics", "Analytics"),
    "users":     ("business", 3.55, "generic.user",      "Business Users"),
    "apps":      ("business", 5.05, "generic.api",       "Applications"),
}

def rect_of(nid):
    zone, y, _, _ = NODES[nid]
    x = ZONE_X[zone] + (ZONE_W - CARD_W) / 2
    return (x, y, CARD_W, CARD_H)

def rmid(nid):
    x, y, w, h = rect_of(nid); return (x + w, y + h / 2)
def lmid(nid):
    x, y, w, h = rect_of(nid); return (x, y + h / 2)
def tmid(nid):
    x, y, w, h = rect_of(nid); return (x + w / 2, y)
def bmid(nid):
    x, y, w, h = rect_of(nid); return (x + w / 2, y + h)
def lpt(nid, frac):
    """A point on the left edge at vertical fraction `frac` (0=top,1=bottom)."""
    x, y, w, h = rect_of(nid); return (x, y + h * frac)

# ---------------------------------------------------------------- build
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Title
text(slide, 0.4, 0.18, 12.5, 0.6, [[("Azure Modern Data Platform", 24, True, INK)]],
     align=PP_ALIGN.LEFT)

# Tier zones + titles
TIERS = [
    ("source",   "SOURCE",       GEN_ZONE_FILL, GEN_ZONE_LINE, GRAY),
    ("ingest",   "INGEST",       AZ_ZONE_FILL,  AZ_ZONE_LINE,  AZ_BLUE),
    ("storage",  "STORAGE",      AZ_ZONE_FILL,  AZ_ZONE_LINE,  AZ_BLUE),
    ("process",  "PROCESS & AI", AZ_ZONE_FILL,  AZ_ZONE_LINE,  AZ_BLUE),
    ("serve",    "SERVE",        AZ_ZONE_FILL,  AZ_ZONE_LINE,  AZ_BLUE),
    ("business", "BUSINESS",     GEN_ZONE_FILL, GEN_ZONE_LINE, GRAY),
]
for zid, title, fill, line, tcolor in TIERS:
    x = ZONE_X[zid]
    rrect(slide, x, ZONE_TOP, ZONE_W, ZONE_H, fill, line, line_pt=1.25, radius=0.035)
    text(slide, x, ZONE_TOP + 0.12, ZONE_W, 0.28,
         [[(title, 10.5, True, tcolor)]], align=PP_ALIGN.CENTER)

# ---- connectors (drawn behind cards) ----
# Sources -> Event Hub (streaming): enter left edge at spread fractions.
for nid, frac in [("traffic", 0.28), ("iot", 0.42), ("mobile", 0.58), ("social", 0.72)]:
    sx, sy = rmid(nid); dx, dy = lpt("eventhub", frac)
    connector(slide, sx, sy, dx, dy)
# Sources -> Data Factory (batch)
for nid, frac in [("web", 0.38), ("oltp", 0.62)]:
    sx, sy = rmid(nid); dx, dy = lpt("adf", frac)
    connector(slide, sx, sy, dx, dy)

def link(a, b, color=FLOW, dashed=False):
    sx, sy = rmid(a); dx, dy = lmid(b)
    connector(slide, sx, sy, dx, dy, color=color, dashed=dashed)

link("eventhub", "stream")
link("adf", "lake")
link("lake", "synapse")
link("lake", "databricks")
link("synapse", "aas")
link("databricks", "cosmos")
link("stream", "pbi_rt")
# AI enrichment: databricks -> cognitive (upward, within Process tier)
sx, sy = tmid("databricks"); dx, dy = bmid("cognitive")
connector(slide, sx, sy, dx, dy, color=FLOW_AI, dashed=True, elbow=False)
# Analysis Services -> Power BI (Reporting) (upward within Serve tier)
sx, sy = tmid("aas"); dx, dy = bmid("pbi_bi")
connector(slide, sx, sy, dx, dy, elbow=False)
# Serve -> Business
link("pbi_rt", "analytics")
link("pbi_bi", "analytics")
link("cosmos", "apps")
# Business funnel -> Business Users (Users sits between Analytics and Apps)
sx, sy = bmid("analytics"); dx, dy = tmid("users")
connector(slide, sx, sy, dx, dy, elbow=False)
sx, sy = tmid("apps"); dx, dy = bmid("users")
connector(slide, sx, sy, dx, dy, elbow=False)

# ---- edge labels ----
def midlabel(a, b, txt, color=GRAY):
    ax, ay = rmid(a); bx, by = lmid(b)
    edge_label(slide, (ax + bx) / 2, (ay + by) / 2 - 0.13, txt, color)

midlabel("eventhub", "stream", "stream")
midlabel("adf", "lake", "batch")
midlabel("stream", "pbi_rt", "live")
# AI enrich label beside the vertical dashed connector, below the synapse->aas
# line so the two don't overlap.
edge_label(slide, tmid("databricks")[0] + 0.62, 4.55, "AI enrich", FLOW_AI)

# ---- cards (icon + label), drawn on top of connectors ----
for nid, (zone, y, key, lbl) in NODES.items():
    x = ZONE_X[zone] + (ZONE_W - CARD_W) / 2
    rrect(slide, x, y, CARD_W, CARD_H, CARD_FILL, CARD_LINE, line_pt=1.0, radius=0.06)
    icon(slide, key, x + (CARD_W - 0.42) / 2, y + 0.07, 0.42)
    text(slide, x + 0.04, y + 0.5, CARD_W - 0.08, 0.26,
         [[(lbl, 8.5, True, INK)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ---- footer legend ----
text(slide, 0.4, 6.92, 12.5, 0.3,
     [[("Solid ", 9, False, GRAY), ("= data flow    ", 9, False, GRAY),
       ("- - - ", 9, True, FLOW_AI), ("= AI enrichment (Databricks → Cognitive Services)", 9, False, GRAY)]],
     align=PP_ALIGN.LEFT)

out = sys.argv[1] if len(sys.argv) > 1 else "azure_data_platform.pptx"
prs.save(out)
print("saved", out)
