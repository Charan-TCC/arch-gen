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
REGION_BG = "F3F6F9"
REGION_LINE = "4078C0"
GLOBAL_BG = "F9F9F9"
GLOBAL_LINE = "666666"
INK = "0F172A"
WHITE = "FFFFFF"

def C(h): return RGBColor.from_string(h)

def icon(slide, key, left, top, size):
    e = R.get(key)
    png = emf.to_picture(e.path, prefer_vector=False)
    asp = emf.svg_aspect(e.path)
    if asp >= 1:
        w = size; h = int(size / asp)
    else:
        h = size; w = int(size * asp)
    slide.shapes.add_picture(str(png), left + (size - w) / 2, top + (size - h) / 2, width=w, height=h)

def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None: shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))

def rrect(slide, x, y, w, h, fill, line, line_pt=1.0, radius=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try: shp.adjustments[0] = radius
    except Exception: pass
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None: shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
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
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    return box

def draw_node(slide, cx, cy, key, label_txt):
    isz = 0.6
    icon(slide, key, Inches(cx - isz/2), Inches(cy - isz/2), Inches(isz))
    text(slide, cx - 1.0, cy + isz/2 + 0.05, 2.0, 0.4, [[(label_txt, 10, False, INK)]])

def zone_rect(slide, x, y, w, h, title, fill=None, line=None, title_color=INK):
    rrect(slide, x, y, w, h, fill, line, 1.25, radius=0.03)
    text(slide, x + 0.1, y + 0.08, w - 0.2, 0.3, [[(title, 11, True, title_color)]], align=PP_ALIGN.LEFT)

def elbow_connector(slide, x1, y1, x2, y2, color=INK, pt=1.5, arrow_end=True, arrow_start=False, dash=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    if dash:
        ln = cn.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    ln = cn.line._get_or_add_ln()
    if arrow_end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 1. Containers
zone_rect(slide, 0.5, 0.5, 12.33, 2.0, "Global / Edge Services", fill=GLOBAL_BG, line=GLOBAL_LINE)
zone_rect(slide, 0.5, 2.8, 6.0, 4.2, "Active Region", fill=REGION_BG, line=REGION_LINE)
zone_rect(slide, 6.83, 2.8, 6.0, 4.2, "Standby Region", fill=REGION_BG, line=REGION_LINE)

# 2. Nodes: Global
draw_node(slide, 1.5, 1.5, "generic.user", "Users")
draw_node(slide, 3.5, 1.5, "azure.dns-zones", "Azure DNS")
draw_node(slide, 5.5, 1.5, "azure.front-door-and-cdn-profiles", "Azure Front Door")
draw_node(slide, 7.5, 1.5, "azure.ad-b2c", "Azure AD")
draw_node(slide, 9.5, 1.5, "azure.cdn-profiles", "Azure CDN")
draw_node(slide, 11.5, 1.5, "azure.storage-accounts", "Blob Storage")

# 3. Nodes: Active Region
draw_node(slide, 1.5, 4.0, "azure.app-services", "Web App")
draw_node(slide, 3.5, 4.0, "azure.storage-queue", "Azure Queue")
draw_node(slide, 5.5, 4.0, "azure.function-apps", "Function App")
draw_node(slide, 1.5, 6.0, "azure.cache-redis", "Redis Cache")
draw_node(slide, 2.8, 6.0, "azure.sql-database", "SQL Database")
draw_node(slide, 4.1, 6.0, "azure.cosmos-db", "Cosmos DB")
draw_node(slide, 5.5, 6.0, "azure.cognitive-search", "Azure Search")

# 4. Nodes: Standby Region
draw_node(slide, 7.83, 4.0, "azure.app-services", "Web App")
draw_node(slide, 9.83, 4.0, "azure.storage-queue", "Azure Queue")
draw_node(slide, 11.83, 4.0, "azure.function-apps", "Function App")
draw_node(slide, 7.83, 6.0, "azure.cache-redis", "Redis Cache")
draw_node(slide, 9.16, 6.0, "azure.sql-database", "SQL Database")
draw_node(slide, 10.49, 6.0, "azure.cosmos-db", "Cosmos DB")
draw_node(slide, 11.83, 6.0, "azure.cognitive-search", "Azure Search")

# 5. Edges
# Global Edges
elbow_connector(slide, 1.8, 1.5, 3.2, 1.5) # User -> DNS
elbow_connector(slide, 3.8, 1.5, 5.2, 1.5) # DNS -> FD
elbow_connector(slide, 1.5, 1.8, 5.5, 1.2, dash=True) # User -> FD
elbow_connector(slide, 1.5, 1.2, 7.5, 1.2, dash=True) # User -> AD
elbow_connector(slide, 1.5, 1.2, 9.5, 1.5) # User -> CDN
elbow_connector(slide, 9.8, 1.5, 11.2, 1.5) # CDN -> Blob

# Global -> Regions
elbow_connector(slide, 5.5, 1.8, 1.5, 3.7) # FD -> Active Web
elbow_connector(slide, 7.5, 1.8, 1.5, 3.7, dash=True) # AD -> Active Web
elbow_connector(slide, 5.5, 1.8, 7.83, 3.7, dash=True) # FD -> Standby Web
elbow_connector(slide, 7.5, 1.8, 7.83, 3.7, dash=True) # AD -> Standby Web

# Active Internal
elbow_connector(slide, 1.8, 4.0, 3.2, 4.0) # Web -> Queue
elbow_connector(slide, 3.8, 4.0, 5.2, 4.0) # Queue -> Func
elbow_connector(slide, 5.5, 4.3, 1.5, 5.7) # Func -> Redis
elbow_connector(slide, 5.5, 4.3, 2.8, 5.7) # Func -> SQL
elbow_connector(slide, 5.5, 4.3, 4.1, 5.7) # Func -> Cosmos
elbow_connector(slide, 5.5, 4.3, 5.5, 5.7) # Func -> Search
elbow_connector(slide, 1.5, 4.3, 1.5, 5.7) # Web -> Redis
elbow_connector(slide, 1.5, 4.3, 5.5, 5.7) # Web -> Search

# Standby Internal
elbow_connector(slide, 8.13, 4.0, 9.53, 4.0) # Web -> Queue
elbow_connector(slide, 10.13, 4.0, 11.53, 4.0) # Queue -> Func
elbow_connector(slide, 11.83, 4.3, 7.83, 5.7) # Func -> Redis
elbow_connector(slide, 11.83, 4.3, 9.16, 5.7) # Func -> SQL
elbow_connector(slide, 11.83, 4.3, 10.49, 5.7) # Func -> Cosmos
elbow_connector(slide, 11.83, 4.3, 11.83, 5.7) # Func -> Search
elbow_connector(slide, 7.83, 4.3, 7.83, 5.7) # Web -> Redis
elbow_connector(slide, 7.83, 4.3, 11.83, 5.7) # Web -> Search

# Geo-Replication
elbow_connector(slide, 3.1, 6.0, 8.86, 6.0, dash=True) # SQL1 -> SQL2
elbow_connector(slide, 4.4, 6.0, 10.19, 6.0, dash=True) # Cosmos1 -> Cosmos2

import os; os.makedirs("out", exist_ok=True)
from archgen.shared.geomcheck import fix_presentation
fix_presentation(prs)  # hug text boxes, snap near-miss alignment
prs.save("out/multi_region_azure_bespoke.pptx")
print("Saved out/multi_region_azure_bespoke.pptx")
