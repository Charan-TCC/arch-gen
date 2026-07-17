"""Bespoke high-fidelity build of the Video Analytics Architecture Diagram."""
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

# --- Palette (Light Mode) ---
BLUE_FILL = "E8F0FE"; BLUE_LINE = "1A73E8"
GREEN_FILL = "E6F4EA"; GREEN_LINE = "1E8E3E"
ORANGE_FILL = "FEF7E0"; ORANGE_LINE = "F9AB00"
SLATE_BG = "F8FAFC"; SLATE_LINE = "CBD5E1"
SLATE_INNER = "F1F5F9"
DARK_SLATE = "1E293B"
INK = "0F172A"; GRAY = "71717A"
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

def rect(slide, x, y, w, h, fill, line, line_pt=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
    _no_shadow(shp)
    return shp

def rrect(slide, x, y, w, h, fill, line, line_pt=1.0, dashed=False, radius=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try: shp.adjustments[0] = radius
    except Exception: pass
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None: shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
        if dashed:
            ln = shp.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
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

def straight_connector(slide, x1, y1, x2, y2, color=GRAY, pt=1.5, arrow_end=True, arrow_start=False, dashed=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if arrow_end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if dashed: ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    cn.shadow.inherit = False
    return cn

def label(slide, x, y, w, h, txt, color=GRAY):
    rect(slide, x, y, w, h, WHITE, None)
    text(slide, x, y, w, h, [[(txt, 9, False, color)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def draw_wide_node(slide, x, y, w, h, key, title, subtitle, fill, line):
    rrect(slide, x, y, w, h, fill, line, radius=0.04)
    if key:
        icon(slide, key, Inches(x + 0.15), Inches(y + h/2 - 0.25), Inches(0.5))
    tx = x + 0.7 if key else x + 0.1
    tw = w - 0.8 if key else w - 0.2
    text(slide, tx, y + h/2 - 0.22, tw, 0.44, 
         [[(title, 11, True, INK)], [(subtitle, 9, False, GRAY)]],
         align=PP_ALIGN.LEFT if key else PP_ALIGN.CENTER)

def draw_node(slide, cx, cy, key, label_txt):
    isz = 0.55
    icon(slide, key, Inches(cx - isz/2), Inches(cy - isz/2 - 0.08), Inches(isz))
    text(slide, cx - 0.8, cy + isz/2 - 0.05, 1.6, 0.2, [[(label_txt, 9, True, INK)]])

def zone_rrect(slide, x, y, w, h, title, fill=None, title_align=PP_ALIGN.LEFT):
    rrect(slide, x, y, w, h, fill, SLATE_LINE, 1.25, radius=0.03)
    text(slide, x + 0.1, y + 0.08, w - 0.2, 0.3, [[(title, 10, True, INK)]], align=title_align)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Video Analytics
text(slide, 0.5, 0.3, 12, 0.5, [[("AI-Powered Video Analytics Pipeline on Azure", 16, True, INK)]], align=PP_ALIGN.LEFT)

# 1. User
draw_node(slide, 1.0, 2.5, "generic.user", "User")

# 2. Upload / Ingest Zone
zone_rrect(slide, 2.0, 1.0, 2.5, 4.0, "Ingest & Storage", fill=SLATE_BG, title_align=PP_ALIGN.CENTER)
draw_node(slide, 3.25, 1.8, "azure.storage-azure-files", "Azure Files")
draw_node(slide, 3.25, 3.0, "azure.container-instances", "Video Ingest")
draw_node(slide, 3.25, 4.2, "azure.storage-accounts", "Azure Storage")

# 3. AI Processing Zone
zone_rrect(slide, 5.0, 1.0, 2.5, 4.0, "AI Processing", fill=SLATE_BG, title_align=PP_ALIGN.CENTER)
draw_node(slide, 6.25, 1.8, "azure.container-instances", "AI Vision")
draw_node(slide, 6.25, 3.0, "azure.cognitive-services", "AI Services")
draw_node(slide, 6.25, 4.2, "azure.container-instances", "Anomaly Detection")

# 4. Web App & Functions Zone
zone_rrect(slide, 8.0, 1.0, 4.0, 4.0, "Application & Events", fill=SLATE_BG, title_align=PP_ALIGN.CENTER)
draw_node(slide, 9.0, 1.8, "azure.sql-database", "Azure SQL DB")
draw_node(slide, 11.0, 1.8, "azure.app-services", "Web App")
draw_node(slide, 11.0, 3.0, "generic.identity", "Azure AD")
draw_node(slide, 9.0, 4.2, "azure.function-apps", "Azure Functions")
draw_node(slide, 11.0, 4.2, "generic.email", "SendGrid")

# 5. Management Plane
zone_rrect(slide, 2.0, 5.5, 10.0, 1.5, "Management & Orchestration", fill=SLATE_INNER, title_align=PP_ALIGN.CENTER)
draw_node(slide, 3.25, 6.25, "azure.kubernetes-services", "AKS")
draw_node(slide, 6.0, 6.25, "azure.container-registries", "ACR")
draw_node(slide, 8.5, 6.25, "azure.key-vaults", "Key Vault")
draw_node(slide, 10.5, 6.25, "azure.monitor", "Monitor")

# Edges
straight_connector(slide, 1.5, 2.5, 2.8, 1.8) # User -> Azure Files
straight_connector(slide, 3.25, 2.2, 3.25, 2.6) # Files -> Ingest
straight_connector(slide, 3.25, 3.4, 3.25, 3.8) # Ingest -> Storage

straight_connector(slide, 3.7, 4.2, 5.8, 1.8) # Storage -> AI Vision
straight_connector(slide, 6.25, 2.2, 6.25, 2.6, arrow_start=True) # Vision <-> AI Services
straight_connector(slide, 6.25, 3.4, 6.25, 3.8) # Vision -> Anomaly

straight_connector(slide, 6.7, 4.2, 8.5, 1.8) # Anomaly -> SQL
straight_connector(slide, 6.7, 4.2, 8.5, 4.2) # Anomaly -> Functions

straight_connector(slide, 9.5, 4.2, 10.5, 4.2) # Functions -> SendGrid
straight_connector(slide, 9.5, 1.8, 10.5, 1.8) # SQL -> Web App

straight_connector(slide, 11.0, 2.2, 11.0, 2.6, arrow_start=True) # Web App <-> Azure AD
straight_connector(slide, 1.2, 2.2, 10.5, 2.9, arrow_start=True) # User <-> Azure AD (long)
straight_connector(slide, 1.5, 2.4, 10.5, 1.7) # User -> Web App

straight_connector(slide, 5.5, 6.25, 3.7, 6.25) # ACR -> AKS
straight_connector(slide, 3.25, 5.8, 3.25, 3.4, arrow_end=True, dashed=True) # AKS -> Ingest
straight_connector(slide, 3.25, 5.8, 6.25, 4.6, arrow_end=True, dashed=True) # AKS -> Anomaly
straight_connector(slide, 3.25, 5.8, 6.25, 2.2, arrow_end=True, dashed=True) # AKS -> Vision

prs.save("video_analytics_python.pptx")
print("saved video_analytics_python.pptx")
