"""Bespoke high-fidelity build of the Shine Control Plane Architecture Diagram in Light Mode."""
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

def straight_connector(slide, x1, y1, x2, y2, color=GRAY, pt=1.5, arrow_end=True, arrow_start=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if arrow_end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
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

# Setup
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 1. Outer Control Plane
zone_rrect(slide, 2.5, 0.5, 8.333, 6.3, "Control plane • Shine on GCP", fill=SLATE_BG, title_align=PP_ALIGN.CENTER)

# 2. AI Foundation (AIF)
draw_wide_node(slide, 3.0, 1.0, 7.333, 0.8, "gcp.vertex-ai", "AI Foundation (AIF)", "governed AI inference endpoints", BLUE_FILL, BLUE_LINE)

# Text annotation between AIF and Cloud Run
text(slide, 3.0, 1.9, 7.333, 0.3, [[("app calls AIF via AI SDK / Agent SDK", 10, False, INK)]])

# 3. Cloud Run box
zone_rrect(slide, 3.0, 2.3, 7.333, 1.4, "Cloud Run", fill=SLATE_INNER)
draw_wide_node(slide, 3.2, 2.7, 3.3, 0.8, "azure.browser", "Frontend", "Web UI • SSO", GREEN_FILL, GREEN_LINE)
draw_wide_node(slide, 6.8, 2.7, 3.3, 0.8, "generic.api", "Backend", "API • orchestration", GREEN_FILL, GREEN_LINE)

# 4. Cloud Run Jobs
draw_wide_node(slide, 3.0, 3.9, 7.333, 0.8, "gcp.cloud-run", "Cloud Run Jobs — migration workers", "batch, long-running tasks", GREEN_FILL, GREEN_LINE)

# 5. Identity Band
zone_rrect(slide, 3.0, 4.9, 7.333, 1.55, "Identity & secrets on GCP", fill=SLATE_INNER)
draw_wide_node(slide, 3.2, 5.4, 2.0, 0.9, "azure.managed-identities", "Azure SPN", "auth to Azure", BLUE_FILL, BLUE_LINE)
draw_wide_node(slide, 5.6, 5.4, 2.0, 0.9, "gcp.secret-manager", "Secret Manager", "Shine secrets store", ORANGE_FILL, ORANGE_LINE)
draw_wide_node(slide, 8.0, 5.4, 2.0, 0.9, "gcp.identity-and-access-management", "GCP SA", "Cloud Run identity", GREEN_FILL, GREEN_LINE)

# 6. External Nodes
draw_node(slide, 1.4, 5.85, "azure.resource-groups", "Azure Source")
draw_node(slide, 11.9, 5.85, "gcp.cloud-generic", "GCP Target")

# 7. Edges
# azure_spn -> azure_source (points left)
straight_connector(slide, 3.2, 5.85, 1.8, 5.85)
label(slide, 2.1, 5.7, 1.0, 0.3, "read-only")

# secret_manager -> azure_spn (points left)
straight_connector(slide, 5.6, 5.85, 5.2, 5.85)
label(slide, 5.25, 5.7, 0.3, 0.3, "SPN secret")

# cloud_run_jobs -> secret_manager (points down)
straight_connector(slide, 6.6, 4.7, 6.6, 5.4)
label(slide, 6.65, 4.95, 0.6, 0.3, "reads")

# cloud_run_jobs -> gcp_sa (points down)
straight_connector(slide, 9.0, 4.7, 9.0, 5.4)
label(slide, 9.05, 4.95, 0.6, 0.3, "runs as")

# gcp_sa -> gcp_target (points right)
straight_connector(slide, 10.0, 5.85, 11.5, 5.85)
label(slide, 10.35, 5.7, 0.8, 0.3, "provision")

# 8. Annotations
# Left side box
rrect(slide, 0.5, 3.1, 2.2, 1.2, DARK_SLATE, None, radius=0.05)
text(slide, 0.6, 3.2, 2.0, 1.0, [[("Shine control plane with added Azure SPN, Secret Manager and GCP Service Account identity band", 10, False, WHITE)]], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

# Bottom text
text(slide, 1.0, 7.0, 11.333, 0.3, [[("Cloud Run runs as the GCP SA → reads Shine secrets (incl. the Azure SPN client secret) from Secret Manager → SPN authenticates to Azure", 11, False, GRAY)]])

prs.save("shine_control_plane_light.pptx")
print("saved shine_control_plane_light.pptx")
