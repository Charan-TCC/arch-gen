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
VPC_BG = "FBF6F0"; VPC_LINE = "D28D3F"
MAIN_SERVER_BG = "FBF6FB"; MAIN_SERVER_LINE = "D132D9"
COMPUTE_NODES_BG = "FCF5F5"; COMPUTE_NODES_LINE = "D64040"
GRAY = "71717A"; INK = "0F172A"; WHITE = "FFFFFF"
PURPLE = "9900FF"

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

def straight_connector(slide, x1, y1, x2, y2, color=INK, pt=1.5, arrow_end=True, arrow_start=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if arrow_end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

def elbow_connector(slide, x1, y1, x2, y2, color=INK, pt=1.5, arrow_end=True, arrow_start=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if arrow_end: ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_start: ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

def draw_node(slide, cx, cy, key, label_txt):
    isz = 0.8
    icon(slide, key, Inches(cx - isz/2), Inches(cy - isz/2), Inches(isz))
    text(slide, cx - 1.0, cy + isz/2 + 0.05, 2.0, 0.4, [[(label_txt, 11, False, "000000")]])

def zone_rect(slide, x, y, w, h, title, icon_key=None, fill=None, line=None, title_color=INK):
    rrect(slide, x, y, w, h, fill, line, 1.25, radius=0.03)
    if icon_key:
        icon(slide, icon_key, Inches(x + 0.1), Inches(y + 0.1), Inches(0.2))
        text(slide, x + 0.35, y + 0.08, w - 0.4, 0.3, [[(title, 10, False, title_color)]], align=PP_ALIGN.LEFT)
    else:
        text(slide, x + 0.1, y + 0.08, w - 0.2, 0.3, [[(title, 10, False, title_color)]], align=PP_ALIGN.LEFT)

# Setup
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 1. Containers
zone_rect(slide, 3.5, 0.5, 7.8, 6.5, "VPC SUBNET", icon_key="aws.virtual-private-cloud", fill=VPC_BG, line=VPC_LINE, title_color="402681")
zone_rect(slide, 3.9, 2.5, 3.0, 3.0, "MAIN SERVER", fill=MAIN_SERVER_BG, line=MAIN_SERVER_LINE, title_color="666666")
zone_rect(slide, 9.4, 1.0, 1.5, 5.6, "COMPUTE\nNODES", fill=COMPUTE_NODES_BG, line=COMPUTE_NODES_LINE, title_color="666666")

# 2. Nodes
draw_node(slide, 1.0, 3.3, "aws.api-gateway", "API gateway")
draw_node(slide, 2.5, 3.3, "aws.lambda", "Lambda")
draw_node(slide, 2.5, 5.2, "aws.s3", "S3")

draw_node(slide, 4.8, 3.5, "aws.ec2", "Server")
draw_node(slide, 6.2, 4.5, "aws.rds", "Data")

draw_node(slide, 8.2, 3.5, "aws.simple-queue-service", "Queue")

draw_node(slide, 10.15, 2.2, "aws.ec2", "Worker1")
draw_node(slide, 10.15, 3.7, "aws.ec2", "Worker2")
draw_node(slide, 10.15, 5.2, "aws.ec2", "Worker3")

draw_node(slide, 12.3, 3.7, "aws.redshift", "Analytics")

# 3. Edges
# API gateway to Lambda
straight_connector(slide, 1.4, 3.3, 2.1, 3.3)
# Lambda to Server
elbow_connector(slide, 2.9, 3.3, 4.4, 3.5)
# Server to Queue
straight_connector(slide, 5.2, 3.5, 7.8, 3.5)
# Server to Data
elbow_connector(slide, 4.8, 3.9, 5.8, 4.5)
# Data to S3
elbow_connector(slide, 5.8, 4.5, 2.9, 5.2)

# Queue to Workers
elbow_connector(slide, 8.6, 3.5, 9.75, 2.2)
elbow_connector(slide, 8.6, 3.5, 9.75, 3.7)
elbow_connector(slide, 8.6, 3.5, 9.75, 5.2)

# Workers to Analytics
straight_connector(slide, 10.9, 3.7, 11.9, 3.7)

prs.save("out/vpc_architecture.pptx")
print("saved out/vpc_architecture.pptx")
