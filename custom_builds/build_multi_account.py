"""Bespoke build of the AWS Multi-Account architecture."""
from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

from archgen.registry.resolver import IconResolver
from archgen.registry import emf

R = IconResolver()

# --- palette ----------------------------------------------------------------
RED = "DB1D3D"
GREEN = "137333"
INK = "202B38"
WHITE = "FFFFFF"

def C(h):
    return RGBColor.from_string(h)

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
    if st is not None:
        shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))

def rrect(slide, x, y, w, h, fill, line, line_pt=1.25, dashed=False, radius=0.0):
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
        if dashed:
            ln = shp.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    _no_shadow(shp)
    return shp

def rect(slide, x, y, w, h, fill, line, line_pt=1.25, dashed=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
        if dashed:
            ln = shp.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    _no_shadow(shp)
    return shp

def text(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, wrap=True, spacing=1.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(9525)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    return box

def connector(slide, x1, y1, x2, y2, color=GREEN, dashed=False, pt=1.5, arrow=True):
    same = (abs(x1 - x2) < 1e-6) or (abs(y1 - y2) < 1e-6)
    kind = MSO_CONNECTOR.ELBOW
    cn = slide.shapes.add_connector(kind, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if arrow:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

def straight_connector(slide, x1, y1, x2, y2, color=GREEN, dashed=False, pt=1.5, arrow=True):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if arrow:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn


def circle_number(slide, x, y, num):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x-0.12), Inches(y-0.12), Inches(0.24), Inches(0.24))
    shp.fill.solid(); shp.fill.fore_color.rgb = C(GREEN)
    shp.line.fill.background()
    _no_shadow(shp)
    tf = shp.text_frame
    tf.margin_bottom = tf.margin_top = tf.margin_left = tf.margin_right = 0
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(num)
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.name = "Segoe UI"
    r.font.color.rgb = C(WHITE)

def draw_node(slide, cx, cy, key, label, label2=None, isz=0.7):
    icon(slide, key, Inches(cx - isz/2), Inches(cy - isz/2), Inches(isz))
    text(slide, cx - 0.7, cy + isz/2 + 0.05, 1.4, 0.4, [[(label, 9, False, INK)]])
    if label2:
        text(slide, cx - 0.7, cy + isz/2 + 0.2, 1.4, 0.4, [[(label2, 8, False, INK)]])

# ============================================================================
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# ---- Zones -----------------------------------------------------------------
def zone(x, y, w, h, title, icon_key=None):
    rect(slide, x, y, w, h, None, RED, 1.5, dashed=False)
    # top-left label
    if icon_key:
        icon(slide, icon_key, Inches(x + 0.05), Inches(y + 0.05), Inches(0.3))
        text(slide, x + 0.35, y + 0.05, 3.0, 0.3, [[(title, 11, False, RED)]], align=PP_ALIGN.LEFT)
    else:
        text(slide, x + 0.05, y + 0.05, 3.0, 0.3, [[(title, 11, False, RED)]], align=PP_ALIGN.LEFT)

# Hub Account
zone(0.3, 0.8, 6.2, 6.5, "Hub account", "aws.virtual-private-cloud") # Just a standard AWS icon for Hub

# Mgmt Account
zone(6.8, 0.8, 6.2, 1.8, "AWS Organizations Management account", "generic.settings")

# Spoke Account N
zone(6.8, 2.8, 6.2, 4.5, "Spoke account N", "aws.virtual-private-cloud")
# Spoke Account 1
zone(7.0, 3.1, 5.8, 4.0, "Spoke account 1", "aws.virtual-private-cloud")

# Policies dashed box
rect(slide, 7.3, 3.6, 5.2, 3.3, None, INK, 1.0, dashed=True)
icon(slide, "aws.identity-and-access-management", Inches(7.4), Inches(3.7), Inches(0.4))
text(slide, 7.8, 3.75, 1.0, 0.3, [[("Policies", 10, False, INK)]], align=PP_ALIGN.LEFT)

# Client (outside)
draw_node(slide, 1.5, 0.4, "generic.client", "Client", isz=0.5)

# ---- Nodes -----------------------------------------------------------------
# Hub Account
# Row 1
cf_x, cf_y = 1.5, 2.0; draw_node(slide, cf_x, cf_y, "aws.cloudfront", "Amazon", "CloudFront")
waf_x, waf_y = 3.3, 2.0; draw_node(slide, waf_x, waf_y, "aws.waf", "AWS WAF")
# Row 2
s3_x, s3_y = 1.5, 4.0; draw_node(slide, s3_x, s3_y, "aws.s3", "Amazon S3 bucket", "(serves web interface)")
apigw_x, apigw_y = 3.3, 3.6; draw_node(slide, apigw_x, apigw_y, "aws.api-gateway", "Amazon API", "Gateway")
cognito_x, cognito_y = 5.3, 2.5; draw_node(slide, cognito_x, cognito_y, "aws.cognito", "Amazon Cognito")
# Row 3
lambda1_x, lambda1_y = 4.8, 4.6; draw_node(slide, lambda1_x, lambda1_y, "aws.lambda", "AWS Lambda")
# Row 4
eb_x, eb_y = 3.3, 5.5; draw_node(slide, eb_x, eb_y, "aws.eventbridge", "Amazon", "EventBridge")
ddb_x, ddb_y = 6.0, 5.2; draw_node(slide, ddb_x, ddb_y, "aws.dynamodb", "Amazon DynamoDB")
# Row 5
step_x, step_y = 5.0, 6.2; draw_node(slide, step_x, step_y, "aws.step-functions", "AWS Step Functions")
lambda2_x, lambda2_y = 6.2, 6.2; draw_node(slide, lambda2_x, lambda2_y, "aws.lambda", "AWS Lambda")

# Mgmt Account
org_x, org_y = 10.0, 1.7; draw_node(slide, org_x, org_y, "aws.organizations", "AWS Organizations")

# Policies
def policy_node(x, y, label):
    icon(slide, "generic.folder", Inches(x), Inches(y - 0.2), Inches(0.4))
    text(slide, x + 0.5, y - 0.1, 2.0, 0.3, [[(label, 10, False, INK)]], align=PP_ALIGN.LEFT)

py = 4.3
policy_node(8.0, py, "S3 Bucket Policies")
py += 0.6
policy_node(8.0, py, "IAM Policies")
py += 0.6
policy_node(8.0, py, "SNS Topic Policies")
py += 0.6
policy_node(8.0, py, "Service Nth")

# ---- Edges with numbers ----------------------------------------------------
def draw_edge(x1, y1, x2, y2, num, nx, ny, straight=False):
    if straight:
        straight_connector(slide, x1, y1, x2, y2)
    else:
        connector(slide, x1, y1, x2, y2)
    circle_number(slide, nx, ny, num)

# Client -> CloudFront (1)
draw_edge(1.5, 0.7, 1.5, 1.6, 1, 1.5, 1.15, straight=True)

# CloudFront -> S3 (2)
draw_edge(1.5, 2.5, 1.5, 3.6, 2, 1.5, 3.0, straight=True)

# Client -> Cognito (3)
# It's an elbow connector going right, then down
connector(slide, 1.8, 0.4, 5.3, 2.1)
circle_number(slide, 4.3, 0.4, 3)

# Client -> WAF (4)
# straight down elbow? Or just straight
connector(slide, 1.8, 0.4, 3.3, 1.6)
circle_number(slide, 2.8, 1.0, 4)

# WAF -> API GW (5)
draw_edge(3.3, 2.4, 3.3, 3.2, 5, 3.3, 2.8, straight=True)

# API GW -> Cognito (6)
connector(slide, 3.6, 3.5, 5.0, 2.7)
circle_number(slide, 4.3, 2.7, 6)

# API GW -> Lambda 1 (7)
draw_edge(3.6, 3.8, 4.5, 4.5, 7, 4.0, 4.2)

# Lambda 1 -> DynamoDB (8)
draw_edge(5.1, 4.7, 5.7, 5.1, 8, 5.4, 4.9)

# Lambda 1 -> Organizations (9)
# From Lambda 1 up and right to Orgs
connector(slide, 4.8, 4.2, 10.0, 2.0)
circle_number(slide, 6.0, 4.0, 9)

# API GW -> EventBridge (10)
draw_edge(3.3, 4.1, 3.3, 5.1, 10, 3.3, 4.6, straight=True)

# EventBridge -> Step Functions (11)
draw_edge(3.6, 5.6, 4.7, 6.1, 11, 4.1, 5.8)

# Step Functions -> Lambda 2 (12)
draw_edge(5.3, 6.2, 5.9, 6.2, 12, 5.6, 6.2, straight=True)

# Lambda 2 -> Policies (13)
# To the dashed box border
draw_edge(6.5, 6.2, 7.3, 5.2, 13, 6.9, 5.7)

# Lambda 2 -> DynamoDB (14)
# From Lambda 2 up to DynamoDB
draw_edge(6.2, 5.9, 6.1, 5.6, 14, 6.1, 5.75, straight=True)

prs.save("multi_account_architecture_bespoke.pptx")
print("saved multi_account_architecture_bespoke.pptx")
