"""Bespoke, image-faithful build of the AWS "Data Transfer Hub" reference
architecture. archgen's generic flow/network layouts can't express this
hand-designed composition (nested account containers, a dashed Authentication
sub-box, a bordered Step Functions container, numbered callouts ①-⑧ and
cross-zone annotation arrows), so we lay it out explicitly here while reusing
archgen's official icon registry + SVG->PNG converter. All AWS service icons
are non-recolorable brand marks, so they render in authentic full colour.
Output is native, fully-editable PowerPoint shapes/text/connectors."""
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

# --- palette (AWS reference architecture style) -----------------------------
INK    = "232F3E"   # AWS squid ink — text, arrows, box borders
GRAY   = "5B6470"
BORDER = "3B4756"   # thin container border
PINK   = "E7157B"   # AWS Step Functions container accent
DASH   = "8592A0"   # dashed sub-box border
WHITE  = "FFFFFF"

ISZ = 0.82          # nominal icon box (inches)


def C(h):
    return RGBColor.from_string(h)


def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))


def icon(slide, key, cx, cy, size=ISZ, force_hex=None):
    """Draw an icon centred on (cx, cy) in inches."""
    e = R.get(key)
    recolor = None
    if force_hex:
        recolor = force_hex
    elif e.recolorable:
        brand = (e.brand_hex or "").lstrip("#")
        recolor = f"#{brand}" if brand else None
    png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)
    asp = emf.svg_aspect(e.path)
    if asp >= 1:
        w = size; h = size / asp
    else:
        h = size; w = size * asp
    slide.shapes.add_picture(str(png), Inches(cx - w / 2), Inches(cy - h / 2),
                             width=Inches(w), height=Inches(h))


def box(slide, x, y, w, h, line, line_pt=1.0, dashed=False, fill=None, radius=None):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    shp.line.color.rgb = C(line); shp.line.width = Pt(line_pt)
    if dashed:
        ln = shp.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    _no_shadow(shp)
    return shp


def text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         wrap=True, spacing=1.0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(6350)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.name = "Segoe UI"; r.font.color.rgb = C(color)
    return tb


def label(slide, cx, top, lines, size=11.0, color=INK, w=1.9):
    """Centred caption under an icon. `lines` = list of strings."""
    runs = [[(ln, size, False, color)] for ln in lines]
    text(slide, cx - w / 2, top, w, 0.6, runs, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.TOP, spacing=1.02)


def seg(slide, x1, y1, x2, y2, arrow=False, pt=2.0, color=INK):
    """A single straight connector segment; arrowhead at the (x2,y2) end."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:round"), {}))
    if arrow:
        ln.append(ln.makeelement(qn("a:tailEnd"),
                                 {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn


def elbow(slide, pts, pt=2.0, color=INK):
    """Poly-line through pts=[(x,y),...]; arrowhead on the final segment."""
    for i in range(len(pts) - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        seg(slide, x1, y1, x2, y2, arrow=(i == len(pts) - 2), pt=pt, color=color)


def badge(slide, cx, cy, num, d=0.34):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2),
                               Inches(d), Inches(d))
    o.fill.solid(); o.fill.fore_color.rgb = C(INK); o.line.fill.background()
    _no_shadow(o)
    text(slide, cx - d / 2, cy - d / 2, d, d, [[(str(num), 12.5, True, WHITE)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def cloud_glyph(slide, x, y):
    c = slide.shapes.add_shape(MSO_SHAPE.CLOUD, Inches(x), Inches(y), Inches(0.42), Inches(0.30))
    c.fill.background(); c.line.color.rgb = C(INK); c.line.width = Pt(1.6)
    _no_shadow(c)


# ============================================================================
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# ---- containers ------------------------------------------------------------
box(slide, 0.20, 0.25, 13.05, 6.95, INK, 1.25)                 # AWS Cloud (outer)
box(slide, 0.42, 1.02, 8.45, 5.95, BORDER, 1.0)                # Customer's AWS Account
box(slide, 9.30, 1.02, 3.75, 5.95, BORDER, 1.0)               # AWS Managed Account

# outer "AWS Cloud" header — AWS logo badge + label
logo = box(slide, 0.36, 0.40, 0.44, 0.44, INK, 0.5, fill=INK, radius=0.12)
text(slide, 0.36, 0.40, 0.44, 0.44, [[("aws", 10, True, WHITE)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(slide, 0.95, 0.44, 4.0, 0.38, [[("AWS Cloud", 14, False, INK)]], anchor=MSO_ANCHOR.MIDDLE)

cloud_glyph(slide, 0.55, 1.12)
text(slide, 1.08, 1.13, 5.0, 0.32, [[("Customer’s AWS Account", 13, False, INK)]],
     anchor=MSO_ANCHOR.MIDDLE)
cloud_glyph(slide, 9.44, 1.12)
text(slide, 9.98, 1.13, 3.2, 0.32, [[("AWS Managed Account", 13, False, INK)]],
     anchor=MSO_ANCHOR.MIDDLE)

# ---- Authentication dashed sub-box -----------------------------------------
box(slide, 1.15, 1.55, 2.40, 1.62, DASH, 1.0, dashed=True)
text(slide, 1.15, 1.60, 2.40, 0.28, [[("Authentication", 11, False, INK)]],
     align=PP_ALIGN.CENTER)
COG = (1.85, 2.30); OID = (2.95, 2.30)
icon(slide, "aws.cognito", *COG)
icon(slide, "generic.identity", *OID)
label(slide, COG[0], COG[1] + 0.44, ["Amazon Cognito"], size=10)
label(slide, OID[0], OID[1] + 0.44, ["OpenID", "Connect"], size=10)

# ---- Customer account service icons ----------------------------------------
DDB   = (5.35, 1.95)
APPS  = (3.70, 3.62)
LAMB  = (5.35, 3.62)
UI    = (1.05, 5.00)
USER  = (1.02, 6.10)
CF    = (3.70, 5.62)
S3C   = (5.35, 5.62)
FARG  = (7.45, 5.62)

icon(slide, "aws.dynamodb", *DDB)
label(slide, DDB[0], DDB[1] + 0.44, ["Amazon DynamoDB"], size=10.5)

icon(slide, "aws.appsync", *APPS)
label(slide, APPS[0], APPS[1] + 0.44, ["AWS AppSync"], size=10.5)

icon(slide, "aws.lambda", *LAMB)
label(slide, LAMB[0], LAMB[1] + 0.44, ["AWS Lambda"], size=10.5)

# Data Transfer Hub UI — monitor glyph
mon = box(slide, UI[0] - 0.42, UI[1] - 0.34, 0.84, 0.60, INK, 2.0, radius=0.06)
base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(UI[0] - 0.10), Inches(UI[1] + 0.26),
                              Inches(0.20), Inches(0.12))
base.fill.solid(); base.fill.fore_color.rgb = C(INK); base.line.fill.background(); _no_shadow(base)
foot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(UI[0] - 0.24), Inches(UI[1] + 0.38),
                              Inches(0.48), Inches(0.06))
foot.fill.solid(); foot.fill.fore_color.rgb = C(INK); foot.line.fill.background(); _no_shadow(foot)
label(slide, UI[0], UI[1] + 0.52, ["Data Transfer", "Hub UI"], size=10.5)

USER = (1.02, 6.25)
icon(slide, "generic.user", *USER, size=0.62)
label(slide, USER[0], USER[1] + 0.36, ["User"], size=10.5)

icon(slide, "aws.cloudfront", *CF)
label(slide, CF[0], CF[1] + 0.44, ["Amazon CloudFront"], size=10.5)
icon(slide, "aws.s3", *S3C)
label(slide, S3C[0], S3C[1] + 0.44, ["Amazon S3"], size=10.5)
icon(slide, "aws.fargate", *FARG)
label(slide, FARG[0], FARG[1] + 0.44, ["AWS Fargate"], size=10.5)

# ---- Step Functions workflow container -------------------------------------
SF_X, SF_Y, SF_W, SF_H = 6.30, 1.42, 2.30, 3.05
box(slide, SF_X, SF_Y, SF_W, SF_H, PINK, 2.0)
icon(slide, "aws.step-functions", SF_X + 0.28, SF_Y + 0.28, size=0.32)
text(slide, SF_X + 0.50, SF_Y + 0.10, SF_W - 0.55, 0.55,
     [[("AWS Step Functions", 10, False, INK)], [("workflow", 10, False, INK)]], spacing=1.0)
SFL = (SF_X + SF_W / 2, SF_Y + 1.05)   # Lambda inside
SFC = (SF_X + SF_W / 2, SF_Y + 2.20)   # CloudFormation inside
icon(slide, "aws.lambda", *SFL)
label(slide, SFL[0], SFL[1] + 0.44, ["AWS Lambda"], size=9.5)
icon(slide, "aws.cloudformation", *SFC)
label(slide, SFC[0], SFC[1] + 0.44, ["AWS CloudFormation"], size=9.5, w=2.1)

# ---- Managed account grid --------------------------------------------------
M_S3T  = (10.15, 2.30)
M_DDBT = (11.75, 2.30)
M_S3   = (10.15, 3.95)
M_ECRT = (11.75, 3.95)
M_ECR  = (10.15, 5.62)
M_DOCK = (11.75, 5.62)

ORANGE = "#ED7100"
icon(slide, "generic.document", *M_S3T, force_hex=ORANGE)
label(slide, M_S3T[0], M_S3T[1] + 0.44, ["S3 replication", "component", "template"], size=9.5, w=1.6)
icon(slide, "generic.document", *M_DDBT, force_hex=ORANGE)
label(slide, M_DDBT[0], M_DDBT[1] + 0.44, ["DynamoDB", "replication", "component template"], size=9.5, w=1.7)
icon(slide, "aws.s3", *M_S3)
label(slide, M_S3[0], M_S3[1] + 0.44, ["Amazon S3"], size=10)
icon(slide, "generic.document", *M_ECRT, force_hex=ORANGE)
label(slide, M_ECRT[0], M_ECRT[1] + 0.44, ["ECR replication", "component", "template"], size=9.5, w=1.6)
icon(slide, "aws.elastic-container-registry", *M_ECR)
label(slide, M_ECR[0], M_ECR[1] + 0.44, ["Amazon ECR"], size=10)
icon(slide, "generic.container", *M_DOCK)
label(slide, M_DOCK[0], M_DOCK[1] + 0.44, ["ECR replication", "Docker image"], size=9.5, w=1.7)

# ---- flow arrows + numbered callouts ---------------------------------------
h = ISZ / 2

# ③ UI -> Authentication
elbow(slide, [(0.85, UI[1] - h), (0.85, 2.30), (1.15, 2.30)])
badge(slide, 1.12, 4.35, 3)

# ② UI -> AppSync
elbow(slide, [(1.30, UI[1] - h), (1.30, APPS[1]), (APPS[0] - h, APPS[1])])
badge(slide, 2.25, 3.42, 2)

# ① UI -> CloudFront
elbow(slide, [(UI[0], UI[1] + h), (UI[0], CF[1]), (CF[0] - h, CF[1])])
badge(slide, 2.30, CF[1], 1)

# CloudFront -> S3
seg(slide, CF[0] + h, CF[1], S3C[0] - h, S3C[1], arrow=True)

# ④ AppSync -> Lambda
seg(slide, APPS[0] + h, APPS[1], LAMB[0] - h, LAMB[1], arrow=True)
badge(slide, 4.53, 3.42, 4)

# ⑤ Lambda -> Step Functions box
seg(slide, LAMB[0] + h, LAMB[1], SF_X, LAMB[1], arrow=True)
badge(slide, 6.02, 3.42, 5)

# ⑧ AppSync -> DynamoDB
elbow(slide, [(APPS[0], APPS[1] - h), (APPS[0], DDB[1]), (DDB[0] - h, DDB[1])])
badge(slide, 4.53, 1.76, 8)

# Step Functions box -> Fargate (down)
seg(slide, SFL[0], SF_Y + SF_H, FARG[0], FARG[1] - h, arrow=True)

# ⑥ Amazon S3 (managed) -> Step Functions box (CloudFormation)
seg(slide, M_S3[0] - h, SFC[1], SF_X + SF_W, SFC[1], arrow=True)
badge(slide, 9.15, SFC[1], 6)

# ⑦ Amazon ECR (managed) -> Fargate
seg(slide, M_ECR[0] - h, FARG[1], FARG[0] + h, FARG[1], arrow=True)
badge(slide, 9.15, FARG[1], 7)

import os; os.makedirs("out", exist_ok=True)
from archgen.shared.geomcheck import fix_presentation
fix_presentation(prs)  # hug text boxes, snap near-miss alignment
prs.save("out/data_transfer_hub.pptx")
print("saved data_transfer_hub.pptx")
