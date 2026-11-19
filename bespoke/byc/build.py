"""Bespoke high-fidelity build of the BYC E-Commerce Architecture Diagram."""
from __future__ import annotations

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
BLUE = "1A73E8"      # Sync Flow (solid)
PURPLE = "7C3AED"    # Auth Flow (dashed)
ORANGE = "F97316"    # Async Flow (dotted)
GRAY = "71717A"      # Observability (dotted-dash/gray)
INK = "0F172A"        # Text
SLATE_LINE = "CBD5E1" # Container borders
SLATE_BG = "F8FAFC"   # Container fills
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

def rect(slide, x, y, w, h, fill, line, line_pt=1.0, dashed=False):
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

def rrect(slide, x, y, w, h, fill, line, line_pt=1.0, dashed=False, radius=0.08):
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

def straight_connector(slide, x1, y1, x2, y2, color=BLUE, dashed=False, dotted=False, pt=1.5, arrow=True):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color); cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    elif dotted:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "sysDot"}))
    if arrow:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

def routed_connector(slide, points, color=BLUE, dashed=False, dotted=False, pt=1.5, arrow=True, label=None):
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i+1]
        has_arrow = arrow and (i == len(points) - 2)
        straight_connector(slide, x1, y1, x2, y2, color=color, dashed=dashed, dotted=dotted, pt=pt, arrow=has_arrow)
    if label and len(points) >= 2:
        # place label in the middle of the longest segment
        mid_idx = len(points) // 2
        mx = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
        my = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
        # small white backing box
        pad_w, pad_h = 0.5, 0.22
        rect(slide, mx - pad_w/2, my - pad_h/2, pad_w, pad_h, WHITE, None)
        text(slide, mx - 0.4, my - 0.1, 0.8, 0.2, [[(label, 8, False, color)]])

# Nodes dictionary to store centers
nodes: dict[str, tuple[float, float]] = {}

def draw_node(slide, nid, cx, cy, key, label, label2=None, isz=0.55):
    nodes[nid] = (cx, cy)
    icon(slide, key, Inches(cx - isz/2), Inches(cy - isz/2 - 0.08), Inches(isz))
    text(slide, cx - 0.8, cy + isz/2 - 0.05, 1.6, 0.2, [[(label, 9, True, INK)]])
    if label2:
        text(slide, cx - 0.8, cy + isz/2 + 0.12, 1.6, 0.2, [[(label2, 8, False, GRAY)]])

def zone(slide, x, y, w, h, title, is_nested=False):
    fill_col = None if is_nested else SLATE_BG
    rrect(slide, x, y, w, h, fill_col, SLATE_LINE, 1.25, radius=0.04 if is_nested else 0.02)
    # top-left title
    text(slide, x + 0.1, y + 0.05, w - 0.2, 0.3, [[(title, 10, True, INK)]], align=PP_ALIGN.LEFT)

# Setup presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Title
text(slide, 0.4, 0.2, 12.5, 0.5, [[("BYC Multi-Tenant E-Commerce Platform", 18, True, INK)]], align=PP_ALIGN.LEFT)

# ---- Layout Zones ----
# 1. Client Layer (outside cloud)
zone(slide, 0.4, 0.9, 1.5, 6.2, "Client Layer")
draw_node(slide, "web", 1.15, 1.8, "generic.client", "Web SPA")
draw_node(slide, "mobile", 1.15, 3.4, "generic.mobile", "Mobile Apps")
draw_node(slide, "api_client", 1.15, 5.0, "generic.api", "Integrators")

# 2. Primary Region Box
zone(slide, 2.1, 0.9, 8.8, 6.2, "Primary Region (us-east-1)")

# Inner Zones in Primary Region
# Top-Left: Edge & Gateway
zone(slide, 2.3, 1.3, 4.4, 1.3, "Edge & Gateway", is_nested=True)
draw_node(slide, "cf", 2.8, 1.9, "aws.cloudfront", "CloudFront")
draw_node(slide, "waf", 3.8, 1.9, "aws.waf", "AWS WAF")
draw_node(slide, "kong", 4.8, 1.9, "tech.kong", "Kong Gateway")
draw_node(slide, "alb", 5.9, 1.9, "aws.elastic-load-balancing", "ALB")

# Top-Right: Data Layer
zone(slide, 6.9, 1.3, 3.8, 2.2, "Data Layer (AWS managed)", is_nested=True)
draw_node(slide, "rds", 7.4, 2.0, "aws.rds", "RDS (Postgres)")
draw_node(slide, "ddb", 8.8, 2.0, "aws.dynamodb", "DynamoDB")
draw_node(slide, "redis", 10.1, 2.0, "aws.elasticache", "ElastiCache")
draw_node(slide, "os", 7.4, 2.9, "aws.opensearch-service", "OpenSearch")
draw_node(slide, "s3", 8.8, 2.9, "aws.s3", "Amazon S3")
draw_node(slide, "redshift", 10.1, 2.9, "aws.redshift", "Redshift")

# Center: EKS Mesh
zone(slide, 2.3, 3.7, 4.4, 2.2, "Amazon EKS / Istio mesh", is_nested=True)
# EKS services in 4 columns, 2 rows
draw_node(slide, "svc_auth", 2.85, 4.3, "aws.cognito", "Auth Service")
draw_node(slide, "svc_user", 3.8, 4.3, "generic.user", "User Service")
draw_node(slide, "svc_catalog", 4.75, 4.3, "generic.database", "Catalog")
draw_node(slide, "svc_search", 5.7, 4.3, "generic.search", "Search Service")

draw_node(slide, "svc_cart", 2.85, 5.2, "generic.workflow", "Cart & Order")
draw_node(slide, "svc_pay", 3.8, 5.2, "generic.security", "Payment Service")
draw_node(slide, "svc_inv", 4.75, 5.2, "generic.container", "Inventory")
draw_node(slide, "svc_notif", 5.7, 5.2, "aws.simple-notification-service", "Notification")

# Bottom-Left EKS (Messaging/Worker block is also in EKS)
zone(slide, 6.9, 3.7, 3.8, 2.2, "Async Messaging & Workers (EKS/MSK)", is_nested=True)
draw_node(slide, "msk", 7.5, 4.3, "aws.managed-streaming-for-apache-kafka", "Amazon MSK")
draw_node(slide, "sqs", 7.5, 5.2, "aws.simple-queue-service", "Amazon SQS")
draw_node(slide, "worker", 9.4, 4.8, "generic.workflow", "Worker Pods")

# Bottom-Left: Observability
zone(slide, 2.3, 6.0, 4.4, 1.0, "Observability Stack", is_nested=True)
draw_node(slide, "otel", 2.8, 6.4, "generic.monitoring", "OpenTelemetry")
draw_node(slide, "prom", 3.8, 6.4, "tech.prometheus", "Prometheus")
draw_node(slide, "loki", 4.9, 6.4, "generic.monitoring", "Loki / ELK")
draw_node(slide, "grafana", 6.0, 6.4, "tech.grafana", "Grafana")

# Bottom-Right: Secondary Region
zone(slide, 6.9, 6.0, 3.8, 1.0, "Secondary Region (eu-west-1) [DR]", is_nested=True)
draw_node(slide, "dr_standby", 8.8, 6.4, "aws.rds", "DR Standby Replica")

# 3. Third-Party Services (outside cloud on the right)
zone(slide, 11.1, 0.9, 1.8, 6.2, "Third-Party Services")
draw_node(slide, "okta", 12.0, 1.8, "tech.okta", "Okta")
draw_node(slide, "stripe", 12.0, 3.0, "tech.stripe", "Stripe")
draw_node(slide, "sendgrid", 12.0, 4.2, "generic.service", "SendGrid")
draw_node(slide, "shipping", 12.0, 5.4, "generic.service", "Logistics API")

# ---- Connections with Precise Anchors ----
def get_anchor(nid, side):
    cx, cy = nodes[nid]
    w, h = 0.7, 0.65
    if side == "top": return cx, cy - h/2 - 0.08
    if side == "bottom": return cx, cy + h/2
    if side == "left": return cx - w/2, cy - 0.05
    if side == "right": return cx + w/2, cy - 0.05
    return cx, cy

# Auth (dashed purple)
# 1. web -> okta (routes over the top of the entire diagram)
web_r = get_anchor("web", "right")
okta_t = get_anchor("okta", "top")
routed_connector(slide, [web_r, (2.0, web_r[1]), (2.0, 0.6), (12.0, 0.6), okta_t], color=PURPLE, dashed=True)

# 2. kong -> okta
kong_b = get_anchor("kong", "bottom")
okta_l = get_anchor("okta", "left")
routed_connector(slide, [kong_b, (4.8, 2.75), (11.0, 2.75), (11.0, okta_l[1]), okta_l], color=PURPLE, dashed=True, label="JWT")

# 3. svc_auth -> okta (Auth microservice fetches config)
auth_l = get_anchor("svc_auth", "left")
routed_connector(slide, [auth_l, (2.2, auth_l[1]), (2.2, 0.6), (12.0, 0.6), okta_t], color=PURPLE, dashed=True)

# Sync (solid blue)
# web/mobile/api -> cf
cf_l = get_anchor("cf", "left")
straight_connector(slide, get_anchor("web", "right")[0], get_anchor("web", "right")[1], cf_l[0], cf_l[1], color=BLUE)
routed_connector(slide, [get_anchor("mobile", "right"), (1.8, get_anchor("mobile", "right")[1]), (1.8, cf_l[1]), cf_l], color=BLUE)
routed_connector(slide, [get_anchor("api_client", "right"), (1.8, get_anchor("api_client", "right")[1]), (1.8, cf_l[1]), cf_l], color=BLUE)

# cf -> waf -> kong -> alb
straight_connector(slide, get_anchor("cf", "right")[0], get_anchor("cf", "right")[1], get_anchor("waf", "left")[0], get_anchor("waf", "left")[1], color=BLUE)
straight_connector(slide, get_anchor("waf", "right")[0], get_anchor("waf", "right")[1], get_anchor("kong", "left")[0], get_anchor("kong", "left")[1], color=BLUE)
straight_connector(slide, get_anchor("kong", "right")[0], get_anchor("kong", "right")[1], get_anchor("alb", "left")[0], get_anchor("alb", "left")[1], color=BLUE)

# cf -> s3 (static hosting)
cf_b = get_anchor("cf", "bottom")
s3_l = get_anchor("s3", "left")
routed_connector(slide, [cf_b, (2.8, 3.4), (8.0, 3.4), (8.0, s3_l[1]), s3_l], color=BLUE, label="Static")

# alb -> microservices (routes in the y=3.2 gap)
alb_b = get_anchor("alb", "bottom")
# alb -> auth
routed_connector(slide, [alb_b, (5.9, 3.2), (2.85, 3.2), get_anchor("svc_auth", "top")], color=BLUE)
# alb -> cart (routes cleanly down the left side at x=2.2)
routed_connector(slide, [alb_b, (5.9, 3.2), (2.85, 3.2), (2.85, 3.5), (2.2, 3.5), (2.2, 5.2), get_anchor("svc_cart", "left")], color=BLUE)
# alb -> catalog
routed_connector(slide, [alb_b, (5.9, 3.2), (4.75, 3.2), get_anchor("svc_catalog", "top")], color=BLUE)
# alb -> search
routed_connector(slide, [alb_b, (5.9, 3.2), (5.7, 3.2), get_anchor("svc_search", "top")], color=BLUE)

# microservices -> data stores
# auth -> redis
routed_connector(slide, [get_anchor("svc_auth", "top"), (2.85, 3.5), (10.1, 3.5), get_anchor("redis", "bottom")], color=BLUE)
# user -> rds
routed_connector(slide, [get_anchor("svc_user", "right"), (4.15, 4.3), (4.15, 3.6), (7.4, 3.6), get_anchor("rds", "bottom")], color=BLUE)
# catalog -> ddb
routed_connector(slide, [get_anchor("svc_catalog", "top"), (4.75, 3.6), (8.8, 3.6), get_anchor("ddb", "bottom")], color=BLUE)
# search -> os
routed_connector(slide, [get_anchor("svc_search", "top"), (5.7, 3.6), (7.4, 3.6), get_anchor("os", "bottom")], color=BLUE)
# cart -> rds
routed_connector(slide, [get_anchor("svc_cart", "right"), (3.2, 5.2), (3.2, 3.6), (7.4, 3.6), get_anchor("rds", "bottom")], color=BLUE)
# pay -> rds & stripe
routed_connector(slide, [get_anchor("svc_pay", "right"), (4.15, 5.2), (4.15, 3.6), (7.4, 3.6), get_anchor("rds", "bottom")], color=BLUE)
routed_connector(slide, [get_anchor("svc_pay", "right"), (4.15, 5.2), (4.15, 3.0), get_anchor("stripe", "left")], color=BLUE)
# inv -> rds
routed_connector(slide, [get_anchor("svc_inv", "right"), (5.1, 5.2), (5.1, 3.6), (7.4, 3.6), get_anchor("rds", "bottom")], color=BLUE)

# Async Flow (dotted orange)
# cart -> msk
routed_connector(slide, [get_anchor("svc_cart", "bottom"), (2.85, 5.6), (7.0, 5.6), (7.0, 4.3), get_anchor("msk", "left")], color=ORANGE, dotted=True)
# inv -> msk (staggered to y=5.75)
routed_connector(slide, [get_anchor("svc_inv", "bottom"), (4.75, 5.75), (7.1, 5.75), (7.1, 4.3), get_anchor("msk", "left")], color=ORANGE, dotted=True)
# svc_notif -> msk (staggered to y=5.75)
routed_connector(slide, [get_anchor("svc_notif", "bottom"), (5.7, 5.75), (7.1, 5.75), (7.1, 4.3), get_anchor("msk", "left")], color=ORANGE, dotted=True)

# msk -> worker
straight_connector(slide, get_anchor("msk", "right")[0], get_anchor("msk", "right")[1], get_anchor("worker", "left")[0], get_anchor("worker", "left")[1], color=ORANGE, dotted=True)

# worker -> redshift (routed at x=10.7 to avoid the Stripe label)
worker_r = get_anchor("worker", "right")
redshift_r = get_anchor("redshift", "right")
routed_connector(slide, [worker_r, (10.7, worker_r[1]), (10.7, redshift_r[1]), redshift_r], color=ORANGE, dotted=True)

# worker -> sqs & sqs -> worker (DLQ loop)
worker_b = get_anchor("worker", "bottom")
sqs_r = get_anchor("sqs", "right")
sqs_b = get_anchor("sqs", "bottom")
routed_connector(slide, [worker_b, (9.4, 5.7), (8.2, 5.7), sqs_r], color=ORANGE, dotted=True, label="DLQ")
routed_connector(slide, [sqs_b, (7.5, 5.7), (9.4, 5.7), worker_b], color=ORANGE, dotted=True)

# worker -> svc_notif
straight_connector(slide, get_anchor("worker", "bottom")[0], get_anchor("worker", "bottom")[1], get_anchor("svc_notif", "top")[0], get_anchor("svc_notif", "top")[1], color=ORANGE, dotted=True)

# svc_notif -> sendgrid
straight_connector(slide, get_anchor("svc_notif", "right")[0], get_anchor("svc_notif", "right")[1], get_anchor("sendgrid", "left")[0], get_anchor("sendgrid", "left")[1], color=ORANGE, dotted=True)

# svc_cart -> shipping
cart_b = get_anchor("svc_cart", "bottom")
shipping_l = get_anchor("shipping", "left")
routed_connector(slide, [cart_b, (2.85, 5.8), (11.0, 5.8), (11.0, shipping_l[1]), shipping_l], color=BLUE)

# Observability Flow (gray dotted-dash)
# Straight vertical telemetry connection between Search and OpenTelemetry
search_b = get_anchor("svc_search", "bottom")
otel_t = get_anchor("otel", "top")
straight_connector(slide, search_b[0], search_b[1], otel_t[0], otel_t[1], color=GRAY, dotted=True)

# otel -> prom, loki, grafana
otel_r = get_anchor("otel", "right")
straight_connector(slide, otel_r[0], otel_r[1], get_anchor("prom", "left")[0], get_anchor("prom", "left")[1], color=GRAY, dotted=True)
straight_connector(slide, get_anchor("prom", "right")[0], get_anchor("prom", "right")[1], get_anchor("loki", "left")[0], get_anchor("loki", "left")[1], color=GRAY, dotted=True)
straight_connector(slide, get_anchor("loki", "right")[0], get_anchor("loki", "right")[1], get_anchor("grafana", "left")[0], get_anchor("grafana", "left")[1], color=GRAY, dotted=True)

# Secondary Region replication (routed at x=11.0 to avoid Redshift)
rds_r = get_anchor("rds", "right")
dr_t = get_anchor("dr_standby", "top")
routed_connector(slide, [rds_r, (8.1, rds_r[1]), (8.1, 3.6), (11.0, 3.6), (11.0, 5.9), (8.8, 5.9), dr_t], color=BLUE, dashed=True, label="Replication")

# ---- Legend (bottom left corner of canvas) ----
# We draw a clean floating box for the legends to complete the design
lx, ly, lw, lh = 0.4, 5.8, 1.5, 1.3
rrect(slide, lx, ly, lw, lh, SLATE_BG, SLATE_LINE, radius=0.02)
text(slide, lx + 0.05, ly + 0.05, lw - 0.1, 0.22, [[("Connections Legend", 9, True, INK)]], align=PP_ALIGN.LEFT)

def add_legend_row(y, label, color, dashed=False, dotted=False):
    # draw a short line
    x1 = lx + 0.1
    x2 = x1 + 0.4
    straight_connector(slide, x1, y, x2, y, color=color, dashed=dashed, dotted=dotted, arrow=True, pt=1.25)
    text(slide, x2 + 0.05, y - 0.1, lw - 0.6, 0.2, [[(label, 8, False, INK)]], align=PP_ALIGN.LEFT)

add_legend_row(ly + 0.35, "Sync Request", BLUE)
add_legend_row(ly + 0.60, "OIDC / Auth", PURPLE, dashed=True)
add_legend_row(ly + 0.85, "Async Event", ORANGE, dotted=True)
add_legend_row(ly + 1.10, "Telemetry", GRAY, dotted=True)

import os; os.makedirs("out", exist_ok=True)
from archgen.shared.geomcheck import fix_presentation
fix_presentation(prs)  # hug text boxes, snap near-miss alignment
prs.save("out/byc_architecture.pptx")
print("saved byc_architecture.pptx")
