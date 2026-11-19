"""App Service architecture (mobile app + App Service backend) — bespoke 16:9 build.

Reference: colour-coded diagram — Entra ID / Mobile app on the left,
App Service (FastAPI, JWT auth, Push trigger, MQTT listener) in the middle,
Data sources / Key Vault / Application Insights on the right, Notification Hubs
under the backend, GitHub CI/CD band at the bottom. Bespoke because the
composition (zones + labelled bidirectional arrows + full band) has no archgen
profile (CLAUDE.md §3).

    .venv/bin/python bespoke/app_service_iot/build.py [out/app_service_iot.pptx]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Inches, Pt, Emu

from archgen.shared import compose as c
from archgen.kit import RegistryIcons, style

st, icons = style("corporate-clean"), RegistryIcons()
OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "out/app_service_iot.pptx")

# Category colours (legend at the bottom)
PEACH, GREEN, BLUE, PURPLE, GREY = "FDE4CC", "D7F2D4", "D8EAFB", "E7E1FB", "ECECEC"
# card border = deeper shade of each light fill (keeps cards distinct on tinted sections)
EDGE = {PEACH: "F2B27A", GREEN: "8FD18A", BLUE: "86B8EC", PURPLE: "B3A6F2", GREY: "BDBDBD"}
INK, MUTED, LINE = "1A1A1A", "4B5563", "333333"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])
c.set_background(s, "FFFFFF")
c.text(s, 0.5, 0.2, 12.3, None, "Backend Service architecture",
       font=st.font_heading, size=24, bold=True, color=INK)
c.text(s, 0.5, 0.68, 12.3, None, "A mobile app talks to one App Service backend.",
       font=st.font_body, size=13, color=MUTED)


def card(key, label, sub, x, y, w, h, fill, icon_in=0.3):
    """Filled card: icon + bold label + sublabel as ONE group (CLAUDE.md §7)."""
    grp = s.shapes.add_group_shape()
    sh = grp.shapes
    bx = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    bx.adjustments[0] = 0.06
    bx.fill.solid()
    bx.fill.fore_color.rgb = RGBColor.from_string(fill)
    bx.line.color.rgb = RGBColor.from_string(EDGE[fill])
    bx.line.width = Pt(1)
    c.kill_effects(bx)
    lines = 1 + sub.count("\n") + (1 if sub else 0)
    block = icon_in + 0.05 + lines * 0.17
    top = y + max(0.08, (h - block) / 2)
    png = icons.png(key, icon_in)
    if png:
        pic = sh.add_picture(str(png), Inches(x + (w - icon_in) / 2), Inches(top),
                             Inches(icon_in), Inches(icon_in))
        c.kill_effects(pic)
    ty = top + icon_in + 0.05
    tb = sh.add_textbox(Inches(x + 0.06), Inches(ty), Inches(w - 0.12),
                        Inches(min(lines * 0.17 + 0.02, y + h - ty - 0.03)))
    c.kill_effects(tb)
    tf = tb.text_frame
    tf.word_wrap, tf.auto_size = True, MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    for i, (t, sz, b) in enumerate([(label, 10, True)] + [(l, 9, False) for l in sub.split("\n") if sub]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text, r.font.size, r.font.bold, r.font.name = t, Pt(sz), b, st.font_body
        r.font.color.rgb = RGBColor.from_string(INK)
    return grp


def zone(x, y, w, h, title, color=MUTED, line="E2E2E2", fill="FAFAFA", icon=None):
    """Section container with an icon + title heading above it."""
    tx = x
    if icon and (png := icons.png(icon, 0.22)):
        c.kill_effects(s.shapes.add_picture(str(png), Inches(x), Inches(y - 0.28),
                                            Inches(0.22), Inches(0.22)))
        tx = x + 0.28
    c.text(s, tx, y - 0.285, min(w, 2.4), None, title, font=st.font_body, size=11,
           bold=True, color=color)
    return c.rrect(s, x, y, w, h, fill=fill, line=line, line_pt=1, radius=0.03)


def arrow(x1, y1, x2, y2, both=False):
    return c.connector(s, x1, y1, x2, y2, color=LINE, pt=1.25, tail=both)


def label(t, cx, cy, w=1.2):
    """Edge label centred on (cx, cy)."""
    h = c.text_height(t, w, 8.5) - 0.02
    return c.text(s, cx - w / 2, cy - h / 2, w, h, t, font=st.font_body, size=8.5,
                  color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---- geometry (inches) — compact, with margin above/below --------------------
LX, LW = 0.5, 1.7                       # left column
ZX, ZY, ZW, ZH = 2.95, 2.3, 4.3, 2.05  # App Service zone
C1, C2, CW, CH = 3.15, 5.3, 1.75, 0.82  # zone columns
R1, R2 = 2.45, 3.4                     # zone rows
RX = 7.75                               # right column
NY, NH = 4.62, 0.82                     # notification row
ZB = ZY + ZH                            # zone bottom (4.30)

# ---- left column ----
card("azure.entra-id-protection", "Microsoft Entra ID", "SSO login + SPN", LX, 1.15, LW, 0.82, PURPLE)
card("generic.mobile", "Mobile app", "SSO login → JWT\nlive view + alerts", LX, ZY, LW, ZH, PEACH, icon_in=0.42)

# ---- App Service ----
zone(ZX, ZY, ZW, ZH, "App Service", color="2E7D32", line="BFE5BB", fill="F3FBF2",
     icon="azure.app-services")
card("tech.python", "FastAPI", "REST + SSE", C1, R1, CW, CH, GREEN)
card("tech.jwt", "JWT auth", "validate token", C2, R1, CW, CH, GREEN)
card("generic.function", "Push trigger", "notify users", C1, R2, CW, CH, GREEN)
card("generic.monitoring", "MQTT listener", "check thresholds\npersist live data", C2, R2, CW, CH, GREEN, icon_in=0.24)
card("azure.notification-hubs", "Notification Hubs", "handles APNs + FCM", C1, NY, CW, NH, BLUE)

# ---- right column ----
DY, DH = ZY, 1.3
zone(RX, DY, 5.0, DH, "Data sources", color="2F6FC0", line="B5D3F3", fill="F2F8FE",
     icon="generic.database")
for x, key, lab, sub in [(7.95, "azure.api-management-services", "Ignition API", "history + user access"),
                         (9.55, "generic.queue", "MQTT broker", "live machine data"),
                         (11.15, "azure.sql-server", "MS SQL", "device ID + details\nlast MQTT data")]:
    card(key, lab, sub, x, R1, 1.4, 0.98, BLUE)
card("azure.key-vaults", "Key Vault", "secrets", RX, 3.72, 2.4, 0.82, PURPLE)
card("azure.application-insights", "Application Insights", "logs + metrics", RX, NY, 2.4, NH, BLUE)

# ---- CI/CD band ----
BY, BH, BCY = 5.85, 1.05, 5.965
zone(LX, BY, 5.05 - LX, BH, "GitHub CI/CD", line="DADADA", fill="F7F7F7",
     icon="tech.github")
card("tech.github", "GitHub Repos", "source + workflow", 0.7, BCY, 1.7, 0.82, GREY)
card("generic.workflow", "GitHub Actions", "build → test → deploy", 2.85, BCY, 2.0, 0.82, GREY)

# ---- arrows + labels ----
mid = LX + LW / 2
# Entra ID <-> App Service (token validation): L-route right, then down into the zone top
EX = C2 + CW / 2   # lands above JWT auth, clear of the "App Service" heading
c.connector(s, LX + LW, 1.56, EX, 1.56, color=LINE, pt=1.25, head=False, tail=True)
arrow(EX, 1.56, EX, ZY);                         label("validate token", (LX + LW + EX) / 2, 1.43, w=0.9)
arrow(mid, 1.97, mid, ZY, both=True);            label("SSO login · JWT", mid + 0.75, 2.15)
arrow(LX + LW, 3.2, ZX, 3.2, both=True);         label("HTTPS · SSE\n· JWT", (LX + LW + ZX) / 2, 3.5, w=0.75)
arrow(C1 + CW, R1 + CH / 2, C2, R1 + CH / 2, both=True)
arrow(C2, R2 + CH / 2, C1 + CW, R2 + CH / 2)      # MQTT listener -> Push trigger
p = C1 + CW / 2
arrow(p, R2 + CH, p, NY);                        label("notify", p + 0.4, (ZB + NY) / 2, w=0.6)
# Notification Hubs -> Mobile app: L-route (left, then up into the app's bottom)
c.connector(s, C1, NY + NH / 2, mid, NY + NH / 2, color=LINE, pt=1.25, head=False)
arrow(mid, NY + NH / 2, mid, ZB);                label("push via APNs / FCM", (mid + C1) / 2 + 0.1, NY + NH / 2 - 0.13, w=1.3)
q = C2 + CW / 2
# GitHub Actions deploys straight to App Service: L-route right, then up
c.connector(s, 4.85, BCY + 0.41, q, BCY + 0.41, color=LINE, pt=1.25, head=False)
arrow(q, BCY + 0.41, q, ZB);                     label("deploy (OIDC)", q + 0.5, (ZB + BY) / 2 + 0.2, w=0.85)
arrow(2.4, BCY + 0.41, 2.85, BCY + 0.41);      label("on push", 2.625, BCY + 0.6, w=0.45)
ZR = ZX + ZW
arrow(ZR, DY + DH / 2, RX, DY + DH / 2, both=True)   # <-> Data sources (straight)
arrow(ZR, 4.03, RX, 4.03, both=True)                   # <-> Key Vault (straight)
# -> Application Insights: clean L from the zone's bottom edge (no stray elbow)
AX = 7.0
c.connector(s, AX, ZB, AX, NY + NH / 2, color=LINE, pt=1.25, head=False)
arrow(AX, NY + NH / 2, RX, NY + NH / 2)

# ---- legend: colour circles + labels ----
lx = LX
for col, name in [(PEACH, "Users"), (GREEN, "Our backend"), (BLUE, "Azure data + messaging"),
                  (PURPLE, "Identity + secrets"), (GREY, "CI/CD")]:
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(lx), Inches(7.08), Inches(0.14), Inches(0.14))
    dot.fill.solid(); dot.fill.fore_color.rgb = RGBColor.from_string(EDGE[col])
    dot.line.color.rgb = RGBColor.from_string(EDGE[col]); dot.line.width = Pt(1)
    c.kill_effects(dot)
    w = 0.065 * len(name) + 0.08
    c.text(s, lx + 0.2, 7.05, w, None, name, font=st.font_body, size=9, color=MUTED)
    lx += 0.2 + w + 0.2

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT}")
