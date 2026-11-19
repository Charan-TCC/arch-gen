"""Clean, minimal architecture diagram for EDTIH matching the exact reference layout.

Layout:
- Left: User -> splits into (1) Azure WAF (in soft blue container) -> Frontend (React), and (2) Microsoft SSO -> App Service (Frontend).
- Center: App Service (in soft green container) containing Frontend (React) <---> Backend (Nodejs).
- Top: Power BI (Report embedding in react app) -> Frontend (React).
- Bottom: Azure Devops (Automated Deployment for Dev, UAT, Prod) --[CI/CD Pipeline]--> App Service.
- Right: Backend (Nodejs) -> splits into (1) Storage (in soft yellow container: ADLS, SQL Server, Databricks Delta Tables), and (2) Azure Data Factory (Triggers ADF pipeline).
"""
from __future__ import annotations

import sys
from pathlib import Path

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
BG_PAGE = "FFFFFF"
INK_DARK = "0F172A"       # Text and card borders
SUBTEXT = "334155"        # Subtitle italic text
ARROW_COLOR = "000000"    # Black connection lines matching reference

# Soft Container Fills & Lines
WAF_BG = "DBEAFE"         # Soft light blue
WAF_LINE = "93C5FD"

APP_BG = "DCFCE7"         # Soft light green
APP_LINE = "86EFAC"

STORAGE_BG = "FEF3C7"     # Soft light yellow / gold
STORAGE_LINE = "FCD34D"

CARD_BG = "FFFFFF"
CARD_LINE = "1E293B"      # Crisp dark card border

def C(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#"))

def _no_shadow(shp):
    shp.shadow.inherit = False
    spPr = shp._element.spPr
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)
    spPr.append(shp._element.makeelement(qn("a:effectLst"), {}))

def rrect(slide, x, y, w, h, fill, line, line_pt=0.9, radius=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_pt)
    _no_shadow(shp)
    return shp

def place_icon(slide, key, cx_in, cy_in, size_in):
    """Place icon centered at (cx_in, cy_in)."""
    e = R.get(key)
    recolor = f"#{e.brand_hex}" if (e.recolorable and e.brand_hex) else None
    png = emf.to_picture(e.path, recolor_hex=recolor, prefer_vector=False)
    asp = emf.svg_aspect(e.path)
    if asp >= 1:
        w = size_in
        h = size_in / asp
    else:
        h = size_in
        w = size_in * asp
    slide.shapes.add_picture(
        str(png),
        Inches(cx_in - w / 2),
        Inches(cy_in - h / 2),
        width=Inches(w),
        height=Inches(h),
    )

def text_box(slide, x, y, w, h, runs, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(4000)
    tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for (txt, size, bold, italic, color) in para:
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Segoe UI"
            r.font.color.rgb = C(color)
    return box

def draw_arrow(slide, x1, y1, x2, y2, color=ARROW_COLOR, pt=1.1, double_headed=False):
    """Straight arrow from (x1, y1) to (x2, y2)."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(pt)
    ln = cn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if double_headed:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    cn.shadow.inherit = False
    return cn

def draw_elbow_arrow(slide, points: list[tuple[float, float]], color=ARROW_COLOR, pt=1.1):
    """Draws multi-segment orthogonal lines with an arrowhead at the last point."""
    if len(points) < 2:
        return
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        is_last = (i == len(points) - 2)
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        cn.line.color.rgb = C(color)
        cn.line.width = Pt(pt)
        if is_last:
            ln = cn.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
        cn.shadow.inherit = False


def build_deck(out_path: str | Path = "out/edtih_architecture.pptx") -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # ---------------------------------------------------------------- GEOMETRY & COORDINATES
    # Main Horizontal Axis
    Y_MAIN = 3.65  # Center Y for User, WAF, Frontend, Backend

    # 1. Containers
    # A. Azure WAF Container (Blue)
    WAF_BOX = {"x": 1.70, "y": 2.60, "w": 2.10, "h": 2.10}
    rrect(slide, WAF_BOX["x"], WAF_BOX["y"], WAF_BOX["w"], WAF_BOX["h"], WAF_BG, WAF_LINE, line_pt=1.0, radius=0.06)

    # B. App Service Container (Green)
    APP_BOX = {"x": 4.30, "y": 2.60, "w": 4.30, "h": 2.10}
    rrect(slide, APP_BOX["x"], APP_BOX["y"], APP_BOX["w"], APP_BOX["h"], APP_BG, APP_LINE, line_pt=1.0, radius=0.06)
    # App Service Badge (Icon + Text pinned at top-left)
    place_icon(slide, "azure.app-services", APP_BOX["x"] + 0.12, APP_BOX["y"] - 0.10, 0.26)
    text_box(
        slide, APP_BOX["x"] + 0.28, APP_BOX["y"] - 0.22, 1.40, 0.25,
        [[("App Service", 9, True, False, INK_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # C. Storage Container (Yellow/Gold)
    STORAGE_BOX = {"x": 9.70, "y": 0.90, "w": 2.65, "h": 4.15}
    rrect(slide, STORAGE_BOX["x"], STORAGE_BOX["y"], STORAGE_BOX["w"], STORAGE_BOX["h"], STORAGE_BG, STORAGE_LINE, line_pt=1.0, radius=0.06)
    # Storage Label at top-left
    text_box(
        slide, STORAGE_BOX["x"] + 0.05, STORAGE_BOX["y"] - 0.28, 1.20, 0.25,
        [[("Storage", 10, True, False, INK_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # ---------------------------------------------------------------- CARDS
    # 1. User
    USER_POS = {"cx": 0.85, "cy": Y_MAIN - 0.12}
    place_icon(slide, "generic.user", USER_POS["cx"], USER_POS["cy"], 0.44)
    text_box(
        slide, USER_POS["cx"] - 0.50, USER_POS["cy"] + 0.28, 1.00, 0.25,
        [[("User", 9, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 2. Azure WAF Card (inside blue container)
    WAF_CARD = {"x": 1.95, "y": 2.88, "w": 1.60, "h": 1.54}
    rrect(slide, WAF_CARD["x"], WAF_CARD["y"], WAF_CARD["w"], WAF_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.web-application-firewall-policies-waf", WAF_CARD["x"] + WAF_CARD["w"] / 2, WAF_CARD["y"] + 0.32, 0.42)
    text_box(
        slide, WAF_CARD["x"] + 0.05, WAF_CARD["y"] + 0.60, WAF_CARD["w"] - 0.10, 0.85,
        [
            [("Azure WAF", 8.5, True, False, INK_DARK)],
            [("Protection Against\nWeb Vulnerabilities", 7.5, False, True, SUBTEXT)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # 3. Microsoft SSO Card (below WAF)
    SSO_CARD = {"x": 1.70, "y": 5.25, "w": 1.55, "h": 1.10}
    rrect(slide, SSO_CARD["x"], SSO_CARD["y"], SSO_CARD["w"], SSO_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.entra-id-protection", SSO_CARD["x"] + SSO_CARD["w"] / 2, SSO_CARD["y"] + 0.36, 0.40)
    text_box(
        slide, SSO_CARD["x"] + 0.05, SSO_CARD["y"] + 0.68, SSO_CARD["w"] - 0.10, 0.35,
        [[("Microsoft SSO", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 4. Frontend (React) Card (inside App Service container)
    FE_CARD = {"x": 4.65, "y": 2.92, "w": 1.45, "h": 1.46}
    rrect(slide, FE_CARD["x"], FE_CARD["y"], FE_CARD["w"], FE_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    text_box(
        slide, FE_CARD["x"] + 0.05, FE_CARD["y"] + 0.10, FE_CARD["w"] - 0.10, 0.24,
        [[("Frontend", 8.5, True, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )
    place_icon(slide, "tech.react", FE_CARD["x"] + FE_CARD["w"] / 2, FE_CARD["y"] + 0.68, 0.44)
    text_box(
        slide, FE_CARD["x"] + 0.05, FE_CARD["y"] + 1.05, FE_CARD["w"] - 0.10, 0.25,
        [[("React", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 5. Backend (Nodejs) Card (inside App Service container)
    BE_CARD = {"x": 6.80, "y": 2.92, "w": 1.45, "h": 1.46}
    rrect(slide, BE_CARD["x"], BE_CARD["y"], BE_CARD["w"], BE_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    text_box(
        slide, BE_CARD["x"] + 0.05, BE_CARD["y"] + 0.10, BE_CARD["w"] - 0.10, 0.24,
        [[("Backend", 8.5, True, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )
    place_icon(slide, "tech.express", BE_CARD["x"] + BE_CARD["w"] / 2, BE_CARD["y"] + 0.68, 0.44)
    text_box(
        slide, BE_CARD["x"] + 0.05, BE_CARD["y"] + 1.05, BE_CARD["w"] - 0.10, 0.25,
        [[("Nodejs", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 6. Power BI Card (Top)
    PBI_CARD = {"x": 4.65, "y": 0.60, "w": 1.65, "h": 1.40}
    rrect(slide, PBI_CARD["x"], PBI_CARD["y"], PBI_CARD["w"], PBI_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.power-bi-embedded", PBI_CARD["x"] + PBI_CARD["w"] / 2, PBI_CARD["y"] + 0.32, 0.38)
    text_box(
        slide, PBI_CARD["x"] + 0.05, PBI_CARD["y"] + 0.58, PBI_CARD["w"] - 0.10, 0.75,
        [
            [("Power BI", 8.5, False, False, INK_DARK)],
            [("Report embedding in\nreact app", 7.5, False, True, SUBTEXT)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # 7. Azure DevOps Card (Bottom)
    DEVOPS_CARD = {"x": 6.30, "y": 5.40, "w": 2.45, "h": 1.45}
    rrect(slide, DEVOPS_CARD["x"], DEVOPS_CARD["y"], DEVOPS_CARD["w"], DEVOPS_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.devops", DEVOPS_CARD["x"] + DEVOPS_CARD["w"] / 2, DEVOPS_CARD["y"] + 0.32, 0.38)
    text_box(
        slide, DEVOPS_CARD["x"] + 0.05, DEVOPS_CARD["y"] + 0.58, DEVOPS_CARD["w"] - 0.10, 0.80,
        [
            [("Azure Devops", 8.5, False, False, INK_DARK)],
            [("Automated Deployment for\nDevelopment, UAT, and Production\nEnvironments.", 7.5, False, True, SUBTEXT)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # 8. Storage Cards (Inside Storage container)
    # ADLS (Azure Storage icon)
    ADLS_CARD = {"x": 9.90, "y": 1.15, "w": 2.25, "h": 1.00}
    rrect(slide, ADLS_CARD["x"], ADLS_CARD["y"], ADLS_CARD["w"], ADLS_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.storage-accounts", ADLS_CARD["x"] + ADLS_CARD["w"] / 2, ADLS_CARD["y"] + 0.38, 0.44)
    text_box(
        slide, ADLS_CARD["x"] + 0.05, ADLS_CARD["y"] + 0.68, ADLS_CARD["w"] - 0.10, 0.25,
        [[("ADLS", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # SQL Server (Red wireframe icon)
    SQL_CARD = {"x": 9.90, "y": 2.48, "w": 2.25, "h": 1.00}
    rrect(slide, SQL_CARD["x"], SQL_CARD["y"], SQL_CARD["w"], SQL_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.sql", SQL_CARD["x"] + SQL_CARD["w"] / 2, SQL_CARD["y"] + 0.38, 0.44)
    text_box(
        slide, SQL_CARD["x"] + 0.05, SQL_CARD["y"] + 0.68, SQL_CARD["w"] - 0.10, 0.25,
        [[("SQL Server", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # Databricks Delta Tables
    DB_CARD = {"x": 9.90, "y": 3.80, "w": 2.25, "h": 1.00}
    rrect(slide, DB_CARD["x"], DB_CARD["y"], DB_CARD["w"], DB_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.databricks", DB_CARD["x"] + DB_CARD["w"] / 2, DB_CARD["y"] + 0.38, 0.44)
    text_box(
        slide, DB_CARD["x"] + 0.05, DB_CARD["y"] + 0.68, DB_CARD["w"] - 0.10, 0.25,
        [[("Databricks Delta Tables", 8.5, False, False, INK_DARK)]],
        align=PP_ALIGN.CENTER,
    )

    # 9. Azure Data Factory Card (Bottom-Right)
    ADF_CARD = {"x": 9.70, "y": 5.30, "w": 2.65, "h": 1.35}
    rrect(slide, ADF_CARD["x"], ADF_CARD["y"], ADF_CARD["w"], ADF_CARD["h"], CARD_BG, CARD_LINE, line_pt=0.9, radius=0.06)
    place_icon(slide, "azure.data-factories", ADF_CARD["x"] + ADF_CARD["w"] / 2, ADF_CARD["y"] + 0.32, 0.40)
    text_box(
        slide, ADF_CARD["x"] + 0.05, ADF_CARD["y"] + 0.58, ADF_CARD["w"] - 0.10, 0.70,
        [
            [("Azure Data Factory", 8.5, False, False, INK_DARK)],
            [("Triggers ADF Data Factory pipeline", 7.5, False, True, SUBTEXT)],
        ],
        align=PP_ALIGN.CENTER,
    )

    # ---------------------------------------------------------------- ARROWS (Crisp, clean black connectors)
    # 1. User -> Azure WAF (Straight horizontal arrow)
    draw_arrow(slide, 1.15, Y_MAIN, WAF_CARD["x"], Y_MAIN, color=ARROW_COLOR, pt=1.1)

    # 2. User -> Microsoft SSO (Elbow arrow down and right)
    draw_elbow_arrow(
        slide,
        [
            (0.85, USER_POS["cy"] + 0.45),
            (0.85, SSO_CARD["y"] + SSO_CARD["h"] / 2),
            (SSO_CARD["x"], SSO_CARD["y"] + SSO_CARD["h"] / 2),
        ],
        color=ARROW_COLOR,
        pt=1.1,
    )

    # 3. Azure WAF -> Frontend (React) (Straight horizontal arrow)
    draw_arrow(slide, WAF_CARD["x"] + WAF_CARD["w"], Y_MAIN, FE_CARD["x"], Y_MAIN, color=ARROW_COLOR, pt=1.1)

    # 4. Microsoft SSO -> App Service (Frontend) (Elbow arrow right and up into bottom of App Service)
    draw_elbow_arrow(
        slide,
        [
            (SSO_CARD["x"] + SSO_CARD["w"], SSO_CARD["y"] + SSO_CARD["h"] / 2),
            (4.70, SSO_CARD["y"] + SSO_CARD["h"] / 2),
            (4.70, APP_BOX["y"] + APP_BOX["h"]),
        ],
        color=ARROW_COLOR,
        pt=1.1,
    )

    # 5. Frontend <---> Backend (Bidirectional arrow)
    draw_arrow(slide, FE_CARD["x"] + FE_CARD["w"], Y_MAIN, BE_CARD["x"], Y_MAIN, color=ARROW_COLOR, pt=1.2, double_headed=True)

    # 6. Power BI -> Frontend (React) (Straight vertical downward arrow)
    draw_arrow(slide, FE_CARD["x"] + FE_CARD["w"] / 2, PBI_CARD["y"] + PBI_CARD["h"], FE_CARD["x"] + FE_CARD["w"] / 2, FE_CARD["y"], color=ARROW_COLOR, pt=1.1)

    # 7. Azure DevOps -> App Service (Straight vertical upward arrow + CI/CD badge)
    devops_arrow_x = 7.50
    draw_arrow(slide, devops_arrow_x, DEVOPS_CARD["y"], devops_arrow_x, APP_BOX["y"] + APP_BOX["h"], color=ARROW_COLOR, pt=1.1)
    # CI/CD Pipeline badge on the arrow (left icon, right text)
    place_icon(slide, "azure.monitor-pipeline", devops_arrow_x - 0.18, 4.95, 0.20)
    text_box(
        slide, devops_arrow_x + 0.06, 4.85, 1.20, 0.24,
        [[("CI/CD Pipeline", 8, False, False, INK_DARK)]],
        align=PP_ALIGN.LEFT,
    )

    # 8. Backend (Nodejs) -> Storage & Azure Data Factory (Forked arrow)
    fork_x = 9.00
    # Stem from Backend right edge
    draw_arrow(slide, BE_CARD["x"] + BE_CARD["w"], Y_MAIN, fork_x, Y_MAIN, color=ARROW_COLOR, pt=1.1)
    # Branch 1 (Up to Storage Container)
    draw_elbow_arrow(
        slide,
        [
            (fork_x, Y_MAIN),
            (fork_x, 2.98),
            (STORAGE_BOX["x"], 2.98),
        ],
        color=ARROW_COLOR,
        pt=1.1,
    )
    # Branch 2 (Down to Azure Data Factory)
    draw_elbow_arrow(
        slide,
        [
            (fork_x, Y_MAIN),
            (fork_x, ADF_CARD["y"] + ADF_CARD["h"] / 2),
            (ADF_CARD["x"], ADF_CARD["y"] + ADF_CARD["h"] / 2),
        ],
        color=ARROW_COLOR,
        pt=1.1,
    )

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    import os; os.makedirs("out", exist_ok=True)
    from archgen.shared.geomcheck import fix_presentation
    fix_presentation(prs)  # hug text boxes, snap near-miss alignment
    prs.save(str(out_path))
    print(f"Rendered clean minimal presentation to {out_path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "out/edtih_architecture.pptx"
    build_deck(out)
