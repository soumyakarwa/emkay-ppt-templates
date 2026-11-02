"""
Emkay presentation templates, built from the Emkay brand guidelines (v3).
Creates three 16:9 decks:
  Emkay_Master_Template.pptx                (Trust Teal)
  Emkay_Wealth_Management_Template.pptx     (Wealth Indigo)
  Emkay_Investment_Managers_Template.pptx   (Investment Pine)

Setup:  pip install python-pptx
Run:    python build_emkay_templates.py
"""
import math
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                            XL_LABEL_POSITION, XL_TICK_LABEL_POSITION)

# ---------------- Canvas and grid ----------------
W, H = 13.333, 7.5      # 16:9, inches
M = 0.75                # right / bottom margin
# Checkerboard corner sizes (print minimums assume A4 handouts at ~0.88 scale):
C1 = 1.25               # knight only: 2.5in mark = ~56mm printed (min 55mm); knight >= 20mm
C2 = 1.9                # knight + emkay: wordmark box ~1.8in = ~40mm printed (min 40mm)
CX = CY = C1 + 0.5      # content well starts clear of the knight-only corner

# ---------------- Brand colours (all from the guide: slides 13, 14, 27, 39) ----------------
TEAL, ORANGE, PEARL, INK = "074974", "EE8A22", "FAF9F2", "201A12"
WHITE, AMBER, GREY, RULE = "FFFFFF", "FCB537", "BCBEC0", "DBDFDE"
BEIGE = "FCF8F4"

# ---------------- Type: Aptos for PowerPoint (slides 17, 18) ----------------
FONT = "Aptos"          # change here (and rebuild) to switch typeface later
F_REG, F_SEMI, F_BLACK = FONT, f"{FONT} SemiBold", f"{FONT} Black"
X = 44                   # slide title = x (Bold)
SUB = round(X / 2.5)     # sub-headline = x / 2.5  -> 18 pt (SemiBold)
BODY = round(SUB / 1.5)  # body = sub / 1.5        -> 12 pt (Regular)
DISPLAY = 54             # opening and section titles (Black)

# The one substitution (slide 37): teal -> Indigo / Pine. Orange, Pearl and the grid never change.
COLOURWAYS = {
    "master": dict(
        file="Emkay_Master_Template.pptx", name="Teal",
        lockup="Stacked lock-up with tagline",
        primary=TEAL, ground=PEARL, grid=GREY,
        table_band="E8DED3", tint_mid="E8DED3", tint_pale=RULE),
    "wealth": dict(
        file="Emkay_Wealth_Management_Template.pptx", name="Indigo",
        lockup="Stacked lock-up with 'Wealth Management' descriptor",
        primary="242B7F", ground=BEIGE, grid="4455B7",
        table_band="E4E7FA", tint_mid="4455B7", tint_pale="E4E7FA"),
    "investment": dict(
        file="Emkay_Investment_Managers_Template.pptx", name="Pine",
        lockup="Stacked lock-up with 'Investment Managers' descriptor",
        primary="215224", ground=BEIGE, grid="416944",
        table_band="D9F1DB", tint_mid="416944", tint_pale="D9F1DB"),
}
SECTION_FILL = ORANGE   # stays orange in every colourway (slide 37 rule); guide visuals show tints

# ---------------- Theme: default fonts and colours for anything users add later ----------------
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

def apply_theme(prs, cw):
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    ns = {"a": A_NS}
    for tag in ("majorFont", "minorFont"):
        root.find(f".//a:fontScheme/a:{tag}/a:latin", ns).set("typeface", FONT)
    colours = dict(dk1=INK, lt1=WHITE, dk2=cw["primary"], lt2=cw["ground"],
                   accent1=cw["primary"], accent2=ORANGE, accent3=AMBER, accent4=GREY,
                   accent5=cw["tint_mid"], accent6=cw["tint_pale"],
                   hlink=cw["primary"], folHlink=cw["primary"])
    for name, hexv in colours.items():
        el = root.find(f".//a:clrScheme/a:{name}", ns)
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, f"{{{A_NS}}}srgbClr", val=hexv)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

# ---------------- Helpers ----------------
def rgb(h):
    return RGBColor.from_string(h)

def new_slide(prs, colour, note):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    f = s.background.fill
    f.solid(); f.fore_color.rgb = rgb(colour)
    s.notes_slide.notes_text_frame.text = note
    return s

def rect(s, x, y, w, h, fill, line=None, name=None):
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line); shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    if name:
        shp.name = name
    return shp

def text(s, x, y, w, h, lines, size=BODY, font=F_REG, color=INK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=0):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(lines if isinstance(lines, list) else [lines]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(after)
        r = p.add_run(); r.text = ln
        r.font.name = font; r.font.size = Pt(size)
        r.font.bold = bold; r.font.color.rgb = rgb(color)
    return tb

def bullets(s, x, y, w, h, items, cw, size=BODY + 2):
    """En-dash bullets in the brand colour (no single squares)."""
    tf = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r1 = p.add_run(); r1.text = "\u2013  "
        r1.font.name = F_SEMI; r1.font.size = Pt(size); r1.font.color.rgb = rgb(cw["primary"])
        r2 = p.add_run(); r2.text = it
        r2.font.name = F_REG; r2.font.size = Pt(size); r2.font.color.rgb = rgb(INK)

def placeholder(s, x, y, w, h, label, fill=WHITE, line=GREY, color=TEAL,
                size=11, name="Placeholder"):
    shp = rect(s, x, y, w, h, fill, line, name)
    shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.text = label
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
            r.font.name = F_REG; r.font.size = Pt(size); r.font.color.rgb = rgb(color)
    return shp

def main_logo(s, x, y, w, h, cw, dark=False):
    """Title slide only: stacked lock-up with tagline (master) or descriptor (divisions)."""
    label = f"LOGO LOCK-UP\n{cw['lockup']}" + ("\nwhite-out version" if dark else "")
    colour = PEARL if dark else cw["primary"]
    placeholder(s, x, y, w, h, label, fill=None, line=PEARL if dark else GREY,
                color=colour, size=11, name="Logo Placeholder")

def corner(s, cw, c, wordmark=False, light=WHITE):
    """Top-left checkerboard corner, the only logo position on inner slides.
    Knight (or knight + emkay) in the light cell; two orange cells step away on the diagonal."""
    rect(s, 0, 0, c, c, light)
    rect(s, c, 0, c, c, ORANGE)
    rect(s, 0, c, c, c, ORANGE)
    if wordmark:
        bw, bh = c - 0.1, c - 0.3
        label = f"KNIGHT + EMKAY\n{cw['name']} wordmark"
    else:
        bw = bh = 0.92
        label = "KNIGHT"
    placeholder(s, (c - bw) / 2, (c - bh) / 2, bw, bh, label, fill=None,
                line=ORANGE, color=ORANGE, size=9, name="Logo Placeholder")

def title(s, cw, t, sub=None, c=C1):
    """Title sits in the top grid row, to the right of the checkerboard."""
    x = 2 * c + 0.4
    y = max(0.2, (c - 1.0) / 2)
    text(s, x, y, W - x - M, 0.65, t, X, F_REG, cw["primary"], bold=True)
    if sub:
        text(s, x, y + 0.65, W - x - M, 0.35, sub, SUB, F_SEMI, INK)

P1 = ("Body copy sits here in Aptos Regular at 12 pt, the sub-headline size divided by 1.5. "
      "Keep paragraphs short and let the grid give the page room to breathe.")
P2 = ("Use sub-headings in Aptos SemiBold to break long arguments into scannable blocks. "
      "Neutrals carry most of the surface; brand colours work as accents.")
P3 = "If the copy runs long, move it to a second slide rather than shrinking the type."

# ---------------- Slides ----------------
def s_opening_light(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Title slide, light. Stacked lock-up with tagline/descriptor, centred: the only full lock-up "
        "in the deck. Orange band = 30% of height, orange in every colourway.")
    main_logo(s, (W - 3.8) / 2, 0.55, 3.8, 3.6, cw)
    band = H * 0.30
    rect(s, 0, H - band, W, band, ORANGE)
    text(s, M, H - band + 0.35, W - 2 * M, 0.9, "PRESENTATION TITLE", DISPLAY,
         F_BLACK, PEARL, align=PP_ALIGN.CENTER)
    text(s, M, H - band + 1.35, W - 2 * M, 0.45,
         "Sub headline  |  Presenter name  |  Month YYYY", SUB, F_SEMI, PEARL,
         align=PP_ALIGN.CENTER)

def s_opening_dark(prs, cw):
    s = new_slide(prs, cw["primary"],
        "Title slide, dark. Brand colour background (Teal / Indigo / Pine) with orange bar. "
        "White-out stacked lock-up with tagline/descriptor.")
    rect(s, 0, H - 0.4, W, 0.4, ORANGE)
    main_logo(s, 1.0, 1.45, 3.9, 3.8, cw, dark=True)
    text(s, 5.6, 2.35, 7.0, 1.0, "TITLE SLIDE", DISPLAY, F_BLACK, PEARL, anchor=MSO_ANCHOR.BOTTOM)
    text(s, 5.6, 3.45, 7.0, 0.5, "Sub headline goes here", SUB + 4, F_SEMI, PEARL)
    text(s, 5.6, 4.1, 7.0, 0.4, "Presenter name  |  Month YYYY", BODY + 2, F_REG, PEARL)

def s_agenda(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Agenda. Brand-colour panel, knight-only checkerboard top left (text slide).")
    pw = 4.2
    rect(s, 0, 0, pw, H, cw["primary"])
    corner(s, cw, C1, light=cw["ground"])
    text(s, M, 2.9, pw - 1.2, 0.8, "Agenda", X, F_REG, PEARL, bold=True)
    text(s, M, 3.75, pw - 1.2, 0.9, "What we will cover today", SUB, F_SEMI, PEARL)
    items = [("Agenda item one", "One line describing what this part covers."),
             ("Agenda item two", "One line describing what this part covers."),
             ("Agenda item three", "One line describing what this part covers."),
             ("Agenda item four", "One line describing what this part covers."),
             ("Agenda item five", "One line describing what this part covers.")]
    x0, y, row = pw + 0.8, 0.8, 1.15
    for i, (t, d) in enumerate(items):
        text(s, x0, y, 0.9, 0.6, f"{i + 1:02d}", 28, F_BLACK, cw["primary"])
        text(s, x0 + 1.1, y + 0.02, W - x0 - 1.1 - M, 0.4, t, SUB, F_SEMI, cw["primary"])
        text(s, x0 + 1.1, y + 0.45, W - x0 - 1.1 - M, 0.4, d, BODY, F_REG, INK)
        rect(s, x0, y + row - 0.15, W - x0 - M, 0.01, RULE)
        y += row

def s_section(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Section divider. Stepped grid with knight + emkay in the corner cell (guide p.21, p.37). "
        "Grid lines: tint, not the full brand colour. Knight + emkay is at its 40mm print minimum.")
    c = H / 4
    for col in range(math.ceil(W / c)):
        rect(s, col * c, 0, min(c, W - col * c), c, cw["ground"], cw["grid"])
    for row in range(1, 4):
        for col in range(2):
            rect(s, col * c, row * c, c, c, cw["ground"], cw["grid"])
    corner(s, cw, c, wordmark=True, light=cw["ground"])
    rect(s, 2 * c, c, W - 2 * c, H - c, SECTION_FILL)
    text(s, 2 * c + 0.6, c + 1.0, W - 2 * c - 1.2, 1.9, ["SECTION", "HEADING"],
         DISPLAY, F_BLACK, INK)
    text(s, 2 * c + 0.6, c + 3.0, W - 2 * c - 1.2, 0.5, "Sub line goes here",
         SUB + 2, F_SEMI, INK)

def s_text_heavy(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Text-heavy slide. Knight-only corner. Title Aptos Bold x, sub SemiBold x/2.5, body Regular sub/1.5.")
    pw = 4.2
    rect(s, 0, 0, pw, H, cw["primary"])
    corner(s, cw, C1, light=cw["ground"])
    text(s, M, 2.85, pw - 1.2, 1.5, ["Slide", "heading"], X, F_REG, PEARL, bold=True)
    text(s, M, 4.45, pw - 1.2, 0.7, "Sub headline in Aptos SemiBold", SUB, F_SEMI, PEARL)
    text(s, M, 5.3, pw - 1.2, 1.7, P3, BODY, F_REG, PEARL)
    x0, gap = pw + 0.7, 0.5
    colw = (W - x0 - M - gap) / 2
    for j in range(2):
        x = x0 + j * (colw + gap)
        text(s, x, 0.75, colw, 0.4, f"Sub-heading {j + 1}", SUB, F_SEMI, cw["primary"])
        text(s, x, 1.3, colw, 5.5, [P1, P2, P3], BODY, F_REG, INK, after=10)

def s_image_content(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Image and content (guide p.36). Knight + emkay corner (image-led slide). Band = 40% of height.")
    ih = H * 0.6
    placeholder(s, 0, 0, W, ih,
        "IMAGE\nEditorial artwork: photo cut-out printed in one colour (halftone) on a flat brand shape.\n"
        "The cut-out breaks one edge of its shape. Never more than two colours.",
        fill=cw["ground"], line=GREY, color=cw["primary"], size=14)
    corner(s, cw, C2, wordmark=True)
    rect(s, 0, ih, W, H - ih, cw["primary"])
    text(s, M, ih + 0.4, 9.5, 0.8, "Headline comes here", X, F_REG, ORANGE, bold=True)
    text(s, M, ih + 1.35, 9.0, 1.2,
         "Supporting copy in Pearl White on the colour band. One message per slide.",
         SUB, F_REG, PEARL)

def style_chart(chart, cw):
    palette = [cw["primary"], ORANGE, AMBER, GREY]   # fixed order (guide p.39)
    chart.font.name = F_REG; chart.font.size = Pt(12); chart.font.color.rgb = rgb(INK)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    plot = chart.plots[0]
    plot.gap_width = 80
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = "0.0"; dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(11)
    for i, ser in enumerate(plot.series):
        ser.invert_if_negative = False
        ser.format.fill.solid(); ser.format.fill.fore_color.rgb = rgb(palette[i % 4])
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb(RULE)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(11)
    ca = chart.category_axis
    ca.format.line.color.rgb = rgb(GREY)
    ca.tick_labels.font.size = Pt(11)
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW

def s_chart(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Chart slide. Knight-only corner. Chart colours in order: brand colour, Growth Orange, "
        "Amber FCB537, Grey BCBEC0. Never use software default palettes.")
    corner(s, cw, C1)
    title(s, cw, "Chart title goes here", "Sub headline: the one takeaway from this chart")
    rect(s, CX, CY, W - M - CX, 4.85, WHITE)
    data = CategoryChartData()
    data.categories = ["1M", "3M", "6M", "1Y", "3Y"]
    data.add_series("Portfolio", (6.4, 18.1, 26.9, 39.6, 13.2))
    data.add_series("Benchmark", (-3.4, 8.7, 16.7, 36.7, 15.9))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(CX + 0.3), Inches(CY + 0.15),
                            Inches(W - M - CX - 0.6), Inches(4.55), data)
    style_chart(gf.chart, cw)
    text(s, CX, 6.75, 9.0, 0.3, "Sample data. Source and notes line in Aptos Regular, 10 pt.",
         10, F_REG, INK)

def s_kpi(prs, cw):
    s = new_slide(prs, cw["ground"],
        "KPI slide. Knight-only corner. Growth Orange marks only the ONE figure that matters.")
    corner(s, cw, C1)
    title(s, cw, "Key figures at a glance", "Sub headline: what these numbers mean")
    gap = 0.4
    cwid = (W - M - CX - 2 * gap) / 3
    y, h = 2.2, 3.2
    items = [("Assets under advice", "Rs 13.8 bn", "As on DD Month YYYY (sample)"),
             ("Portfolio TWRR", "39.6%", "Net of fees and expenses (sample)"),
             ("Net inflows, YTD", "Rs 21.4 bn", "As on DD Month YYYY (sample)")]
    for i, (lab, val, cap) in enumerate(items):
        x = CX + i * (cwid + gap)
        rect(s, x, y, cwid, h, WHITE)
        rect(s, x, y, cwid, 0.06, cw["primary"])      # thin rule, not a square
        text(s, x + 0.4, y + 0.6, cwid - 0.8, 0.4, lab, 14, F_SEMI, INK)
        text(s, x + 0.4, y + 1.15, cwid - 0.8, 0.9, val, 48, F_BLACK,
             ORANGE if i == 0 else cw["primary"])
        text(s, x + 0.4, y + 2.35, cwid - 0.8, 0.4, cap, BODY, F_REG, INK)

def s_table(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Table. Knight-only corner. Header in the brand colour; tints only for table bands (guide p.14).")
    corner(s, cw, C1)
    title(s, cw, "Table title goes here", "Coverage summary (sample data)")
    rows = [("Company", "Rating", "Target price (Rs)", "Upside"),
            ("Allied Blenders", "Buy", "800", "+17%"),
            ("Radico Khaitan", "Buy", "3,700", "+16%"),
            ("Sula Vineyards", "Add", "275", "+8%"),
            ("United Breweries", "Reduce", "1,700", "-6%")]
    tw = W - M - CX
    gf = s.shapes.add_table(len(rows), 4, Inches(CX), Inches(2.2),
                            Inches(tw), Inches(0.6 * len(rows)))
    tbl = gf.table
    tbl.horz_banding = False
    for i, wd in enumerate([tw - 6.8, 2.1, 2.6, 2.1]):
        tbl.columns[i].width = Inches(wd)
    for r, row in enumerate(rows):
        tbl.rows[r].height = Inches(0.6)
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(cw["primary"] if r == 0 else
                                           (cw["table_band"] if r % 2 == 0 else WHITE))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.2)
            tf = cell.text_frame
            tf.text = val
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            for run in p.runs:
                run.font.name = F_SEMI if r == 0 else F_REG
                run.font.size = Pt(14)
                run.font.color.rgb = rgb(PEARL if r == 0 else INK)
    text(s, CX, 5.4, 9.0, 0.3, "Source and notes line in Aptos Regular, 10 pt.", 10, F_REG, INK)

def s_comparison(prs, cw):
    s = new_slide(prs, cw["ground"], "Two-column comparison. Knight-only corner.")
    corner(s, cw, C1)
    title(s, cw, "Comparison title goes here", "Sub headline framing the choice")
    gap = 0.5
    cwid = (W - M - CX - gap) / 2
    y, h = 2.2, 4.5
    cols = [("Option A", ["First point of comparison", "Second point of comparison",
                          "Third point of comparison", "Fourth point of comparison"]),
            ("Option B", ["First point of comparison", "Second point of comparison",
                          "Third point of comparison", "Fourth point of comparison"])]
    for j, (head, pts) in enumerate(cols):
        x = CX + j * (cwid + gap)
        rect(s, x, y, cwid, h, WHITE)
        rect(s, x, y, cwid, 0.7, cw["primary"])
        text(s, x + 0.35, y, cwid - 0.7, 0.7, head, 20, F_SEMI, PEARL, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x + 0.35, y + 1.0, cwid - 0.7, h - 1.3, pts, cw)

def s_quote(prs, cw):
    s = new_slide(prs, cw["primary"],
        "Quote. Knight-only corner. One message only. The cut-out portrait breaks the edge of the "
        "orange block (guide p.25).")
    corner(s, cw, C1, light=cw["ground"])
    text(s, CX, 1.45, 1.5, 1.1, "\u201C", 120, F_BLACK, ORANGE)
    text(s, CX, 2.55, 6.9, 2.4,
         "One quote per slide, set large enough to read from the back of the room.",
         32, F_REG, PEARL, bold=True)
    text(s, CX, 5.2, 6.9, 0.4, "Speaker Name", SUB, F_SEMI, PEARL)
    text(s, CX, 5.6, 6.9, 0.4, "Designation, Emkay", BODY + 2, F_REG, PEARL)
    rect(s, 8.9, 3.0, 3.7, H - 3.0, ORANGE)
    placeholder(s, 9.2, 1.5, 3.7, H - 1.5,
        "PORTRAIT CUT-OUT\nBreaks the top edge of the orange block.\nHint of brand colour in clothing or background.",
        fill=None, line=PEARL, color=PEARL, size=11)

def s_team(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Team. Knight + emkay corner (image-led slide). People imagery: natural light, looking out of "
        "frame (guide p.23). Colour OR monochrome per slide, never both (guide p.24).")
    corner(s, cw, C2, wordmark=True)
    title(s, cw, "Leadership team", "One grade per slide: colour or monochrome", c=C2)
    x0, gap = C2 + 0.5, 0.35
    pw = (W - M - x0 - 3 * gap) / 4
    y, ph = 2.3, 3.2
    for i in range(4):
        x = x0 + i * (pw + gap)
        placeholder(s, x, y, pw, ph, "PORTRAIT\nNatural light, looking out of frame",
                    fill=WHITE, line=GREY, color=cw["primary"], size=11)
        text(s, x, y + ph + 0.2, pw, 0.4, "Name Surname", SUB, F_SEMI, cw["primary"])
        text(s, x, y + ph + 0.6, pw, 0.4, "Designation", BODY, F_REG, INK)

def s_full_bleed(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Full-bleed image (guide p.26, p.36). When artwork fills the frame, the checkerboard corner "
        "top left is the only place the identity appears.")
    placeholder(s, 0, 0, W, H,
        "FULL-BLEED IMAGE\nEditorial artwork: one cut-out, one flat brand shape",
        fill=cw["ground"], line=GREY, color=cw["primary"], size=16)
    corner(s, cw, C2, wordmark=True)

def s_thank_you(prs, cw):
    s = new_slide(prs, cw["primary"],
        "Closing slide. Knight-only corner. Orange rule and checkerboard stay orange in every colourway.")
    corner(s, cw, 1.5)
    text(s, 1.15, 3.9, 9.0, 1.0, "Thank you", 60, F_BLACK, PEARL)
    rect(s, 1.15, 5.05, 6.1, 0.04, ORANGE)
    text(s, 1.15, 5.35, 10.5, 1.0,
         ["EMKAY GLOBAL FINANCIAL SERVICES LIMITED",
          "Paragon Centre, C-06, Ground Floor, Pandurang Budhkar Marg, Opp. Century Mills, Worli, Mumbai 400013",
          "www.emkayglobal.com  |  CIN L67120MH1995PLC084899"],
         BODY, F_REG, PEARL, after=4)

SLIDES = [s_opening_light, s_opening_dark, s_agenda, s_section, s_text_heavy,
          s_image_content, s_chart, s_kpi, s_table, s_comparison, s_quote,
          s_team, s_full_bleed, s_thank_you]

def build(key):
    cw = COLOURWAYS[key]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    prs.core_properties.title = cw["file"].replace("_", " ").replace(".pptx", "")
    apply_theme(prs, cw)
    for make in SLIDES:
        make(prs, cw)
    prs.save(cw["file"])
    print("Saved", cw["file"])

if __name__ == "__main__":
    for key in COLOURWAYS:
        build(key)