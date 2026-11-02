"""
Emkay presentation templates (v4), built from the Emkay brand guidelines.
Creates three 16:9 decks:
  Emkay_Master_Template.pptx                (Trust Teal)
  Emkay_Wealth_Management_Template.pptx     (Wealth Indigo)
  Emkay_Investment_Managers_Template.pptx   (Investment Pine)

Setup:  pip install python-pptx
Run:    python build_emkay_templates.py   (close the decks in PowerPoint first)
"""
import math
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                            XL_LABEL_POSITION, XL_TICK_LABEL_POSITION)

# ---------------- Canvas and grid ----------------
W, H = 13.333, 7.5      # 16:9, inches
M = 0.75                # outer margin
C1 = 0.45               # knight-only corner cell (presentation scale, guide p.35/37)
C2 = 1.0                # knight + emkay corner cell (stacked, guide p.22 Template 1)
CX, CY = M, 2.5         # content starts below the corner + title block
FOOT_Y = H - 0.42       # footer baseline row

# ---------------- Brand colours (guide slides 13, 14, 27, 39) ----------------
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

COMPLIANCE = "[Compliance line placeholder]"

# Master follows the master pages. Divisions follow the guide's division pages (slides 35-37,
# 42, 48-53): accents take the division tint; orange stays on the checkerboard corner.
# NOTE: B1B9E6 and 82C487 are sampled from the guide's images, not its swatch tables.
COLOURWAYS = {
    "master": dict(
        file="Emkay_Master_Template.pptx", name="Teal",
        entity="Emkay Global Financial Services Limited",
        lockup_desc="tagline", title_lockup="Stacked",
        primary=TEAL, ground=PEARL,
        title_ground=PEARL, title_band=ORANGE,
        dark_bg=TEAL, dark_bar=ORANGE,
        section_fill=ORANGE, section_text=INK, grid=GREY,
        accent=ORANGE, headline_on_band=ORANGE,
        quote_bg=TEAL, closing_bg=TEAL,
        table_band="E8DED3", theme5="E8DED3", theme6=RULE),
    "wealth": dict(
        file="Emkay_Wealth_Management_Template.pptx", name="Indigo",
        entity="Emkay Wealth Advisory Limited",
        lockup_desc="'Wealth Management' descriptor", title_lockup="Linear",
        primary="242B7F", ground=BEIGE,
        title_ground="E4E7FA", title_band="0D0F26",
        dark_bg="0D0F26", dark_bar="4455B7",
        section_fill="B1B9E6", section_text="242B7F", grid="B1B9E6",
        accent="B1B9E6", headline_on_band=PEARL,
        quote_bg="0D0F26", closing_bg="0D0F26",
        table_band="E4E7FA", theme5="4455B7", theme6="E4E7FA"),
    "investment": dict(
        file="Emkay_Investment_Managers_Template.pptx", name="Pine",
        entity="[Investment Managers legal entity name]",
        lockup_desc="'Investment Managers' descriptor", title_lockup="Linear",
        primary="215224", ground=BEIGE,
        title_ground="D9F1DB", title_band="0F220F",
        dark_bg="0F220F", dark_bar="82C487",
        section_fill="82C487", section_text="215224", grid="416944",
        accent="82C487", headline_on_band=PEARL,
        quote_bg="0F220F", closing_bg="0F220F",
        table_band="D9F1DB", theme5="416944", theme6="D9F1DB"),
}

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
                   accent5=cw["theme5"], accent6=cw["theme6"],
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
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=0, name=None):
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
    if name:
        tb.name = name
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

def main_logo(s, x, y, w, h, cw, kind, dark=False):
    """Title slides only: full lock-up with tagline (master) or descriptor (divisions)."""
    label = f"LOGO LOCK-UP\n{kind} lock-up with {cw['lockup_desc']}" + ("\nwhite-out version" if dark else "")
    colour = PEARL if dark else cw["primary"]
    placeholder(s, x, y, w, h, label, fill=None, line=PEARL if dark else GREY,
                color=colour, size=11, name="Logo Placeholder")

def corner(s, cw, c, wordmark=False, light=WHITE):
    """Top-left checkerboard corner, the only logo position on inner slides.
    Knight (or knight + emkay) in the light cell; two orange cells step away on the diagonal."""
    rect(s, 0, 0, c, c, light)
    rect(s, c, 0, c, c, ORANGE)
    rect(s, 0, c, c, c, ORANGE)
    pad = 0.05
    if wordmark:
        label = "KNIGHT +\nEMKAY"
    else:
        label = "KNIGHT" if c >= 0.8 else ""
    placeholder(s, pad, pad, c - 2 * pad, c - 2 * pad, label, fill=None,
                line=ORANGE, color=ORANGE, size=7, name="Logo Placeholder")

def slide_number(s, x, y, w, h, color):
    """Auto-updating slide number field."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = "Slide Number"
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    fld = etree.Element(qn("a:fld"))
    fld.set("id", "{5C1E7B3A-9D2F-4E8B-A1C6-3F7D2B9E4A10}")
    fld.set("type", "slidenum")
    rpr = etree.SubElement(fld, qn("a:rPr")); rpr.set("lang", "en-US"); rpr.set("sz", "900")
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    etree.SubElement(rpr, qn("a:latin")).set("typeface", FONT)
    etree.SubElement(fld, qn("a:t")).text = "\u2039#\u203A"
    end = p._p.find(qn("a:endParaRPr"))
    if end is not None:
        end.addprevious(fld)
    else:
        p._p.append(fld)

def footer(s, cw, x=CX, color=INK, number=True, w=None):
    """Entity name + compliance line bottom-left; slide number bottom-right."""
    w = w or (W - x - M - 0.8)
    text(s, x, FOOT_Y, w, 0.22, f"{cw['entity']}  |  {COMPLIANCE}", 9, F_REG, color,
         name="Footer")
    if number:
        slide_number(s, W - M - 0.6, FOOT_Y, 0.6, 0.22, color)

def title(s, cw, t, sub=None, c=C1, rule=False):
    """Title sits below the checkerboard corner, left-aligned to the margin (guide p.36, p.44, p.46)."""
    y = 2 * c + 0.3
    text(s, M, y, W - 2 * M, 0.65, t, X, F_REG, cw["primary"], bold=True)
    if sub:
        text(s, M, y + 0.65, W - 2 * M, 0.35, sub, SUB, F_SEMI, INK)
    if rule:
        rect(s, M, y + 1.1, W - 2 * M, 0.02, cw["primary"])

P1 = ("Body copy sits here in Aptos Regular at 12 pt, the sub-headline size divided by 1.5. "
      "Keep paragraphs short and let the grid give the page room to breathe.")
P2 = ("Use sub-headings in Aptos SemiBold to break long arguments into scannable blocks. "
      "Neutrals carry most of the surface; brand colours work as accents.")
P3 = "If the copy runs long, move it to a second slide rather than shrinking the type."

# ---------------- Slides ----------------
def s_opening_light(prs, cw):
    s = new_slide(prs, cw["title_ground"],
        "Title slide, light (guide p.35). The only full lock-up in the deck, with tagline/descriptor. "
        "Band = 30% of height: orange for master, deep shade for divisions.")
    band = H * 0.30
    if cw["title_lockup"] == "Stacked":
        main_logo(s, (W - 3.8) / 2, 0.55, 3.8, 3.6, cw, "Stacked")
    else:  # linear with descriptor: >= 84.2mm printed, so ~3.8in+ on the slide
        lw, lh = 4.8, 1.9
        main_logo(s, (W - lw) / 2, (H - band - lh) / 2, lw, lh, cw, "Linear")
    rect(s, 0, H - band, W, band, cw["title_band"])
    text(s, M, H - band + 0.35, W - 2 * M, 0.9, "PRESENTATION TITLE", DISPLAY,
         F_BLACK, PEARL, align=PP_ALIGN.CENTER)
    text(s, M, H - band + 1.35, W - 2 * M, 0.45,
         "Sub headline  |  Presenter name  |  Month YYYY", SUB, F_SEMI, PEARL,
         align=PP_ALIGN.CENTER)

def s_opening_dark(prs, cw):
    s = new_slide(prs, cw["dark_bg"],
        "Title slide, dark (guide p.35). White-out stacked lock-up with tagline/descriptor. "
        "Master: teal + orange bar. Divisions: deep shade + tint bar.")
    rect(s, 0, H - 0.4, W, 0.4, cw["dark_bar"])
    main_logo(s, 1.0, 1.45, 3.9, 3.8, cw, "Stacked", dark=True)
    text(s, 5.6, 2.35, 7.0, 1.0, "TITLE SLIDE", DISPLAY, F_BLACK, PEARL, anchor=MSO_ANCHOR.BOTTOM)
    text(s, 5.6, 3.45, 7.0, 0.5, "Sub headline goes here", SUB + 4, F_SEMI, PEARL)
    text(s, 5.6, 4.1, 7.0, 0.4, "Presenter name  |  Month YYYY", BODY + 2, F_REG, PEARL)

def s_agenda(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Agenda. Brand-colour panel, knight-only checkerboard top left (text slide).")
    pw = 4.2
    rect(s, 0, 0, pw, H, cw["primary"])
    corner(s, cw, C1, light=cw["ground"])
    text(s, M, 1.5, pw - 1.2, 0.8, "Agenda", X, F_REG, PEARL, bold=True)
    text(s, M, 2.35, pw - 1.2, 0.9, "What we will cover today", SUB, F_SEMI, PEARL)
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
    footer(s, cw, x=x0)

def s_section(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Section divider (guide p.21, p.37, p.48, p.50). Stepped grid, knight + emkay in the corner cell "
        "(at its 40mm print minimum). Master: orange block. Divisions: light tint block.")
    c = H / 4
    for col in range(math.ceil(W / c)):
        rect(s, col * c, 0, min(c, W - col * c), c, cw["ground"], cw["grid"])
    for row in range(1, 4):
        for col in range(2):
            rect(s, col * c, row * c, c, c, cw["ground"], cw["grid"])
    corner(s, cw, c, wordmark=True, light=cw["ground"])
    rect(s, 2 * c, c, W - 2 * c, H - c, cw["section_fill"])
    text(s, 2 * c + 0.6, c + 1.0, W - 2 * c - 1.2, 1.9, ["SECTION", "HEADING"],
         DISPLAY, F_BLACK, cw["section_text"])
    text(s, 2 * c + 0.6, c + 3.0, W - 2 * c - 1.2, 0.5, "Sub line goes here",
         SUB + 2, F_SEMI, cw["section_text"])
    footer(s, cw, x=2 * c + 0.6, color=cw["section_text"])

def s_text_heavy(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Text-heavy slide. Knight-only corner. Title Aptos Bold x, sub SemiBold x/2.5, body Regular sub/1.5.")
    pw = 4.2
    rect(s, 0, 0, pw, H, cw["primary"])
    corner(s, cw, C1, light=cw["ground"])
    text(s, M, 1.5, pw - 1.2, 1.5, ["Slide", "heading"], X, F_REG, PEARL, bold=True)
    text(s, M, 3.1, pw - 1.2, 0.7, "Sub headline in Aptos SemiBold", SUB, F_SEMI, PEARL)
    text(s, M, 3.95, pw - 1.2, 2.5, P3, BODY, F_REG, PEARL)
    x0, gap = pw + 0.7, 0.5
    colw = (W - x0 - M - gap) / 2
    for j in range(2):
        x = x0 + j * (colw + gap)
        text(s, x, 0.75, colw, 0.4, f"Sub-heading {j + 1}", SUB, F_SEMI, cw["primary"])
        text(s, x, 1.3, colw, 5.3, [P1, P2, P3], BODY, F_REG, INK, after=10)
    footer(s, cw, x=x0)

def s_image_content(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Image and content (guide p.36). Knight + emkay corner (image-led slide). Band = 40% of height. "
        "Headline: orange (master) / Pearl White (divisions).")
    ih = H * 0.6
    placeholder(s, 0, 0, W, ih,
        "IMAGE\nEditorial artwork: photo cut-out printed in one colour (halftone) on a flat brand shape.\n"
        "The cut-out breaks one edge of its shape. Never more than two colours.",
        fill=cw["ground"], line=GREY, color=cw["primary"], size=14)
    corner(s, cw, C2, wordmark=True)
    rect(s, 0, ih, W, H - ih, cw["primary"])
    text(s, M, ih + 0.4, 9.5, 0.8, "Headline comes here", X, F_REG,
         cw["headline_on_band"], bold=True)
    text(s, M, ih + 1.35, 9.0, 1.2,
         "Supporting copy in Pearl White on the colour band. One message per slide.",
         SUB, F_REG, PEARL)
    footer(s, cw, x=M, color=PEARL)

def style_chart(chart, cw):
    palette = [cw["primary"], ORANGE, AMBER, GREY]   # fixed order (guide p.36, p.39)
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
    title(s, cw, "Chart title goes here", "Sub headline: the one takeaway from this chart", rule=True)
    rect(s, CX, CY, W - M - CX, 4.0, WHITE)
    data = CategoryChartData()
    data.categories = ["1M", "3M", "6M", "1Y", "3Y"]
    data.add_series("Portfolio", (6.4, 18.1, 26.9, 39.6, 13.2))
    data.add_series("Benchmark", (-3.4, 8.7, 16.7, 36.7, 15.9))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(CX + 0.3), Inches(CY + 0.15),
                            Inches(W - M - CX - 0.6), Inches(3.7), data)
    style_chart(gf.chart, cw)
    text(s, CX, 6.6, 9.0, 0.3, "Sample data. Source and notes line in Aptos Regular, 10 pt.",
         10, F_REG, INK)
    footer(s, cw)

def s_kpi(prs, cw):
    s = new_slide(prs, cw["ground"],
        "KPI slide. Knight-only corner. Growth Orange marks only the ONE figure that matters.")
    corner(s, cw, C1)
    title(s, cw, "Key figures at a glance", "Sub headline: what these numbers mean")
    gap = 0.4
    cwid = (W - M - CX - 2 * gap) / 3
    y, h = CY, 3.2
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
    footer(s, cw)

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
    gf = s.shapes.add_table(len(rows), 4, Inches(CX), Inches(CY),
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
    text(s, CX, 5.75, 9.0, 0.3, "Source and notes line in Aptos Regular, 10 pt.", 10, F_REG, INK)
    footer(s, cw)

def s_comparison(prs, cw):
    s = new_slide(prs, cw["ground"], "Two-column comparison. Knight-only corner.")
    corner(s, cw, C1)
    title(s, cw, "Comparison title goes here", "Sub headline framing the choice")
    gap = 0.5
    cwid = (W - M - CX - gap) / 2
    y, h = CY, 4.1
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
    footer(s, cw)

def s_quote(prs, cw):
    s = new_slide(prs, cw["quote_bg"],
        "Quote. Knight-only corner. One message only. The cut-out portrait breaks the edge of the "
        "accent block (guide p.25). Master: orange block. Divisions: light tint block (p.49, p.51).")
    corner(s, cw, C1, light=cw["ground"])
    text(s, CX, 1.45, 1.5, 1.1, "\u201C", 120, F_BLACK, cw["accent"])
    text(s, CX, 2.55, 6.9, 2.4,
         "One quote per slide, set large enough to read from the back of the room.",
         32, F_REG, PEARL, bold=True)
    text(s, CX, 5.2, 6.9, 0.4, "Speaker Name", SUB, F_SEMI, PEARL)
    text(s, CX, 5.6, 6.9, 0.4, "Designation, Emkay", BODY + 2, F_REG, PEARL)
    rect(s, 8.9, 3.0, 3.7, H - 3.0, cw["accent"])
    placeholder(s, 9.2, 1.5, 3.7, H - 1.5,
        "PORTRAIT CUT-OUT\nBreaks the top edge of the accent block.\nHint of brand colour in clothing or background.",
        fill=None, line=PEARL, color=PEARL, size=11)
    footer(s, cw, color=PEARL, number=False, w=6.9)   # portrait covers the bottom right

def s_team(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Team. Knight + emkay corner (image-led slide). People imagery: natural light, looking out of "
        "frame (guide p.23). Colour OR monochrome per slide, never both (guide p.24).")
    corner(s, cw, C2, wordmark=True)
    title(s, cw, "Leadership team", "One grade per slide: colour or monochrome", c=C2)
    x0, gap = M, 0.35
    pw = (W - M - x0 - 3 * gap) / 4
    y, ph = 3.6, 2.3
    for i in range(4):
        x = x0 + i * (pw + gap)
        placeholder(s, x, y, pw, ph, "PORTRAIT\nNatural light, looking out of frame",
                    fill=WHITE, line=GREY, color=cw["primary"], size=11)
        text(s, x, y + ph + 0.2, pw, 0.4, "Name Surname", SUB, F_SEMI, cw["primary"])
        text(s, x, y + ph + 0.6, pw, 0.4, "Designation", BODY, F_REG, INK)
    footer(s, cw, x=x0)

def s_full_bleed(prs, cw):
    s = new_slide(prs, cw["ground"],
        "Full-bleed image (guide p.26, p.36). When artwork fills the frame, the checkerboard corner "
        "top left is the only place the identity appears. No footer.")
    placeholder(s, 0, 0, W, H,
        "FULL-BLEED IMAGE\nEditorial artwork: one cut-out, one flat brand shape",
        fill=cw["ground"], line=GREY, color=cw["primary"], size=16)
    corner(s, cw, C2, wordmark=True)

def s_thank_you(prs, cw):
    s = new_slide(prs, cw["closing_bg"],
        "Closing slide. Knight-only corner. Master: teal + orange rule. Divisions: deep shade + tint rule "
        "(as the division email signatures, p.53). Carries the legal details, so no footer.")
    corner(s, cw, 1.5)
    text(s, 1.15, 3.9, 9.0, 1.0, "Thank you", 60, F_BLACK, PEARL)
    rect(s, 1.15, 5.05, 6.1, 0.04, cw["accent"])
    text(s, 1.15, 5.35, 10.5, 1.0,
         [cw["entity"].upper(),
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