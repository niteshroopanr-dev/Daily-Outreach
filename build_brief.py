"""
ProfitPulse Brief Builder, v3.3 house style.
Target: AH Fencing | Date: 11 Jul 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
White body background, black header band, amber left stripe, teal accented
black stat cards, muted footer line, per Section 6 of the nightly instructions.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours, seven only
BLACK = RGBColor(0x00, 0x00, 0x00)
TEAL = RGBColor(0x01, 0xA2, 0x96)
AMBER_B = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D = RGBColor(0xF6, 0xA1, 0x02)
GOLD = RGBColor(0xE3, 0xA7, 0x12)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STAMP = "11 Jul 2026"
COMPANY = "AH Fencing"
STRIPE_LINK = "https://buy.stripe.com/aFa7sMedigKQ81ibEz3ks0i"
QUEST_URL_CLEAN = "https://profit-pulse.com.au/services/find-your-fit/"
BOOKING_LINK = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    shape.shadow.inherit = False
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    # Strip the theme p:style block entirely, its effectRef renders a shadow
    # in LibreOffice even when spPr carries an empty effectLst override.
    style_el = shape._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style_el is not None:
        shape._element.remove(style_el)
    return shape


def add_text(slide, text, left, top, width, height, font_name="Calibri", font_size=14,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False,
             anchor=MSO_ANCHOR.TOP, wrap=True, heading=False, hyperlink=None,
             shrink=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = "Georgia" if heading else font_name
    f.size = Pt(font_size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = colour
    if hyperlink:
        f.underline = False
        run.hyperlink.address = hyperlink
    return box


def add_multiline(slide, lines, left, top, width, height, font_name="Calibri",
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   line_spacing=1.15, space_after=4, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        cfg = line if isinstance(line, dict) else {"text": line}
        cfg.setdefault("size", default_size)
        cfg.setdefault("colour", default_colour)
        cfg.setdefault("bold", False)
        cfg.setdefault("italic", False)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = cfg["text"]
        f = run.font
        f.name = font_name
        f.size = Pt(cfg["size"])
        f.bold = cfg["bold"]
        f.italic = cfg["italic"]
        f.color.rgb = cfg["colour"]
    return box


def add_chrome(slide, eyebrow, right_label):
    set_background(slide, WHITE)
    add_rect(slide, 0, 0, Inches(0.1), H, AMBER_B)
    add_rect(slide, 0, 0, W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.34), Inches(8.0), Inches(0.4),
             font_size=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, right_label, Inches(6.0), Inches(0.30), Inches(7.0), Inches(0.5),
             font_size=13, bold=True, colour=WHITE, align=PP_ALIGN.RIGHT)
    add_text(slide, f"Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au        {DATE_STAMP}",
             Inches(0.35), Inches(7.08), Inches(12.6), Inches(0.3), font_size=8, bold=False,
             colour=RGBColor(0x77, 0x77, 0x77), align=PP_ALIGN.LEFT)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    add_rect(slide, left, top, width, height, BLACK)
    add_rect(slide, left, top, width, Pt(3), TEAL)
    add_text(slide, number, left + Inches(0.15), top + Inches(0.16), width - Inches(0.3), Inches(0.55),
              font_size=28, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_multiline(slide, label_lines, left + Inches(0.15), top + Inches(0.74), width - Inches(0.3), Inches(0.5),
                  font_name="Calibri", default_size=11, default_colour=OFF_WHITE, line_spacing=1.0, space_after=0)
    add_text(slide, source, left + Inches(0.15), top + height - Inches(0.28), width - Inches(0.3), Inches(0.24),
              font_size=7, bold=False, colour=RGBColor(0x9a, 0x9a, 0x9a), align=PP_ALIGN.LEFT, italic=True)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ═══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
add_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

add_text(s1, "AH Fencing", Inches(0.35), Inches(1.18), Inches(9.0), Inches(0.75),
         font_size=42, bold=True, colour=BLACK, align=PP_ALIGN.LEFT, heading=True)
add_text(s1, "Commercial and security fencing contractor, Northgate, Brisbane QLD",
         Inches(0.35), Inches(1.95), Inches(11.0), Inches(0.35),
         font_size=13, bold=False, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)

cards = [
    ("$16.2M", ["FY2025", "revenue"], "SmartCompany, 13 Nov 2025"),
    ("55%", ["Revenue growth,", "year on year"], "SmartCompany, 13 Nov 2025"),
    ("50", ["Team", "members"], "Smart50 2025 citation"),
    ("2017", ["Company", "founded"], "SmartCompany, 13 Nov 2025"),
    ("#20", ["Smart50 2025", "national rank"], "SmartCompany, 13 Nov 2025"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
x = Inches(0.35)
for number, label_lines, source in cards:
    stat_card(s1, x, Inches(2.45), card_w, Inches(1.6), number, label_lines, source)
    x = Inches(x.inches + card_w.inches + gap.inches)

add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.28), Inches(8.0), Inches(0.3),
         font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    "Revenue grew $3.2M to $16.2M, FY21 to FY25, 55% YoY. SmartCompany, 13 Nov 2025.",
    "Team grew to 50 staff, up from near 40 in 2023. SmartCompany, 13 Nov 2025.",
    "Melbourne office opened 2023, New South Wales entry planned 2026. SmartCompany, Nov 2025.",
    "Aiming to trade in every state and territory by 2028. SmartCompany, 13 Nov 2025.",
    "New Brisbane headquarters warehouse opened 2023. SmartCompany, 5 Apr 2023.",
    "Industrial Commercial Fencing Award winner. Fencing Industry Australia.",
]
add_multiline(s1, [{"text": f"•  {t}", "size": 12, "colour": RGBColor(0x1a, 0x1a, 0x1a)} for t in signals],
              Inches(0.35), Inches(4.62), Inches(12.4), Inches(2.2),
              default_size=12, line_spacing=1.25, space_after=8)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ═══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
add_chrome(s2, "THE OPPORTUNITY", "AH FENCING, THREE COMMERCIAL OBSERVATIONS")

col_top = Inches(1.15)
col_h = Inches(5.35)
col_w = Inches(4.411)
col_positions = [Inches(0.1), Inches(4.511), Inches(8.922)]
col_fills = [TEAL, BLACK, GOLD]
col_text = [BLACK, WHITE, BLACK]

observations = [
    ("01", "Three states, one growth plan?",
     "Revenue moved from 3.2 million in FY21 to 16.2 million in FY25, and the next "
     "step is New South Wales in 2026, on the way to national coverage by 2028. Each "
     "new state carries setup cost, working capital lag and a hiring curve. A costed "
     "growth plan puts numbers behind that before the decision is made, not after."),
    ("02", "Fifty people, one finance seat?",
     "AH Fencing has built a genuine leadership bench, named sales, estimating and "
     "site leads under two managing directors. What is not visible publicly is a "
     "dedicated finance seat. A business growing 55 percent a year needs a permanent "
     "owner of the numbers, not a periodic look, with a third state about to add its "
     "own overhead."),
    ("03", "Sixteen million, one capital story?",
     "Every dollar going into a new warehouse, a new state's setup cost or working "
     "capital for the New South Wales launch is a capital allocation decision. At this "
     "pace it is worth knowing, dollar for dollar, whether that capital earns the "
     "return the fencing contracts themselves generate, before the next site is "
     "committed to."),
]

for (idx, header, para), x, fill, txt_colour in zip(observations, col_positions, col_fills, col_text):
    add_rect(s2, x, col_top, col_w, col_h, fill)
    inset = Inches(0.32)
    inner_w = Inches(col_w.inches - 0.64)
    add_text(s2, idx, x + inset, col_top + Inches(1.05), inner_w, Inches(0.65),
              font_size=32, bold=True, colour=txt_colour, align=PP_ALIGN.LEFT, heading=True)
    add_text(s2, header, x + inset, col_top + Inches(1.78), inner_w, Inches(0.8),
              font_size=16, bold=True, colour=txt_colour, align=PP_ALIGN.LEFT)
    add_multiline(s2, [{"text": para, "size": 11.5, "colour": txt_colour}],
                  x + inset, col_top + Inches(2.68), inner_w, Inches(2.5),
                  default_size=11.5, line_spacing=1.2, space_after=0)

add_text(s2,
         "These are observations offered in good faith. The company has built something "
         "impressive. The question is simply whether the financial architecture matches "
         "the ambition.",
         Inches(0.35), Inches(6.58), Inches(12.6), Inches(0.4),
         font_size=11, italic=True, bold=False, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ═══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
add_chrome(s3, "THE RECOMMENDATION", "AH FENCING")

# LEFT COLUMN
lx = Inches(0.35)
lw = Inches(7.3)

add_text(s3, "Strategic Growth Diagnostic", lx, Inches(1.2), lw, Inches(0.55),
         font_size=24, bold=True, colour=BLACK, align=PP_ALIGN.LEFT, heading=True)
add_text(s3, "$7,000 one off, ProfitPulse verified price", lx, Inches(1.78), lw, Inches(0.4),
         font_size=15, bold=True, colour=AMBER_D, align=PP_ALIGN.LEFT)
add_multiline(s3, [
    {"text": "Maps revenue, capacity and margin headroom into a costed twelve month growth "
             "plan, useful timing given the New South Wales entry planned for 2026.",
     "size": 12.5, "colour": RGBColor(0x1a, 0x1a, 0x1a)}],
    lx, Inches(2.22), lw, Inches(0.75), default_size=12.5, line_spacing=1.2)

add_rect(s3, lx, Inches(3.05), lw, Inches(1.55), RGBColor(0xEF, 0xF8, 0xF7), line_colour=TEAL, line_width=Pt(1.25))
add_text(s3, "Step one, answer a few quick questions", lx + Inches(0.25), Inches(3.2), lw - Inches(0.5), Inches(0.35),
         font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry.",
         lx + Inches(0.25), Inches(3.58), lw - Inches(0.5), Inches(0.35),
         font_size=11.5, bold=False, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)
add_text(s3, "profit-pulse.com.au/services/find-your-fit",
         lx + Inches(0.25), Inches(3.98), lw - Inches(0.5), Inches(0.4),
         font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.LEFT,
         hyperlink=QUEST_URL_CLEAN)

add_rect(s3, lx, Inches(4.78), lw, Inches(0.6), WHITE, line_colour=AMBER_D, line_width=Pt(1.5))
add_text(s3, "Purchase the suggested product now to get started",
         lx, Inches(4.78), lw, Inches(0.6), font_size=13, bold=True, colour=AMBER_D,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, hyperlink=STRIPE_LINK)

add_text(s3, "Prefer a conversation first?", lx, Inches(5.62), lw, Inches(0.35),
         font_size=12, bold=False, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)
add_text(s3, "Book a complimentary discovery call", lx, Inches(5.97), lw, Inches(0.35),
         font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.LEFT, hyperlink=BOOKING_LINK)

# RIGHT COLUMN
rx = Inches(8.0)
rw = Inches(4.98)
add_rect(s3, rx, Inches(1.2), rw, Inches(5.35), BLACK)
add_rect(s3, rx, Inches(1.2), rw, Pt(3), TEAL)

add_text(s3, "Nitesh Roopa", rx + Inches(0.3), Inches(1.42), rw - Inches(0.6), Inches(0.45),
         font_size=19, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT, heading=True)
add_text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.3), Inches(1.9), rw - Inches(0.6), Inches(0.35),
         font_size=12.5, bold=False, colour=WHITE, align=PP_ALIGN.LEFT)
add_rect(s3, rx + Inches(0.3), Inches(2.32), rw - Inches(0.6), Pt(1.25), TEAL)

cred_lines = [
    {"text": "16 years across 4 countries", "size": 11.5, "colour": OFF_WHITE},
    {"text": "52 deals executed and managed", "size": 11.5, "colour": OFF_WHITE},
    {"text": "Largest single deal, USD 1.3 billion, Cahora Bassa", "size": 11.5, "colour": OFF_WHITE},
    {"text": "Total GRBT project value over AUD 10 billion", "size": 11.5, "colour": OFF_WHITE},
]
add_multiline(s3, cred_lines, rx + Inches(0.3), Inches(2.5), rw - Inches(0.6), Inches(1.55),
              default_size=11.5, line_spacing=1.25, space_after=7)

add_rect(s3, rx + Inches(0.3), Inches(4.12), rw - Inches(0.6), Pt(1.25), TEAL)
contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 12, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 12, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 11, "colour": OFF_WHITE},
]
add_multiline(s3, contact_lines, rx + Inches(0.3), Inches(4.3), rw - Inches(0.6), Inches(1.5),
              default_size=12, line_spacing=1.3, space_after=6)

# Retheme hyperlink colours to brand teal. Left uncorrected, both PowerPoint
# and LibreOffice render hlinkClick runs in the theme's default link blue
# regardless of an explicit run colour, which would break Rule 3.
from lxml import etree as _etree
_master = prs.slide_masters[0]
_theme_part = _master.part.part_related_by(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
_root = _etree.fromstring(_theme_part.blob)
_ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
_root.find(".//a:clrScheme/a:hlink/a:srgbClr", _ns).set("val", "01A296")
_root.find(".//a:clrScheme/a:folHlink/a:srgbClr", _ns).set("val", "01A296")
_theme_part._blob = _etree.tostring(_root, xml_declaration=True, encoding="UTF-8", standalone=True)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_AH_Fencing_11Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
