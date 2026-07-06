"""
ProfitPulse Brief Builder
Target: Paire | Date: 07 Jul 2026
Three slide prospect facing deck. House style per Section 6.0 / 6.0A.
White background body, black header band, amber left stripe, black stat tiles
with a teal top edge. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Georgia"
SANS  = "Calibri"

W = Inches(13.333)
H = Inches(7.5)

DATE_STR = "07 Jul 2026"
COMPANY = "Paire"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    # Strip the theme p:style block entirely, it carries a default effectRef
    # (drop shadow) that some renderers apply even when shadow.inherit is False.
    style_el = shape._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style_el is not None:
        shape._element.remove(style_el)
    return shape


def add_text(slide, text, left, top, width, height, font_name=SANS, font_size=12,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False,
             anchor=None, line_spacing=None, hyperlink=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if anchor:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    if hyperlink:
        box.click_action.hyperlink.address = hyperlink
    return box


def add_multiline(slide, lines, left, top, width, height, font_name=SANS,
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   line_spacing=None, space_after=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        cfg = {"text": line, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False} if isinstance(line, str) else {
            "text": line.get("text", ""), "size": line.get("size", default_size),
            "colour": line.get("colour", default_colour), "bold": line.get("bold", False),
            "italic": line.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        if space_after is not None:
            p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def add_chrome(slide, eyebrow_text):
    """Fixed chrome per Section 6.0A: left stripe, header band, footer line."""
    set_background(slide, WHITE)
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)
    add_rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    add_text(slide, eyebrow_text, Inches(0.45), Inches(0.32), Inches(8.5), Inches(0.4),
              font_name=SANS, font_size=13, bold=True, colour=OFF_WHITE)
    add_text(slide, "PROFITPULSE", Inches(9.5), Inches(0.32), Inches(3.4), Inches(0.4),
              font_name=SANS, font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse",
              Inches(0.45), Inches(7.05), Inches(8), Inches(0.3),
              font_name=SANS, font_size=9, colour=TEAL, italic=True)
    add_text(slide, DATE_STR, Inches(10.5), Inches(7.05), Inches(2.38), Inches(0.3),
              font_name=SANS, font_size=9, colour=TEAL, align=PP_ALIGN.RIGHT, italic=True)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# NOTE: hyperlinks are applied at the shape level (shape.click_action.hyperlink),
# not as a run-level a:hlinkClick. LibreOffice's PDF export forces any run that
# carries a text-level hyperlink to the theme's hlink colour (blue, underlined),
# ignoring explicit run colour. A shape-level click action stays fully clickable
# in the exported PDF while leaving the run's own brand colour untouched.

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
add_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, COMPANY, Inches(0.45), Inches(1.15), Inches(8), Inches(0.75),
          font_name=SERIF, font_size=42, bold=True, colour=BLACK)
add_text(s1, "Direct to consumer apparel and retail brand, South Melbourne VIC",
          Inches(0.45), Inches(1.9), Inches(9.5), Inches(0.35),
          font_name=SANS, font_size=13, colour=TEAL)

# Stat cards row: 5 cards, top y=2.4, height 1.6, width 2.3, gap 0.15
CARD_Y = Inches(2.4)
CARD_W = Inches(2.3)
CARD_H = Inches(1.6)
CARD_GAP = Inches(0.15)
CARD_X0 = Inches(0.45)

stat_cards = [
    ("$10M", "Revenue,\nFY2025", "SmartCompany, Smart50 2025"),
    ("26", "Team, full time\n+ casual", "SmartCompany, Smart50 2025"),
    ("#15", "2025 Smart50\nnational rank", "SmartCompany Smart50 2025"),
    ("2020", "Year\nfounded", "SmartCompany Smart50 profiles"),
    ("10x", "Revenue growth\n2021 to 2024", "SmartCompany growth profile"),
]

for i, (num, label, src) in enumerate(stat_cards):
    x = CARD_X0 + i * (CARD_W + CARD_GAP)
    add_rect(s1, x, CARD_Y, CARD_W, CARD_H, BLACK)
    add_rect(s1, x, CARD_Y, CARD_W, Pt(4), TEAL)
    add_text(s1, num, x + Inches(0.15), CARD_Y + Inches(0.18), CARD_W - Inches(0.3), Inches(0.5),
              font_name=SERIF, font_size=28, bold=True, colour=AMBER_B)
    add_multiline(s1, label.split("\n"), x + Inches(0.15), CARD_Y + Inches(0.72),
                   CARD_W - Inches(0.3), Inches(0.55), font_name=SANS, default_size=11,
                   default_colour=OFF_WHITE, line_spacing=1.0)
    add_text(s1, src, x + Inches(0.15), CARD_Y + Inches(1.32), CARD_W - Inches(0.3), Inches(0.24),
              font_name=SANS, font_size=8, colour=OFF_WHITE, italic=True)

# Key Commercial Signals (left) + Revenue chart (right)
SIG_Y = Inches(4.15)
add_text(s1, "KEY COMMERCIAL SIGNALS", CARD_X0, SIG_Y, Inches(7.6), Inches(0.3),
          font_name=SANS, font_size=12, bold=True, colour=TEAL)

signals = [
    "Ranked 15th nationally at the 2025 Smart50 Awards, up from 13th in 2024, 38th in 2023",
    "Won the Rising Star and Retail Award categories at the 2024 Smart50 Awards",
    "Opened first physical flagship at QV Melbourne in January 2025, after years online only",
    "Plans two to three more Sydney stores plus additional Melbourne sites within 18 months",
    "Pitched on Shark Tank Australia in 2024, seeking $500,000 for 2.5 percent equity",
    "Expanded product sales into Singapore, its first international market",
]
sig_lines = [{"text": "•  " + s, "size": 11.5} for s in signals]
add_multiline(s1, sig_lines, CARD_X0, SIG_Y + Inches(0.35), Inches(7.6), Inches(2.4),
               font_name=SANS, default_colour=BLACK, space_after=8,
               line_spacing=1.05)

# Revenue chart panel, right side
CHART_X = Inches(8.35)
CHART_W = Inches(4.55)
add_text(s1, "REVENUE, VERIFIED", CHART_X, SIG_Y, CHART_W, Inches(0.3),
          font_name=SANS, font_size=12, bold=True, colour=TEAL)

BASE_Y = 6.35
bar_w = Inches(0.9)
tall_h = Inches(1.7)
short_h = Inches(0.17)
bar1_x = CHART_X + Inches(0.6)
bar2_x = CHART_X + Inches(2.6)

add_rect(s1, bar1_x, Inches(BASE_Y) - short_h, bar_w, short_h, TEAL)
add_text(s1, "$1M", bar1_x - Inches(0.15), Inches(BASE_Y) - short_h - Inches(0.32),
          Inches(1.2), Inches(0.3), font_name=SERIF, font_size=13, bold=True, colour=BLACK)
add_text(s1, "2021", bar1_x - Inches(0.15), Inches(BASE_Y) + Inches(0.05), Inches(1.2), Inches(0.25),
          font_name=SANS, font_size=10, colour=TEAL)

add_rect(s1, bar2_x, Inches(BASE_Y) - tall_h, bar_w, tall_h, AMBER_B)
add_text(s1, "$10M", bar2_x - Inches(0.15), Inches(BASE_Y) - tall_h - Inches(0.32),
          Inches(1.3), Inches(0.3), font_name=SERIF, font_size=13, bold=True, colour=BLACK)
add_text(s1, "2024", bar2_x - Inches(0.15), Inches(BASE_Y) + Inches(0.05), Inches(1.2), Inches(0.25),
          font_name=SANS, font_size=10, colour=TEAL)

add_text(s1, "Source: SmartCompany, growth profile", CHART_X, Inches(6.75), CHART_W, Inches(0.25),
          font_name=SANS, font_size=8, colour=TEAL, italic=True)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
add_chrome(s2, "THE OPPORTUNITY")

add_text(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
          Inches(0.45), Inches(1.15), Inches(11.5), Inches(0.4),
          font_name=SERIF, font_size=18, bold=True, colour=BLACK)

COL_Y = Inches(1.65)
COL_H = Inches(4.5)
COL_W = Inches(4.411)

observations = [
    ("01", "The online engine outgrew its plan",
     "Revenue moved from one million to ten million between 2021 and 2024, with the Smart50 "
     "rank climbing from 38th to 13th to 15th. That pace rarely arrives paired with the multi "
     "year financial model a physical rollout now needs before it, not after.",
     TEAL, BLACK, BLACK),
    ("02", "One store is a pilot, not a program",
     "The QV Melbourne flagship opened January 2025. Two to three Sydney stores and more "
     "Melbourne sites are planned within eighteen months. Each carries its own fit out, stock, "
     "and staffing cost. A rollout needs a shared model ranking each site's return before the "
     "lease is signed.",
     BLACK, WHITE, OFF_WHITE),
    ("03", "Capital is already part of the story",
     "The 2024 Shark Tank pitch sought $500,000 for 2.5 percent equity alongside a new "
     "Singapore line. That confirms growth capital is already being weighed, on a lean sixteen "
     "person team. A Growth Diagnostic ranks trading cashflow, debt, and equity against each "
     "store's expected return.",
     GOLD, BLACK, BLACK),
]

for i, (idx, header, body, fill, textcol, indexcol) in enumerate(observations):
    x = Inches(0.1) + i * COL_W
    add_rect(s2, x, COL_Y, COL_W, COL_H, fill)
    pad = Inches(0.3)
    add_text(s2, idx, x + pad, COL_Y + Inches(0.25), COL_W - 2 * pad, Inches(0.7),
              font_name=SERIF, font_size=40, bold=True, colour=indexcol)
    add_text(s2, header, x + pad, COL_Y + Inches(1.0), COL_W - 2 * pad, Inches(0.75),
              font_name=SERIF, font_size=16, bold=True, colour=textcol)
    add_text(s2, body, x + pad, COL_Y + Inches(1.85), COL_W - 2 * pad, Inches(2.5),
              font_name=SANS, font_size=11.5, colour=textcol, line_spacing=1.15)

add_text(s2, "These observations are offered in good faith. Paire has built something genuinely "
              "fast growing. The question is simply whether the financial architecture keeps pace "
              "with the ambition.",
          Inches(0.45), Inches(6.3), Inches(12.4), Inches(0.55),
          font_name=SANS, font_size=11, italic=True, colour=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ════════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
add_chrome(s3, "THE RECOMMENDATION")

LEFT_X = Inches(0.45)
LEFT_W = Inches(7.1)

add_text(s3, "Strategic Growth Diagnostic", LEFT_X, Inches(1.2), LEFT_W, Inches(0.55),
          font_name=SERIF, font_size=27, bold=True, colour=BLACK)
add_text(s3, "$5,000 one off", LEFT_X, Inches(1.78), LEFT_W, Inches(0.4),
          font_name=SANS, font_size=17, bold=True, colour=AMBER_D)
add_text(s3, "Maps revenue, capacity and margin headroom, then builds a twelve month growth "
              "plan with funding and capital steps for each new site.",
          LEFT_X, Inches(2.25), LEFT_W, Inches(0.65),
          font_name=SANS, font_size=12.5, colour=BLACK, line_spacing=1.15)

# Step one block
add_rect(s3, LEFT_X, Inches(3.05), LEFT_W, Inches(1.55), WHITE,
          line_colour=TEAL, line_width=Pt(1.25))
add_text(s3, "Step one, answer a few quick questions", LEFT_X + Inches(0.25), Inches(3.22),
          LEFT_W - Inches(0.5), Inches(0.35), font_name=SANS, font_size=13, bold=True, colour=TEAL)
add_text(s3, "See the solutions matched to your size and industry.",
          LEFT_X + Inches(0.25), Inches(3.6), LEFT_W - Inches(0.5), Inches(0.35),
          font_name=SANS, font_size=11.5, colour=BLACK)
add_text(s3, "profit-pulse.com.au/services/find-your-fit", LEFT_X + Inches(0.25), Inches(4.0),
          LEFT_W - Inches(0.5), Inches(0.45), font_name=SANS, font_size=13, bold=True,
          colour=TEAL, hyperlink="https://profit-pulse.com.au/services/find-your-fit/")

# Direct CTA
cta_box = add_rect(s3, LEFT_X, Inches(4.85), Inches(5.6), Inches(0.62), AMBER_D)
tf = cta_box.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "Purchase the suggested product now to get started"
run.font.name = SANS
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = BLACK
cta_box.click_action.hyperlink.address = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"

add_text(s3, "Prefer a conversation first?", LEFT_X, Inches(5.75), LEFT_W, Inches(0.32),
          font_name=SANS, font_size=11.5, colour=BLACK)
add_text(s3, "Book a complimentary discovery call", LEFT_X, Inches(6.08), LEFT_W, Inches(0.35),
          font_name=SANS, font_size=12, bold=True, colour=TEAL,
          hyperlink="https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right column: About ProfitPulse and Nitesh, black credibility panel
PANEL_X = Inches(7.95)
PANEL_W = Inches(4.9)
add_rect(s3, PANEL_X, Inches(1.2), PANEL_W, Inches(5.55), BLACK)
add_rect(s3, PANEL_X, Inches(1.2), PANEL_W, Pt(4), TEAL)

add_text(s3, "NITESH ROOPA", PANEL_X + Inches(0.3), Inches(1.5), PANEL_W - Inches(0.6), Inches(0.45),
          font_name=SERIF, font_size=19, bold=True, colour=AMBER_B)
add_text(s3, "CA, Managing Partner, ProfitPulse", PANEL_X + Inches(0.3), Inches(1.98),
          PANEL_W - Inches(0.6), Inches(0.35), font_name=SANS, font_size=12.5, colour=WHITE)

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
add_multiline(s3, [{"text": "•  " + c, "size": 11.5} for c in cred_lines],
               PANEL_X + Inches(0.3), Inches(2.45), PANEL_W - Inches(0.6), Inches(1.6),
               font_name=SANS, default_colour=OFF_WHITE, space_after=6, line_spacing=1.1)

add_rect(s3, PANEL_X + Inches(0.3), Inches(4.2), PANEL_W - Inches(0.6), Pt(1.25), TEAL)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 12},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 12},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 11},
]
add_multiline(s3, contact_lines, PANEL_X + Inches(0.3), Inches(4.4), PANEL_W - Inches(0.6),
               Inches(1.5), font_name=SANS, default_colour=OFF_WHITE, space_after=6)

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_07Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
