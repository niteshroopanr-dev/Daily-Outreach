"""
ProfitPulse Brief Builder, Version 3.3 house style
Target: Equality Media + Marketing | Date: 09 Jul 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Brand colours, the only seven permitted anywhere in this file
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
SERIF = "Cambria"
SANS = "Calibri"

DATE_STAMP = "09 Jul 2026"
COMPANY = "Equality Media + Marketing"
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_URL = "https://buy.stripe.com/3cI4gA1qw5280yQgYT3ks0w"
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def strip_style(shape):
    # Remove the theme <p:style> reference (line/fill/effect/font refs to
    # accent colours) so nothing but our explicit direct formatting renders.
    # LibreOffice applies the style's effectRef (a drop shadow) even when
    # spPr carries an empty <a:effectLst/>, which would blend non brand
    # colours into the edges of every shape.
    el = shape._element
    style = el.find("{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style is not None:
        el.remove(style)


def add_rect(slide, left, top, width, height, fill_colour, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    shape.line.fill.background()
    shape.shadow.inherit = False
    strip_style(shape)
    return shape


def add_text(slide, text, left, top, width, height, font_size=12, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, font_name=SANS, italic=False,
             anchor=None, line_spacing=None):
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
    return box, run


def add_multiline(slide, lines, left, top, width, height, font_name=SANS,
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   line_spacing=1.15, space_after=2):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size, "colour": default_colour, "bold": False}
        else:
            cfg = {"text": line.get("text", ""), "size": line.get("size", default_size),
                   "colour": line.get("colour", default_colour), "bold": line.get("bold", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.color.rgb = cfg["colour"]
    return box


def add_button(slide, text, left, top, width, height, fill_colour, text_colour, url,
               font_size=13, outline=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.5
    if outline:
        shape.fill.background()
        shape.line.color.rgb = fill_colour
        shape.line.width = Pt(1.5)
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_colour
        shape.line.fill.background()
    shape.shadow.inherit = False
    strip_style(shape)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = SANS
    run.font.size = Pt(font_size)
    run.font.bold = True
    run.font.color.rgb = text_colour
    shape.click_action.hyperlink.address = url
    return shape


def header_band(slide, eyebrow):
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8), Inches(0.4),
              font_size=13, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, "PROFITPULSE", Inches(8.5), Inches(0.32), Inches(4.483), Inches(0.4),
              font_size=13, bold=True, colour=AMBER_B, align=PP_ALIGN.RIGHT,
              anchor=MSO_ANCHOR.MIDDLE)


def left_stripe(slide):
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)


def footer(slide):
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.35), Inches(7.05), Inches(9.5), Inches(0.3), font_size=9, colour=BLACK,
              align=PP_ALIGN.LEFT)
    add_text(slide, DATE_STAMP, Inches(10.0), Inches(7.05), Inches(2.98), Inches(0.3),
              font_size=9, colour=BLACK, align=PP_ALIGN.RIGHT)


def add_hyperlink_run(slide, text, left, top, width, height, url, font_size=12,
                       colour=TEAL, bold=True, align=PP_ALIGN.LEFT):
    # Hyperlink is attached at shape level (click_action), not run level.
    # LibreOffice's Impress renderer forces run-level hyperlinks to a
    # non-brand blue regardless of explicit run colour; shape-level click
    # actions do not trigger that recolour, so the brand teal survives export.
    box, run = add_text(slide, text, left, top, width, height, font_size=font_size,
                         bold=bold, colour=colour, align=align)
    run.font.underline = True
    box.click_action.hyperlink.address = url
    return box


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
header_band(s1, "COMMERCIAL INTELLIGENCE BRIEF")
left_stripe(s1)

add_text(s1, COMPANY, Inches(0.35), Inches(1.15), Inches(11.5), Inches(0.7),
          font_size=40, bold=True, colour=BLACK, font_name=SERIF)
add_text(s1, "Independent media and marketing agency, Richmond, Melbourne VIC",
          Inches(0.35), Inches(1.9), Inches(11.5), Inches(0.4), font_size=14, colour=BLACK)

# Stat cards
cards = [
    ("$13M", ["Annual revenue,", "Smart50 rank 24"], "SmartCompany Smart50 2025"),
    ("49%",  ["Three year average", "revenue growth"], "SmartCompany Smart50 2025"),
    ("30",   ["People on the", "Richmond team"], "AFR BOSS coverage, 2025"),
    ("2018", ["Year founded by", "Marilla Akkermans"], "Company and press records"),
    ("#1",   ["Best Place to Work", "in Media, 2025"], "AdNews, April 2025"),
    ("$250K",["Donated to Women's", "Property Initiatives"], "B&T, April 2025"),
]
card_w = Inches(1.98)
gap = Inches(0.15)
card_top = Inches(2.4)
card_h = Inches(1.6)
x = Inches(0.35)
for num, label_lines, source in cards:
    add_rect(s1, x, card_top, card_w, card_h, BLACK)
    add_rect(s1, x, card_top, card_w, Pt(4), TEAL)
    add_text(s1, num, x + Inches(0.12), card_top + Inches(0.12), card_w - Inches(0.24), Inches(0.5),
              font_size=28, bold=True, colour=AMBER_B, font_name=SERIF)
    add_multiline(s1, label_lines, x + Inches(0.12), card_top + Inches(0.62), card_w - Inches(0.24), Inches(0.55),
                  default_size=11, default_colour=OFF_WHITE, line_spacing=1.0, space_after=0)
    add_text(s1, source, x + Inches(0.12), card_top + Inches(1.28), card_w - Inches(0.24), Inches(0.25),
              font_size=8, colour=OFF_WHITE, italic=True)
    x += card_w + gap

# Key commercial signals
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.18), Inches(8), Inches(0.3),
          font_size=13, bold=True, colour=TEAL)
signals = [
    "Team grew from 15 to 30 across 2024 to 25 while the studio footprint expanded (Smart50; AdNews)",
    "Named number one AFR BOSS Best Place to Work in Media, third year running (AdNews, April 2025)",
    "Runs a four day, full pay 32 hour week with 19 extra leave days across the team (B&T, 2025)",
    "Donates 10 percent of annual profit to Women's Property Initiatives, $250,000 in 2025 (B&T)",
    "Won the BIODERMA ANZ media account, full funnel strategy and channel remit (AdNews, March 2026)",
    "Great Place to Work certified, single studio location in Richmond VIC (Great Place to Work AU)",
]
add_multiline(s1, [{"text": "• " + s, "size": 12, "colour": BLACK} for s in signals],
              Inches(0.35), Inches(4.52), Inches(12.6), Inches(2.3),
              line_spacing=1.05, space_after=6)
footer(s1)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
header_band(s2, "THE OPPORTUNITY")
left_stripe(s2)

add_text(s2, f"{COMPANY}: Three commercial observations from ProfitPulse",
          Inches(0.35), Inches(1.12), Inches(12.6), Inches(0.45), font_size=16, bold=True, colour=TEAL)

col_top = Inches(1.65)
col_h = Inches(3.95)
col_w = Inches(4.131)
col_gap = Inches(0.12)
cols = [
    (TEAL, BLACK, "01", "Growth has outpaced the finance rhythm",
     "The team has roughly doubled in the past year, from 15 to 30 people, while the studio kept its four "
     "day week at full pay and an expanded leave framework in place. Payroll now moves ahead of revenue "
     "recognition on retainer and campaign work, and without a rolling forecast the exact week that gap "
     "bites is hard to see in advance."),
    (BLACK, WHITE, "02", "Self funded growth raises the stakes on cash timing",
     "No funding round or facility has been publicly reported alongside this expansion, which points to "
     "growth funded from trading cash flow. That is a strong result, but it also means every hire and every "
     "new client onboarding draws on the same pool of cash the business needs for its own payroll and "
     "commitments."),
    (GOLD, BLACK, "03", "The culture commitments are a genuine cost line",
     "The four day week, the extended leave framework, and the $250,000 given to Women's Property "
     "Initiatives in 2025 are a deliberate cost structure, not an incidental one. Protecting it through "
     "further growth depends on knowing, week by week, what the business can actually afford to commit."),
]
x = Inches(0.35)
for fill, text_colour, idx, header, para in cols:
    add_rect(s2, x, col_top, col_w, col_h, fill)
    add_text(s2, idx, x + Inches(0.25), col_top + Inches(0.15), col_w - Inches(0.5), Inches(0.7),
              font_size=40, bold=True, colour=text_colour, font_name=SERIF)
    add_text(s2, header, x + Inches(0.25), col_top + Inches(0.95), col_w - Inches(0.5), Inches(0.7),
              font_size=16, bold=True, colour=text_colour, font_name=SERIF)
    add_text(s2, para, x + Inches(0.25), col_top + Inches(1.65), col_w - Inches(0.5), Inches(2.85),
              font_size=11, colour=text_colour, line_spacing=1.12)
    x += col_w + col_gap

add_text(s2, "These are observations offered in good faith. The agency has built a workplace people want to "
             "be part of. The question is simply whether the cash rhythm keeps pace with the growth and the "
             "culture it is funding.",
          Inches(0.35), Inches(5.95), Inches(12.6), Inches(0.6), font_size=12, italic=True, colour=BLACK,
          line_spacing=1.15)
footer(s2)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
header_band(s3, "THE RECOMMENDATION")
left_stripe(s3)

# Divider
add_rect(s3, Inches(8.15), Inches(1.15), Pt(1.5), Inches(5.75), TEAL)

# Left column
add_text(s3, "13 Week Cash Flow Build", Inches(0.35), Inches(1.2), Inches(7.6), Inches(0.55),
          font_size=24, bold=True, colour=BLACK, font_name=SERIF)
add_text(s3, "$2,650 one off", Inches(0.35), Inches(1.78), Inches(7.6), Inches(0.4),
          font_size=18, bold=True, colour=AMBER_D)
add_text(s3, "A rolling 13 week cash flow forecast with three scenarios, so payroll timing is planned, not "
             "guessed, while the team keeps growing.",
          Inches(0.35), Inches(2.24), Inches(7.6), Inches(0.6), font_size=12, colour=BLACK, line_spacing=1.1)

add_rect(s3, Inches(0.35), Inches(3.0), Inches(7.6), Inches(1.5), WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
step_box = s3.shapes[-1]
step_box.line.color.rgb = TEAL
step_box.line.width = Pt(1.25)
step_box.fill.solid()
step_box.fill.fore_color.rgb = WHITE

add_text(s3, "Step one: answer a few quick questions", Inches(0.55), Inches(3.12), Inches(7.2), Inches(0.35),
          font_size=13, bold=True, colour=BLACK)
add_text(s3, "See the solutions matched to your size and industry.",
          Inches(0.55), Inches(3.5), Inches(7.2), Inches(0.35), font_size=11, colour=BLACK)
add_hyperlink_run(s3, QUESTIONNAIRE_CLEAN, Inches(0.55), Inches(3.9), Inches(7.2), Inches(0.4),
                   QUESTIONNAIRE_URL, font_size=13, colour=TEAL)

add_button(s3, "Purchase the suggested product now to get started",
           Inches(0.55), Inches(4.85), Inches(4.9), Inches(0.6), AMBER_D, WHITE, STRIPE_URL, font_size=12)

add_text(s3, "Prefer a conversation first?", Inches(0.35), Inches(5.75), Inches(7.6), Inches(0.35),
          font_size=11, colour=BLACK)
add_hyperlink_run(s3, "Book a complimentary discovery call", Inches(0.35), Inches(6.1), Inches(7.6), Inches(0.4),
                   BOOKING_URL, font_size=12, colour=TEAL)

# Right panel
panel_x = Inches(8.3)
add_rect(s3, panel_x, Inches(1.15), Inches(4.683), Inches(5.6), BLACK)
add_text(s3, "NITESH ROOPA", panel_x + Inches(0.25), Inches(1.35), Inches(4.2), Inches(0.5),
          font_size=19, bold=True, colour=AMBER_B, font_name=SERIF)
add_text(s3, "CA, Managing Partner, ProfitPulse", panel_x + Inches(0.25), Inches(1.85), Inches(4.2), Inches(0.35),
          font_size=13, colour=WHITE)
add_rect(s3, panel_x + Inches(0.25), Inches(2.28), Inches(4.0), Pt(1.5), TEAL)

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
add_multiline(s3, [{"text": "• " + c, "size": 11, "colour": OFF_WHITE} for c in cred_lines],
              panel_x + Inches(0.25), Inches(2.48), Inches(4.2), Inches(1.7), line_spacing=1.2, space_after=6)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 12, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 11, "colour": TEAL},
]
add_multiline(s3, contact_lines, panel_x + Inches(0.25), Inches(4.35), Inches(4.2), Inches(1.6),
              line_spacing=1.3, space_after=4)

footer(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMedia_09Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
