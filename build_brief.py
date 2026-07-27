"""
ProfitPulse Brief Builder
Target: MuraConnect | Date: 28 Jul 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
Built to the Section 6 house style: black header band, amber left stripe,
white body, black stat tiles with teal top edge, three colour observation
columns, two column recommendation and credibility slide.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Brand colours, the only seven permitted anywhere in this deck
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Georgia"
SANS = "Arial"

W = Inches(13.333)
H = Inches(7.5)

COMPANY = "MuraConnect"
DATE_STAMP = "28 Jul 2026"
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_URL = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, font_name=SANS, font_size=14,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False,
             hyperlink=None, underline=False, anchor=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = font_name
    f.size = Pt(font_size)
    f.bold = bold
    f.italic = italic
    f.underline = underline
    f.color.rgb = colour
    if hyperlink:
        run.hyperlink.address = hyperlink
    return box


def add_lines(slide, lines, left, top, width, height, font_name=SANS,
              default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
              line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        cfg = {"text": "", "size": default_size, "colour": default_colour,
               "bold": False, "italic": False}
        if isinstance(line, str):
            cfg["text"] = line
        else:
            cfg.update(line)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def add_header(slide, eyebrow):
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.3), Inches(0.32), Inches(8.5), Inches(0.4),
              font_size=13, bold=True, colour=OFF_WHITE)
    add_text(slide, "PROFITPULSE", Inches(9.5), Inches(0.32), Inches(3.5), Inches(0.4),
              font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)


def add_stripe(slide):
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)


def add_footer(slide):
    add_rect(slide, Inches(0.3), Inches(7.0), Inches(12.7), Pt(1), BLACK)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.3), Inches(7.08), Inches(9.5), Inches(0.3), font_size=9, colour=BLACK)
    add_text(slide, DATE_STAMP, Inches(10.0), Inches(7.08), Inches(3.0), Inches(0.3),
              font_size=9, colour=BLACK, align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ============================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ============================================================
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
add_header(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, COMPANY, Inches(0.3), Inches(1.12), Inches(10.5), Inches(0.85),
          font_name=SERIF, font_size=42, bold=True, colour=TEAL)
add_text(s1, "Indigenous owned recruitment and IT consultancy, Brisbane QLD",
          Inches(0.3), Inches(1.95), Inches(11.0), Inches(0.4), font_size=13, colour=BLACK)

stat_cards = [
    ("$10.1M", "Annual revenue, FY2025", "Smart50 2025, Nov 2025"),
    ("90%", "Revenue growth year on year", "Smart50 2025 citation"),
    ("#6", "Smart50 2025 national rank", "SmartCompany, Nov 2025"),
    ("2017", "Year founded", "Company records"),
    ("51%", "Indigenous ownership stake", "Company website"),
]
card_w = Inches(2.3)
card_h = Inches(1.6)
card_gap = Inches(0.15)
card_x0 = Inches(0.3)
card_y = Inches(2.4)
for i, (num, label, src) in enumerate(stat_cards):
    x = Inches(0.3 + i * (2.3 + 0.15))
    add_rect(s1, x, card_y, card_w, card_h, BLACK)
    add_rect(s1, x, card_y, card_w, Pt(4), TEAL)
    add_text(s1, num, x + Inches(0.15), card_y + Inches(0.16), card_w - Inches(0.3), Inches(0.55),
              font_name=SERIF, font_size=28, bold=True, colour=AMBER_B)
    add_lines(s1, [label], x + Inches(0.15), card_y + Inches(0.78), card_w - Inches(0.3), Inches(0.5),
              default_size=11, default_colour=OFF_WHITE)
    add_text(s1, src, x + Inches(0.15), card_y + Inches(1.30), card_w - Inches(0.3), Inches(0.25),
              font_size=8, colour=OFF_WHITE, italic=True)

add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.3), Inches(4.15), Inches(6), Inches(0.3),
          font_size=12, bold=True, colour=TEAL)

signals = [
    "6th of 50, 2025 Smart50 Awards: $10.1M revenue, 90% growth (SmartCompany)",
    "Founded 2017 by Phil Ahmat, Anthony Singh, Ash Kumar (Company records)",
    "Supply Nation accredited: 51% owned by CEO Phil Ahmat (MuraConnect site)",
    "Registered NSW Government supplier, buy.nsw panel (buy.nsw.gov.au)",
    "Team draws on more than 25 years serving major public, private clients (Company website)",
    "HQ at Fortitude Valley, inner Brisbane QLD (Company contact page)",
]
y = 4.5
for sig in signals:
    add_text(s1, "•  " + sig, Inches(0.3), Inches(y), Inches(12.4), Inches(0.32), font_size=11, colour=BLACK)
    y += 0.37

add_stripe(s1)
add_footer(s1)

# ============================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
add_header(s2, "THE OPPORTUNITY")

add_text(s2, "MuraConnect: three commercial observations from ProfitPulse",
          Inches(0.3), Inches(1.1), Inches(12.5), Inches(0.4),
          font_name=SERIF, font_size=17, bold=True, colour=BLACK)

columns = [
    {
        "fill": TEAL, "text_colour": WHITE, "para_colour": OFF_WHITE,
        "num": "01", "header": "Growth is outrunning cash conversion",
        "body": ("MuraConnect grew revenue 90 percent to $10.1 million on a recruitment "
                 "and IT project model where placed staff are usually paid weekly while "
                 "enterprise and government clients settle on 30 to 60 day terms. Every "
                 "new placement widens that gap before the cash lands. A Working Capital "
                 "Unlock maps exactly where it sits."),
    },
    {
        "fill": BLACK, "text_colour": AMBER_B, "para_colour": OFF_WHITE,
        "num": "02", "header": "Public sector clients concentrate the risk",
        "body": ("As a Supply Nation accredited, Indigenous owned provider on the buy.nsw "
                 "panel, MuraConnect likely draws a large share of its $10.1 million "
                 "revenue from a small number of public sector and enterprise "
                 "relationships. Knowing which panels and clients drive margin, not just "
                 "turnover, matters as contracts renew."),
    },
    {
        "fill": GOLD, "text_colour": BLACK, "para_colour": BLACK,
        "num": "03", "header": "People cost is the business, so utilisation is the lever",
        "body": ("In a recruitment and IT consultancy, gross margin is set by billable "
                 "utilisation and placement fee capture, not headcount growth. At 90 "
                 "percent growth it is easy to add delivery capacity faster than "
                 "utilisation is tracked. A monthly view of revenue per consultant keeps "
                 "growth profitable, not just larger."),
    },
]

col_w = Inches(4.17)
col_h = Inches(3.9)
col_y = Inches(1.65)
col_xs = [Inches(0.3), Inches(4.57), Inches(8.84)]
for col, x in zip(columns, col_xs):
    add_rect(s2, x, col_y, col_w, col_h, col["fill"])
    add_text(s2, col["num"], x + Inches(0.28), col_y + Inches(0.22), col_w - Inches(0.5), Inches(0.7),
              font_name=SERIF, font_size=36, bold=True, colour=col["text_colour"])
    add_text(s2, col["header"], x + Inches(0.28), col_y + Inches(0.95), col_w - Inches(0.55), Inches(0.85),
              font_size=14, bold=True, colour=col["text_colour"])
    add_text(s2, col["body"], x + Inches(0.28), col_y + Inches(1.85), col_w - Inches(0.55), Inches(2.9),
              font_size=11, colour=col["para_colour"])

add_text(s2, "These are observations offered in good faith. MuraConnect has built "
              "something genuinely impressive in a short time. The question is simply "
              "whether the financial architecture is keeping pace with the growth.",
          Inches(0.3), Inches(5.85), Inches(12.7), Inches(0.6), font_size=11, italic=True, colour=BLACK)

add_stripe(s2)
add_footer(s2)

# ============================================================
# SLIDE 3: THE RECOMMENDATION
# ============================================================
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
add_header(s3, "THE RECOMMENDATION")

# Left column
add_text(s3, "Working Capital Unlock", Inches(0.3), Inches(1.15), Inches(7.4), Inches(0.55),
          font_name=SERIF, font_size=26, bold=True, colour=TEAL)
add_text(s3, "$6,000 one off", Inches(0.3), Inches(1.72), Inches(7.4), Inches(0.4),
          font_size=17, bold=True, colour=BLACK)
add_text(s3, "A four week project mapping cash trapped in debtors, work in progress "
              "and supplier terms, with a prioritised action list to release it.",
          Inches(0.3), Inches(2.18), Inches(7.4), Inches(0.6), font_size=12, colour=BLACK)

add_rect(s3, Inches(0.3), Inches(2.9), Inches(7.4), Inches(1.55), WHITE, line_colour=TEAL, line_width=Pt(1.25))
add_text(s3, "Step one, answer a few quick questions", Inches(0.5), Inches(3.0), Inches(7.0), Inches(0.35),
          font_size=13, bold=True, colour=TEAL)
add_text(s3, "See the solutions matched to your size and industry.",
          Inches(0.5), Inches(3.38), Inches(7.0), Inches(0.32), font_size=11, colour=BLACK)
add_text(s3, QUESTIONNAIRE_CLEAN, Inches(0.5), Inches(3.72), Inches(7.0), Inches(0.35),
          font_size=13, bold=True, colour=TEAL, underline=True, hyperlink=QUESTIONNAIRE_URL)
add_rect(s3, Inches(0.5), Inches(4.12), Inches(3.1), Inches(0.24), TEAL, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s3, "Find your fit in two minutes", Inches(0.5), Inches(4.115), Inches(3.1), Inches(0.24),
          font_size=10, bold=True, colour=WHITE, align=PP_ALIGN.CENTER, hyperlink=QUESTIONNAIRE_URL)

add_rect(s3, Inches(0.3), Inches(4.75), Inches(4.9), Inches(0.5), WHITE, line_colour=AMBER_D, line_width=Pt(1.25),
         shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s3, "Purchase the suggested product now to get started", Inches(0.3), Inches(4.75), Inches(4.9), Inches(0.5),
          font_size=11, bold=True, colour=AMBER_D, align=PP_ALIGN.CENTER, anchor=None, hyperlink=STRIPE_URL)
add_text(s3, "$6,000 one off, ProfitPulse verified price", Inches(0.3), Inches(5.32), Inches(6), Inches(0.3),
          font_size=10, colour=BLACK)

add_text(s3, "Prefer a conversation first?", Inches(0.3), Inches(5.85), Inches(6), Inches(0.32),
          font_size=12, colour=BLACK)
add_text(s3, "Book a complimentary discovery call", Inches(0.3), Inches(6.2), Inches(6), Inches(0.35),
          font_size=12, bold=True, colour=TEAL, underline=True, hyperlink=BOOKING_URL)

# Right panel
panel_x = Inches(8.05)
panel_w = Inches(4.98)
add_rect(s3, panel_x, Inches(1.15), panel_w, Inches(5.75), BLACK)
add_text(s3, "NITESH ROOPA", panel_x + Inches(0.25), Inches(1.35), panel_w - Inches(0.5), Inches(0.45),
          font_size=18, bold=True, colour=AMBER_B)
add_text(s3, "CA, Managing Partner, ProfitPulse", panel_x + Inches(0.25), Inches(1.78), panel_w - Inches(0.5), Inches(0.35),
          font_size=13, colour=WHITE)
add_rect(s3, panel_x + Inches(0.25), Inches(2.18), panel_w - Inches(0.5), Pt(1.5), TEAL)

add_lines(s3, [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal USD 1.3 billion, Cahora Bassa",
    "AUD 10 billion Queensland infrastructure value",
], panel_x + Inches(0.25), Inches(2.35), panel_w - Inches(0.5), Inches(1.6),
    default_size=11, default_colour=OFF_WHITE, line_spacing=1.15)

add_text(s3, "Fractional CFO for growth stage businesses", panel_x + Inches(0.25), Inches(3.85),
          panel_w - Inches(0.5), Inches(0.35), font_size=10.5, italic=True, colour=OFF_WHITE)

add_rect(s3, panel_x + Inches(0.25), Inches(4.35), panel_w - Inches(0.5), Pt(1.5), TEAL)
add_lines(s3, [
    "Profit-Pulse.com.au",
    {"text": "Nitesh@Profit-Pulse.com.au", "colour": TEAL},
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
], panel_x + Inches(0.25), Inches(4.5), panel_w - Inches(0.5), Inches(1.6),
    default_size=11, default_colour=OFF_WHITE, line_spacing=1.15)

add_stripe(s3)
add_footer(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MuraConnect_28Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
