"""
ProfitPulse Brief Builder v3.3 house style
Target: TMC Fine Jewellers | Date: 12 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
White body, black header band, amber left stripe, teal topped stat cards.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import pptx.util as ptu

# Brand colours, exactly seven
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Georgia"
SANS  = "Arial"

W = Inches(13.333)
H = Inches(7.5)

DATE_STAMP = "12 Aug 2026"
COMPANY = "TMC Fine Jewellers"


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
    return shape


def add_text(slide, text, left, top, width, height, font_name=SANS, font_size=14,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False,
             anchor=MSO_ANCHOR.TOP, line_spacing=None, hyperlink=None, wrap=True,
             shrink=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
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
        run.hyperlink.address = hyperlink
    return box


def add_multipara(slide, paras, left, top, width, height, font_name=SANS,
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   line_spacing=1.12, space_after=6, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for para in paras:
        cfg = {"text": para, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False} if isinstance(para, str) else {
            "text": para.get("text", ""), "size": para.get("size", default_size),
            "colour": para.get("colour", default_colour), "bold": para.get("bold", False),
            "italic": para.get("italic", False)}
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
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def house_chrome(slide, eyebrow):
    """Left amber stripe, black header band, eyebrow + PROFITPULSE, footer line."""
    set_background(slide, WHITE)
    # Header band
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.36), Inches(8.5), Inches(0.35),
              font_name=SANS, font_size=12, bold=True, colour=OFF_WHITE)
    add_text(slide, "PROFITPULSE", Inches(9.4), Inches(0.36), Inches(3.6), Inches(0.35),
              font_name=SANS, font_size=12, bold=True, colour=WHITE, align=PP_ALIGN.RIGHT)
    # Left accent stripe, full height, drawn on top so it reads unbroken
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    # Footer
    add_rect(slide, Inches(0.35), Inches(7.02), Inches(12.6), Pt(0.75), TEAL)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.35), Inches(7.1), Inches(9.5), Inches(0.3),
              font_name=SANS, font_size=8, colour=BLACK)
    add_text(slide, DATE_STAMP, Inches(10.5), Inches(7.1), Inches(2.5), Inches(0.3),
              font_name=SANS, font_size=8, colour=BLACK, align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
house_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, COMPANY, Inches(0.35), Inches(1.15), Inches(11.5), Inches(0.75),
          font_name=SERIF, font_size=40, bold=True, colour=BLACK)
add_text(s1, "Lab grown diamond and moissanite bridal jewellery house, New Farm, Brisbane QLD",
          Inches(0.35), Inches(2.05), Inches(12.5), Inches(0.4),
          font_name=SANS, font_size=13, colour=BLACK)

# Stat cards, single row of five, shared top edge and height
card_w = Inches(2.3)
card_h = Inches(1.6)
card_top = Inches(2.55)
gap = Inches(0.15)
start_x = Inches(0.35)

cards = [
    ("$12.7M", "FY2025 revenue", "Smart50 2025, rank 5"),
    ("95%",    "Revenue growth, FY25", "Smart50 2025 citation"),
    ("35",     "Team members", "Smart50 2025 citation"),
    ("4",      "Global showroom cities", "TMC showroom pages"),
    ("2020",   "Year founded", "TMC Our Story page"),
]

x = start_x
for number, label, source in cards:
    add_rect(s1, x, card_top, card_w, card_h, BLACK)
    add_rect(s1, x, card_top, card_w, Pt(4), TEAL)
    add_text(s1, number, x + Inches(0.18), card_top + Inches(0.18), card_w - Inches(0.36), Inches(0.55),
              font_name=SERIF, font_size=28, bold=True, colour=AMBER_B)
    add_multipara(s1, [label], x + Inches(0.18), card_top + Inches(0.8), card_w - Inches(0.36), Inches(0.5),
                   default_size=11, default_colour=OFF_WHITE, line_spacing=1.05, space_after=0)
    add_text(s1, source, x + Inches(0.18), card_top + Inches(1.34), card_w - Inches(0.36), Inches(0.22),
              font_name=SANS, font_size=8, colour=OFF_WHITE)
    x += card_w + gap

# Key Commercial Signals
sig_top = Inches(4.5)
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), sig_top, Inches(6), Inches(0.3),
          font_name=SANS, font_size=12, bold=True, colour=TEAL)

signals = [
    "Smart50 2025 ranks TMC fifth nationally with 95 percent three year growth to $12.7M, per SmartCompany.",
    "Fourth global showroom opened in Armadale, Melbourne, 6 December 2025, after Brisbane, Sydney and a Mayfair, London pop up in June 2025, per Inside Retail.",
    "BFCM week 2025 was targeted to deliver over $3.75 million, against $2.6 million and 1,200 customers in the 2024 event, per SmartCompany.",
    "Brand renamed from The Moissanite Company to TMC Fine Jewellers as it broadened into lab grown diamonds, per Ragtrader.",
    "Founded 2020 by Makayla and Tom Donovan on a $10,000 initial investment, per Onya Magazine and company published material.",
]
add_multipara(s1, signals, Inches(0.35), sig_top + Inches(0.42), Inches(12.6), Inches(2.1),
               default_size=11.5, default_colour=BLACK, line_spacing=1.18, space_after=8)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
house_chrome(s2, "THE OPPORTUNITY")
add_text(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
          Inches(0.35), Inches(1.08), Inches(12.5), Inches(0.4),
          font_name=SERIF, font_size=16, bold=True, colour=BLACK)

col_top = Inches(1.65)
col_h = Inches(4.85)
col_w = Inches(4.24)
gaps2 = Inches(0.11)
col_x = [Inches(0.35), Inches(0.35) + col_w + gaps2, Inches(0.35) + 2 * (col_w + gaps2)]
fills = [TEAL, BLACK, GOLD]
text_colours = [BLACK, WHITE, BLACK]
sub_colours = [BLACK, TEAL, BLACK]

observations = [
    ("01", "Four showrooms, one working capital question",
     "Brisbane, Sydney, Melbourne and a London pop up now each hold moissanite and lab grown "
     "diamond stock ahead of sale. Every new showroom and every pre BFCM build locks up cash "
     "before it converts to revenue, and four sites multiply that exposure at once."),
    ("02", "One week now carries close to a third of the year",
     "The 2025 BFCM target of $3.75 million is planned six to eight months out, up from $2.6 "
     "million in 2024. That concentration is a strength when forecast well and a liquidity "
     "shock when the stock build outpaces the cash available to fund it."),
    ("03", "Growth has outrun the reporting cadence",
     "Four showrooms across three countries within about eighteen months is a rapid build for "
     "a five year old business. Matching that pace with a costed twelve month plan keeps the "
     "next site funded by design rather than by whichever quarter had cash to spare."),
]

for i, (idx, header, para) in enumerate(observations):
    add_rect(s2, col_x[i], col_top, col_w, col_h, fills[i])
    pad = Inches(0.28)
    add_text(s2, idx, col_x[i] + pad, col_top + Inches(0.22), col_w - 2*pad, Inches(0.7),
              font_name=SERIF, font_size=34, bold=True, colour=sub_colours[i])
    add_text(s2, header, col_x[i] + pad, col_top + Inches(0.95), col_w - 2*pad, Inches(0.85),
              font_name=SERIF, font_size=15, bold=True, colour=text_colours[i], line_spacing=1.08)
    add_multipara(s2, [para], col_x[i] + pad, col_top + Inches(1.95), col_w - 2*pad, Inches(2.8),
                   default_size=11, default_colour=text_colours[i], line_spacing=1.22, space_after=0)

add_text(s2,
         "These observations are offered in good faith. Makayla and Tom have built something rare "
         "in five years. The question is simply whether the financial architecture keeps pace with the shopfronts.",
         Inches(0.35), Inches(6.6), Inches(12.6), Inches(0.5),
         font_name=SANS, font_size=10.5, italic=True, colour=BLACK)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
house_chrome(s3, "THE RECOMMENDATION")

left_x = Inches(0.35)
left_w = Inches(7.65)

add_text(s3, "Working Capital Unlock", left_x, Inches(1.15), left_w, Inches(0.55),
          font_name=SERIF, font_size=27, bold=True, colour=BLACK)
add_text(s3, "$7,500 one off", left_x, Inches(1.72), left_w, Inches(0.4),
          font_name=SANS, font_size=16, bold=True, colour=TEAL)
add_multipara(s3, [
    "Maps cash trapped in showroom stock, workshop inventory and supplier terms across all "
    "four sites, with a prioritised plan to release it before the next major sale week."],
    left_x, Inches(2.18), left_w, Inches(0.75), default_size=11.5,
    default_colour=BLACK, line_spacing=1.2)

# Step one block
add_rect(s3, left_x, Inches(3.05), left_w, Inches(1.55), OFF_WHITE,
          line_colour=TEAL, line_width=Pt(1))
add_text(s3, "Step one, answer a few quick questions", left_x + Inches(0.25), Inches(3.22),
          left_w - Inches(0.5), Inches(0.35), font_name=SANS, font_size=13, bold=True, colour=BLACK)
add_text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.25), Inches(3.62),
          left_w - Inches(0.5), Inches(0.3), font_name=SANS, font_size=11, colour=BLACK)
add_text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.25), Inches(3.98),
          left_w - Inches(0.5), Inches(0.35), font_name=SANS, font_size=13, bold=True, colour=TEAL,
          hyperlink="https://profit-pulse.com.au/services/find-your-fit/")

# Direct CTA
add_rect(s3, left_x, Inches(4.78), left_w, Inches(0.7), AMBER_D)
add_text(s3, "Purchase the suggested product now to get started", left_x, Inches(4.78), left_w, Inches(0.7),
          font_name=SANS, font_size=13, bold=True, colour=BLACK, align=PP_ALIGN.CENTER,
          anchor=MSO_ANCHOR.MIDDLE, hyperlink="https://buy.stripe.com/cNi28s0msgKQ0yQeQL3ks0A")

add_text(s3, "Prefer a conversation first?", left_x, Inches(5.68), left_w, Inches(0.3),
          font_name=SANS, font_size=11, colour=BLACK)
add_text(s3, "Book a complimentary discovery call", left_x, Inches(5.98), left_w, Inches(0.35),
          font_name=SANS, font_size=12, bold=True, colour=TEAL,
          hyperlink="https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right column: credibility panel
right_x = Inches(8.25)
right_w = Inches(4.73)
add_rect(s3, right_x, Inches(1.15), right_w, Inches(5.5), BLACK)
add_rect(s3, right_x, Inches(1.15), right_w, Pt(4), TEAL)
pad = Inches(0.3)
add_text(s3, "Nitesh Roopa", right_x + pad, Inches(1.45), right_w - 2*pad, Inches(0.45),
          font_name=SERIF, font_size=19, bold=True, colour=AMBER_B)
add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + pad, Inches(1.92), right_w - 2*pad, Inches(0.35),
          font_name=SANS, font_size=12, colour=WHITE)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
add_multipara(s3, cred, right_x + pad, Inches(2.45), right_w - 2*pad, Inches(1.7),
               default_size=11, default_colour=OFF_WHITE, line_spacing=1.15, space_after=8)

add_rect(s3, right_x + pad, Inches(4.25), right_w - 2*pad, Pt(1), TEAL)

contact = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
add_multipara(s3, contact, right_x + pad, Inches(4.45), right_w - 2*pad, Inches(1.8),
               default_size=11, default_colour=OFF_WHITE, line_spacing=1.25, space_after=6)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TMCFineJewellers_12Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
