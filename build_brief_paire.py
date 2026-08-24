"""
ProfitPulse Brief Builder v3.3 house style
Target: Paire (Paire Pty Ltd) | Date: 25 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
White body background, black header band, amber left stripe, black stat tiles
with teal top edge, per Section 6 of the nightly outreach engine spec.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import copy

# Brand colours, official ProfitPulse palette, no other colours used anywhere
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

SERIF = "Georgia"
SANS  = "Calibri"

DATE_STR = "25 Aug 2026"
COMPANY  = "Paire"

QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL   = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_C2_COMMAND   = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_LINK        = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_colour
    if not line:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = fill_colour
    shp.shadow.inherit = False
    return shp


def add_text(slide, text, left, top, width, height, font_name=SANS, font_size=14,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False,
             anchor=MSO_ANCHOR.TOP, line_spacing=None, shrink=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if shrink:
        try:
            from pptx.enum.text import MSO_AUTO_SIZE
            tf.auto_size = MSO_AUTO_SIZE.NONE
        except Exception:
            pass
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
    return box


def add_hyperlink_text(slide, text, url, left, top, width, height, font_name=SANS,
                        font_size=14, bold=False, colour=TEAL, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.font.underline = True
    run.hyperlink.address = url
    return box


def add_multiline(slide, lines, left, top, width, height, font_name=SANS,
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   space_after=4, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
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
        p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.color.rgb = cfg["colour"]
    return box


def house_chrome(slide, eyebrow, dark_header=True):
    """Left amber stripe, black header band, footer line. Per Section 6.0A."""
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    add_rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
             font_name=SANS, font_size=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, "PROFITPULSE", Inches(9.5), Inches(0.32), Inches(3.6), Inches(0.4),
             font_name=SANS, font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
    # Footer
    add_rect(slide, Inches(0.35), Inches(7.02), Inches(12.6), Pt(0.75), RGBColor(0xCC, 0xCC, 0xCC))
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
             Inches(0.35), Inches(7.08), Inches(9.5), Inches(0.35),
             font_name=SANS, font_size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.LEFT)
    add_text(slide, DATE_STR, Inches(10.5), Inches(7.08), Inches(2.5), Inches(0.35),
             font_name=SANS, font_size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.RIGHT)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    add_rect(slide, left, top, width, height, BLACK)
    add_rect(slide, left, top, width, Pt(3), TEAL)
    add_text(slide, number, left + Inches(0.12), top + Inches(0.12), width - Inches(0.24), Inches(0.5),
              font_name=SERIF, font_size=27, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_multiline(slide, label_lines, left + Inches(0.12), top + Inches(0.62), width - Inches(0.24), Inches(0.55),
                  font_name=SANS, default_size=10, default_colour=OFF_WHITE, space_after=0, line_spacing=1.0)
    add_text(slide, source, left + Inches(0.12), top + height - Inches(0.32), width - Inches(0.24), Inches(0.28),
              font_name=SANS, font_size=7, italic=True, colour=RGBColor(0x9A, 0x9A, 0x9A), align=PP_ALIGN.LEFT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COMMERCIAL INTELLIGENCE BRIEF
# ═════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
house_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, "Paire", Inches(0.35), Inches(1.12), Inches(8.5), Inches(0.85),
         font_name=SERIF, font_size=42, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s1, "Direct to consumer sock and essentials brand, South Melbourne VIC",
         Inches(0.35), Inches(1.9), Inches(11.5), Inches(0.35),
         font_name=SANS, font_size=13, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)

cards = [
    ("$10.0M", ["Revenue, FY2025"], "SmartCompany Smart50, Nov 2025"),
    ("15/50",  ["Smart50 2025 rank"], "SmartCompany, Nov 2025"),
    ("300K+",  ["Global customers"], "Inside Retail, Aug 2026"),
    ("2020",   ["Founded, garage", "start in Melbourne"], "Power Retail profile"),
    ("60+",    ["Products across", "four categories"], "Inside Retail, Aug 2026"),
    ("26",     ["Team, full time", "and casual staff"], "SmartCompany Smart50 profile"),
]
card_w = Inches(1.98)
gap = Inches(0.15)
x = Inches(0.35)
y_cards = Inches(2.35)
card_h = Inches(1.5)
for number, label_lines, source in cards:
    stat_card(s1, x, y_cards, card_w, card_h, number, label_lines, source)
    x = x + card_w + gap

# Chart: revenue growth FY22 to FY25, two verified data points
chart_x = Inches(0.35)
chart_y = Inches(4.12)
chart_w = Inches(3.9)
chart_h = Inches(2.55)
add_rect(s1, chart_x, chart_y, chart_w, chart_h, RGBColor(0xF4, 0xF4, 0xF2))
add_text(s1, "REVENUE, FY22 TO FY25", chart_x + Inches(0.18), chart_y + Inches(0.12),
         chart_w - Inches(0.36), Inches(0.3), font_name=SANS, font_size=10, bold=True,
         colour=TEAL, align=PP_ALIGN.LEFT)

base_y = chart_y + Inches(2.05)
max_bar_h = Inches(1.35)
# FY22 bar: $1.7M
bar1_h = Emu(int(max_bar_h * (1.7 / 10.0)))
bar1_x = chart_x + Inches(1.0)
bar1_w = Inches(0.85)
add_rect(s1, bar1_x, base_y - bar1_h, bar1_w, bar1_h, TEAL)
add_text(s1, "$1.7M", bar1_x - Inches(0.2), base_y - bar1_h - Inches(0.32), Inches(1.25), Inches(0.28),
         font_name=SANS, font_size=10, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
add_text(s1, "FY2022", bar1_x - Inches(0.2), base_y + Inches(0.06), Inches(1.25), Inches(0.28),
         font_name=SANS, font_size=9, colour=RGBColor(0x55, 0x55, 0x55), align=PP_ALIGN.CENTER)
# FY25 bar: $10.0M
bar2_h = max_bar_h
bar2_x = chart_x + Inches(2.3)
bar2_w = Inches(0.85)
add_rect(s1, bar2_x, base_y - bar2_h, bar2_w, bar2_h, AMBER_D)
add_text(s1, "$10.0M", bar2_x - Inches(0.2), base_y - bar2_h - Inches(0.32), Inches(1.25), Inches(0.28),
         font_name=SANS, font_size=10, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
add_text(s1, "FY2025", bar2_x - Inches(0.2), base_y + Inches(0.06), Inches(1.25), Inches(0.28),
         font_name=SANS, font_size=9, colour=RGBColor(0x55, 0x55, 0x55), align=PP_ALIGN.CENTER)
add_rect(s1, chart_x + Inches(0.18), base_y, chart_w - Inches(0.36), Pt(1), RGBColor(0xAA, 0xAA, 0xAA))
add_text(s1, "Source: SmartCompany Smart50 2025 growth profile", chart_x + Inches(0.18),
         chart_y + chart_h - Inches(0.3), chart_w - Inches(0.36), Inches(0.26),
         font_name=SANS, font_size=7, italic=True, colour=RGBColor(0x77, 0x77, 0x77), align=PP_ALIGN.LEFT)

# Key Commercial Signals
sig_x = Inches(4.55)
sig_y = Inches(4.12)
sig_w = Inches(8.43)
add_text(s1, "KEY COMMERCIAL SIGNALS", sig_x, sig_y, sig_w, Inches(0.3),
         font_name=SANS, font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
signals = [
    "Revenue grew from $1.7M (FY22) to $10.0M (FY25), Smart50 2025 rank 15 of 50.",
    "Converted pop up trials into two permanent stores, QV Melbourne and South Melbourne.",
    "Entered the United States market in 2026, its first launch outside ANZ.",
    "Also trading into Singapore and Malaysia; range now exceeds 60 SKUs.",
    "Declined Shark Tank Australia offers in November 2024, staying founder funded.",
    "Planning further Melbourne and Sydney stores within the next 12 to 18 months.",
]
add_multiline(s1, signals, sig_x, sig_y + Inches(0.4), sig_w, Inches(2.1),
              font_name=SANS, default_size=11, default_colour=RGBColor(0x22, 0x22, 0x22),
              space_after=8, line_spacing=1.05)
add_text(s1, "Sources: SmartCompany Smart50 2025 series, Inside Retail Aug 2026, FashionUnited Aug 2026",
         sig_x, Inches(6.55), sig_w, Inches(0.3),
         font_name=SANS, font_size=7, italic=True, colour=RGBColor(0x77, 0x77, 0x77), align=PP_ALIGN.LEFT)

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 2 — THE OPPORTUNITY
# ═════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
house_chrome(s2, "THE OPPORTUNITY")
add_text(s2, "Paire: three commercial observations from ProfitPulse",
         Inches(0.35), Inches(1.02), Inches(12.6), Inches(0.4),
         font_name=SANS, font_size=13, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)

col_y = Inches(1.55)
col_h = Inches(4.0)
col_w = Inches(4.14)
col_gap = Inches(0.15)
col_x = [Inches(0.35), Inches(0.35) + col_w + col_gap, Inches(0.35) + 2 * (col_w + col_gap)]
fills = [TEAL, BLACK, GOLD]
text_colours = [BLACK, WHITE, BLACK]

headers = [
    "Cash is funding four fronts at once",
    "SKU growth has outpaced margin visibility",
    "Growth is entirely founder funded",
]
bodies = [
    ("Paire now trades in five markets at once: Australia, the US, Singapore and "
     "Malaysia online, plus two Melbourne stores. Each market and each store needs "
     "its own stock position before a single sale converts back to cash. Funding "
     "this many fronts simultaneously, on top of a wider domestic rollout already "
     "planned, is a working capital question before it is anything else. "
     "A Working Capital Unlock maps exactly where that cash is trapped."),
    ("Sixty plus products now sit across socks, underwear, activewear and outerwear, "
     "sold through direct online, two flagship stores and three overseas markets. At "
     "$1.7 million the product mix was simple to read by eye. At $10.0 million across "
     "this many channels, blended revenue growth can mask lines and markets that are "
     "actually thin on margin once freight, duty and store overhead are allocated "
     "properly against each one."),
    ("Paire walked away from Shark Tank offers in November 2024 rather than give up "
     "2.5 percent of the business. That discipline means every store opening and "
     "every new market this year is funded from trading cash, not investor cash, "
     "which raises the cost of any dollar left sitting idle in stock or receivables, "
     "and makes disciplined cash management the difference between funding the next "
     "market and stalling before it."),
]

for i in range(3):
    add_rect(s2, col_x[i], col_y, col_w, col_h, fills[i])
    add_text(s2, f"0{i+1}", col_x[i] + Inches(0.22), col_y + Inches(0.18), col_w - Inches(0.44), Inches(0.75),
             font_name=SERIF, font_size=34, bold=True, colour=text_colours[i], align=PP_ALIGN.LEFT)
    add_text(s2, headers[i], col_x[i] + Inches(0.22), col_y + Inches(0.95), col_w - Inches(0.44), Inches(0.75),
             font_name=SERIF, font_size=15, bold=True, colour=text_colours[i], align=PP_ALIGN.LEFT)
    add_text(s2, bodies[i], col_x[i] + Inches(0.22), col_y + Inches(1.75), col_w - Inches(0.44), Inches(2.15),
             font_name=SANS, font_size=11, colour=text_colours[i], align=PP_ALIGN.LEFT, line_spacing=1.15)

add_text(s2, ("These are observations offered in good faith. Paire has built something genuinely "
              "impressive from a garage start. The question is simply whether the cash engine is "
              "built to match the ambition of five markets at once."),
         Inches(0.35), Inches(5.95), Inches(12.6), Inches(0.55),
         font_name=SANS, font_size=11.5, italic=True, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 3 — THE RECOMMENDATION AND HOW TO START
# ═════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
house_chrome(s3, "THE RECOMMENDATION")

left_x = Inches(0.35)
left_w = Inches(7.6)
right_x = Inches(8.25)
right_w = Inches(4.73)

add_text(s3, "Working Capital Unlock", left_x, Inches(1.15), left_w, Inches(0.55),
         font_name=SERIF, font_size=25, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "$6,000 one off", left_x, Inches(1.68), left_w, Inches(0.35),
         font_name=SANS, font_size=15, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, ("A four week project that maps cash trapped in stock, receivables, supplier terms "
              "and banking facilities across every Paire channel and market, then hands over a "
              "prioritised release plan. Typical clients release 8 to 15 percent of revenue back "
              "into the business without raising outside capital."),
         left_x, Inches(2.08), left_w, Inches(1.0),
         font_name=SANS, font_size=11, colour=RGBColor(0x22, 0x22, 0x22), align=PP_ALIGN.LEFT, line_spacing=1.1)

# Step one block
add_rect(s3, left_x, Inches(3.25), left_w, Inches(1.35), RGBColor(0xF4, 0xF4, 0xF2))
add_text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.38), left_w - Inches(0.4), Inches(0.32),
         font_name=SANS, font_size=12, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.2), Inches(3.72), left_w - Inches(0.4), Inches(0.3),
         font_name=SANS, font_size=10.5, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)
add_hyperlink_text(s3, QUESTIONNAIRE_CLEAN, QUESTIONNAIRE_URL, left_x + Inches(0.2), Inches(4.05),
                    left_w - Inches(0.4), Inches(0.35), font_size=12, bold=True, colour=TEAL)

# Direct CTA
add_rect(s3, left_x, Inches(4.78), left_w, Inches(0.62), TEAL)
cta_box = s3.shapes.add_textbox(left_x, Inches(4.78), left_w, Inches(0.62))
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
run.hyperlink.address = STRIPE_C2_COMMAND

add_text(s3, "Prefer a conversation first?", left_x, Inches(5.62), left_w, Inches(0.3),
         font_name=SANS, font_size=11, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)
add_hyperlink_text(s3, "Book a complimentary discovery call", BOOKING_LINK, left_x, Inches(5.92),
                    left_w, Inches(0.35), font_size=11, bold=True, colour=TEAL)

# Right column: credibility panel
add_rect(s3, right_x, Inches(1.15), right_w, Inches(4.75), BLACK)
add_text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.35), right_w - Inches(0.5), Inches(0.4),
         font_name=SERIF, font_size=18, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.78), right_w - Inches(0.5), Inches(0.32),
         font_name=SANS, font_size=11.5, colour=WHITE, align=PP_ALIGN.LEFT)
add_rect(s3, right_x + Inches(0.25), Inches(2.18), right_w - Inches(0.5), Pt(1.5), TEAL)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
add_multiline(s3, cred, right_x + Inches(0.25), Inches(2.35), right_w - Inches(0.5), Inches(1.5),
              font_name=SANS, default_size=10.5, default_colour=OFF_WHITE, space_after=6, line_spacing=1.1)

add_rect(s3, right_x + Inches(0.25), Inches(3.95), right_w - Inches(0.5), Pt(1.5), TEAL)
contact = [
    {"text": "Profit-Pulse.com.au", "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "colour": TEAL},
    {"text": "+61 411 876 267", "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "colour": TEAL},
]
add_multiline(s3, contact, right_x + Inches(0.25), Inches(4.12), right_w - Inches(0.5), Inches(1.4),
              font_name=SANS, default_size=10.5, space_after=6, line_spacing=1.1)

add_rect(s3, right_x + Inches(0.25), Inches(5.15), right_w - Inches(0.5), Pt(1.5), TEAL)
add_text(s3, "Fractional CFO for growth stage SMEs: cash flow, investor ready "
             "reporting and growth strategy.",
         right_x + Inches(0.25), Inches(5.3), right_w - Inches(0.5), Inches(0.55),
         font_name=SANS, font_size=10, italic=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_25Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
