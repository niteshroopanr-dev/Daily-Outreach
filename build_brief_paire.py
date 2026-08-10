"""
ProfitPulse Brief Builder v3.3 house style
Target: Paire | Date: 11 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
White body background, black header band, amber left accent stripe,
black stat cards with teal top edge, serif headings, sans body.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

# Brand colours, exactly seven, no others used anywhere in this file.
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Georgia"
SANS  = "Calibri"

W = Inches(13.333)
H = Inches(7.5)
DATE_STAMP = "11 Aug 2026"
COMPANY = "Paire"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shp = slide.shapes.add_shape(1, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_width or Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, left, top, width, height, size=12, bold=False, colour=BLACK,
         align=PP_ALIGN.LEFT, font=SANS, italic=False, hyperlink=None, wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = s
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    if hyperlink:
        run.hyperlink.address = hyperlink
    return box


def multiline(slide, lines, left, top, width, height, font=SANS, size=11,
              colour=BLACK, align=PP_ALIGN.LEFT, space_after=4, bold=False,
              line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        cfg = {"text": line, "size": size, "colour": colour, "bold": bold,
               "italic": False, "font": font} if isinstance(line, str) else {
               "text": line.get("text", ""), "size": line.get("size", size),
               "colour": line.get("colour", colour), "bold": line.get("bold", bold),
               "italic": line.get("italic", False), "font": line.get("font", font)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = cfg["font"]
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def header_band(slide, eyebrow, right_label, right_colour=TEAL):
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.34), Inches(7.5), Inches(0.4),
         size=12, bold=True, colour=OFF_WHITE, font=SANS)
    text(slide, right_label, Inches(9.0), Inches(0.34), Inches(4.0), Inches(0.4),
         size=13, bold=True, colour=right_colour, align=PP_ALIGN.RIGHT, font=SANS)


def footer(slide):
    text(slide, "Prepared by Nitesh Roopa, CA, Managing Partner, ProfitPulse. Profit-Pulse.com.au",
         Inches(0.35), Inches(7.08), Inches(9.5), Inches(0.3), size=9, colour=BLACK, font=SANS)
    text(slide, DATE_STAMP, Inches(10.5), Inches(7.08), Inches(2.48), Inches(0.3),
         size=9, colour=BLACK, align=PP_ALIGN.RIGHT, font=SANS)


def left_stripe(slide):
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)


def click_overlay(slide, left, top, width, height, url):
    """Invisible rectangle carrying a click hyperlink, laid over styled text.
    Avoids PowerPoint/LibreOffice forcing the theme hyperlink colour (blue,
    underlined) onto a run that has run.hyperlink.address set directly."""
    shp = slide.shapes.add_shape(1, left, top, width, height)
    shp.fill.background()
    shp.line.fill.background()
    shp.shadow.inherit = False
    shp.click_action.hyperlink.address = url
    return shp


# ════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, WHITE)
header_band(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE", TEAL)

text(s1, "Paire", Inches(0.35), Inches(1.12), Inches(9), Inches(0.75),
     size=42, bold=True, colour=BLACK, font=SERIF)
text(s1, "Direct to consumer apparel brand, Melbourne VIC",
     Inches(0.35), Inches(1.9), Inches(10), Inches(0.35), size=13, colour=TEAL, font=SANS)

# Stat cards
stat_cards = [
    ("$9.4M", ["FY2025 revenue"], "Smart50 2025 citation"),
    ("64%",   ["3yr avg growth"], "Smart50 2025 citation"),
    ("40",    ["Team members"], "Inside Retail, Aug 2026"),
    ("2020",  ["Founded"], "SmartCompany profile"),
    ("4",     ["Markets", "selling in"], "Inside Retail, Aug 2026"),
]
card_w = Inches(2.3)
card_h = Inches(1.6)
gap = Inches(0.15)
card_y = Inches(2.4)
x = Inches(0.35)
for num, label_lines, src in stat_cards:
    rect(s1, x, card_y, card_w, card_h, BLACK)
    rect(s1, x, card_y, card_w, Pt(4), TEAL)
    text(s1, num, x + Inches(0.15), card_y + Inches(0.14), card_w - Inches(0.3), Inches(0.5),
         size=28, bold=True, colour=WHITE, font=SERIF)
    multiline(s1, label_lines, x + Inches(0.15), card_y + Inches(0.68), card_w - Inches(0.3),
              Inches(0.5), size=11, colour=OFF_WHITE, font=SANS, space_after=0)
    text(s1, src, x + Inches(0.15), card_y + Inches(1.26), card_w - Inches(0.3), Inches(0.28),
         size=7, colour=OFF_WHITE, font=SANS, italic=True)
    x += card_w + gap

# Simple verified revenue chart, two data points
chart_x = Inches(0.35)
chart_y = Inches(4.25)
chart_w = Inches(4.4)
text(s1, "REVENUE, VERIFIED", chart_x, chart_y, chart_w, Inches(0.3),
     size=11, bold=True, colour=TEAL, font=SANS)
base_y = chart_y + Inches(2.15)
bar_w = Inches(0.9)
b1_h = Inches(0.35)
b2_h = Inches(1.5)
b1_x = chart_x + Inches(0.5)
b2_x = chart_x + Inches(2.3)
text(s1, "$10M", b2_x, base_y - b2_h - Inches(0.32), bar_w + Inches(0.3), Inches(0.3),
     size=13, bold=True, colour=BLACK, font=SERIF)
rect(s1, b2_x, base_y - b2_h, bar_w, b2_h, TEAL)
text(s1, "2024", b2_x, base_y + Inches(0.05), bar_w, Inches(0.28), size=10, colour=BLACK, font=SANS)
text(s1, "$1M", b1_x, base_y - b1_h - Inches(0.32), bar_w + Inches(0.3), Inches(0.3),
     size=13, bold=True, colour=BLACK, font=SERIF)
rect(s1, b1_x, base_y - b1_h, bar_w, b1_h, TEAL)
text(s1, "2021", b1_x, base_y + Inches(0.05), bar_w, Inches(0.28), size=10, colour=BLACK, font=SANS)
text(s1, "Source: SmartCompany, Smart50 2025 profile (co founder statement)",
     chart_x, base_y + Inches(0.4), chart_w, Inches(0.3), size=7, colour=BLACK, italic=True, font=SANS)

# Key commercial signals
sig_x = Inches(5.15)
sig_w = Inches(7.85)
text(s1, "KEY COMMERCIAL SIGNALS", sig_x, Inches(4.25), sig_w, Inches(0.3),
     size=11, bold=True, colour=TEAL, font=SANS)
signals = [
    "Revenue grew from $1M in 2021 to $10M in 2024, a tenfold rise. (SmartCompany, Smart50 2025)",
    "Ranked 15th on the 2025 Smart50 list, 64% three year average growth. (SmartCompany citation)",
    "Opened a permanent flagship at QV Melbourne after pop up trials. (SmartCompany retail coverage)",
    "Trades bricks and mortar in Singapore and Malaysia alongside Melbourne. (SmartCompany profile)",
    "Launched direct to consumer in the United States, August 2026. (Inside Retail Australia)",
    "Team has grown to 40 people supporting the international rollout. (Inside Retail, Aug 2026)",
]
multiline(s1, signals, sig_x, Inches(4.62), sig_w, Inches(2.2), size=10.5, colour=BLACK,
          font=SANS, space_after=8, line_spacing=1.05)

footer(s1)
left_stripe(s1)

# ════════════════════════════════════════════════════════════════════════
# SLIDE 2 — THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, WHITE)
header_band(s2, "THE OPPORTUNITY", COMPANY, TEAL)
text(s2, "Paire: three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.05), Inches(12.6), Inches(0.35),
     size=14, bold=True, colour=BLACK, align=PP_ALIGN.LEFT, font=SERIF)

col_y = Inches(1.55)
col_h = Inches(4.75)
col_w = Inches(4.411)
cols = [
    (Inches(0.1), TEAL, BLACK, "01", "Four markets in three years",
     "Paire has taken direct to consumer retail from a single Melbourne flagship at "
     "QV to stores in Singapore and Malaysia, then a United States launch in August "
     "2026. Four markets in six years is rapid capital deployment across inventory, "
     "fit out and marketing. The Strategic Growth Diagnostic maps that expansion "
     "into one costed twelve month plan before the next market is chosen."),
    (Inches(4.511), BLACK, OFF_WHITE, "02", "Revenue times ten, team times eight",
     "Revenue grew from $1 million in 2021 to $10 million in 2024, and the team "
     "has grown to 40 people, yet public coverage names only the two co founders "
     "across creative, textile and growth roles. A monthly Fractional CFO "
     "Partnership puts a senior financial voice at the table as the business "
     "scales into new markets and inventory commitments grow heavier."),
    (Inches(8.922), GOLD, BLACK, "03", "Where the next dollar of capital goes",
     "Three new markets, a permanent flagship, and a product range grown from "
     "three sock styles to more than 60 lines all draw on the same capital base. "
     "A Capital Allocation Review would test how marketing, inventory and fit out "
     "spend across Melbourne, Singapore, Malaysia and the United States are "
     "earning their return before the next market decision is made."),
]
for cx, fill, txt_colour, idx, head, para in cols:
    rect(s2, cx, col_y, col_w, col_h, fill)
    pad = Inches(0.3)
    text(s2, idx, cx + pad, col_y + Inches(0.22), col_w - Inches(0.6), Inches(0.6),
         size=34, bold=True, colour=txt_colour, font=SERIF)
    text(s2, head, cx + pad, col_y + Inches(0.92), col_w - Inches(0.6), Inches(0.65),
         size=14, bold=True, colour=txt_colour, font=SERIF)
    text(s2, para, cx + pad, col_y + Inches(1.6), col_w - Inches(0.6), Inches(2.9),
         size=11, colour=txt_colour, font=SANS)

text(s2, "These observations are offered in good faith. Paire has built something "
         "genuinely impressive. The question is simply whether the financial "
         "architecture matches the ambition.",
     Inches(0.35), Inches(6.42), Inches(12.6), Inches(0.5),
     size=10.5, italic=True, colour=BLACK, align=PP_ALIGN.CENTER, font=SANS)

footer(s2)
left_stripe(s2)

# ════════════════════════════════════════════════════════════════════════
# SLIDE 3 — THE RECOMMENDATION AND HOW TO START
# ════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)
header_band(s3, "THE RECOMMENDATION", COMPANY, TEAL)

left_x = Inches(0.35)
left_w = Inches(7.35)

text(s3, "Strategic Growth Diagnostic", left_x, Inches(1.15), left_w, Inches(0.55),
     size=24, bold=True, colour=BLACK, font=SERIF)
text(s3, "$5,000 one off", left_x, Inches(1.72), left_w, Inches(0.4),
     size=19, bold=True, colour=AMBER_D, font=SERIF)
text(s3, "A six week engagement mapping revenue, capacity and margin headroom into "
         "a 12 month growth plan, with funding and capital allocation steps for "
         "the next market spelled out.",
     left_x, Inches(2.18), left_w, Inches(0.75), size=12, colour=BLACK, font=SANS)

# Step one block
rect(s3, left_x, Inches(3.0), left_w, Inches(1.35), TEAL)
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.25), Inches(3.15),
     left_w - Inches(0.5), Inches(0.35), size=13, bold=True, colour=BLACK, font=SANS)
text(s3, "See the solutions matched to your size and industry.",
     left_x + Inches(0.25), Inches(3.52), left_w - Inches(0.5), Inches(0.3),
     size=11, colour=BLACK, font=SANS)
text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.25), Inches(3.88),
     left_w - Inches(0.5), Inches(0.35), size=13, bold=True, colour=BLACK, font=SANS)
click_overlay(s3, left_x + Inches(0.25), Inches(3.88), left_w - Inches(0.5), Inches(0.35),
              "https://profit-pulse.com.au/services/find-your-fit/")

# Direct purchase CTA
rect(s3, left_x, Inches(4.55), left_w, Inches(0.65), WHITE, line_colour=AMBER_D, line_width=Pt(1.5))
text(s3, "Purchase the suggested product now to get started", left_x + Inches(0.3), Inches(4.72),
     left_w - Inches(0.6), Inches(0.35), size=13, bold=True, colour=AMBER_D, font=SANS,
     align=PP_ALIGN.CENTER)
click_overlay(s3, left_x, Inches(4.55), left_w, Inches(0.65),
              "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h")

text(s3, "Prefer a conversation first?", left_x, Inches(5.45), left_w, Inches(0.3),
     size=11, colour=BLACK, font=SANS)
text(s3, "Book a complimentary discovery call", left_x, Inches(5.78), left_w, Inches(0.35),
     size=12, bold=True, colour=TEAL, font=SANS)
click_overlay(s3, left_x, Inches(5.78), left_w, Inches(0.35),
              "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right panel, credibility
panel_x = Inches(8.05)
panel_w = Inches(4.93)
rect(s3, panel_x, Inches(1.15), panel_w, Inches(5.55), BLACK)
rect(s3, panel_x, Inches(1.15), panel_w, Pt(4), TEAL)
text(s3, "NITESH ROOPA", panel_x + Inches(0.3), Inches(1.4), panel_w - Inches(0.6), Inches(0.4),
     size=17, bold=True, colour=AMBER_B, font=SERIF)
text(s3, "CA, Managing Partner, ProfitPulse", panel_x + Inches(0.3), Inches(1.82),
     panel_w - Inches(0.6), Inches(0.35), size=12, colour=WHITE, font=SANS)
rect(s3, panel_x + Inches(0.3), Inches(2.28), panel_w - Inches(0.6), Pt(1.5), TEAL)

credibility = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3 billion, Cahora Bassa",
    "Over AUD 10 billion GRBT project value",
]
multiline(s3, credibility, panel_x + Inches(0.3), Inches(2.5), panel_w - Inches(0.6), Inches(1.5),
          size=11, colour=OFF_WHITE, font=SANS, space_after=8)

contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
multiline(s3, contact_lines, panel_x + Inches(0.3), Inches(4.3), panel_w - Inches(0.6), Inches(1.5),
          size=11, colour=TEAL, font=SANS, space_after=6)

footer(s3)
left_stripe(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_11Aug2026.pptx"
prs.save(out_path)
print("PPTX saved:", out_path)
