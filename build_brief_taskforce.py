"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 01 Aug 2026
Three slide prospect facing deck. House style per Section 6.0 and 6.0A.
Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
HEAD_SERIF = "Cambria"
BODY_SANS = "Calibri"

DATE_STR = "01 Aug 2026"
COMPANY = "Taskforce Australia"


def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, l, t, w, h, fill_colour, line_colour=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, l, t, w, h, size=12, bold=False, italic=False,
         colour=BLACK, align=PP_ALIGN.LEFT, font=BODY_SANS,
         anchor=MSO_ANCHOR.TOP, wrap=True, shrink=True, line_spacing=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    if shrink:
        from pptx.enum.text import MSO_AUTO_SIZE
        tf.auto_size = MSO_AUTO_SIZE.NONE
    lines = s.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = colour
    return box


def header_band(slide, eyebrow, company_or_pp="PROFITPULSE"):
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(7.5), Inches(0.4),
         size=12, bold=True, colour=OFF_WHITE, font=BODY_SANS,
         anchor=MSO_ANCHOR.MIDDLE)
    text(slide, company_or_pp, Inches(9.0), Inches(0.32), Inches(3.98), Inches(0.4),
         size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT, font=BODY_SANS,
         anchor=MSO_ANCHOR.MIDDLE)


def footer(slide):
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.35), Inches(7.08), Inches(9.0), Inches(0.3),
         size=8, colour=RGBColor(0x66, 0x66, 0x66), font=BODY_SANS)
    text(slide, DATE_STR, Inches(10.6), Inches(7.08), Inches(2.38), Inches(0.3),
         size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.RIGHT, font=BODY_SANS)


def stat_card(slide, l, top, w, h, number, label_lines, source):
    rect(slide, l, top, w, h, BLACK)
    rect(slide, l, top, w, Pt(4), TEAL)
    text(slide, number, l + Inches(0.12), top + Inches(0.14), w - Inches(0.24), Inches(0.5),
         size=27, bold=True, colour=AMBER_B, font=HEAD_SERIF)
    text(slide, "\n".join(label_lines), l + Inches(0.12), top + Inches(0.68), w - Inches(0.24), Inches(0.55),
         size=10.5, colour=OFF_WHITE, font=BODY_SANS, line_spacing=1.05)
    text(slide, source, l + Inches(0.12), top + h - Inches(0.28), w - Inches(0.24), Inches(0.24),
         size=7.5, italic=True, colour=RGBColor(0x9a, 0x9a, 0x9a), font=BODY_SANS)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ═════════════════════════════════════════════════════════════════
# SLIDE 1
# ═════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, WHITE)
header_band(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

text(s1, COMPANY, Inches(0.35), Inches(1.12), Inches(9.5), Inches(0.68),
     size=40, bold=True, colour=BLACK, font=HEAD_SERIF)
text(s1, "National rental property compliance and maintenance platform, Burnley VIC",
     Inches(0.35), Inches(2.02), Inches(12.4), Inches(0.32),
     size=12, italic=True, colour=RGBColor(0x33, 0x33, 0x33), font=BODY_SANS)

cards = [
    ("$12.8M", ["FY25 revenue"], "Smart50 2025 award, Nov 2025"),
    ("72%", ["Revenue growth,", "FY23 to FY25"], "Smart50 2023 and 2025"),
    ("19", ["People on", "the team"], "Smart50 2025 award"),
    ("2014", ["Founded,", "Burnley VIC"], "Company website"),
    ("140K+", ["RentSafe jobs", "completed"], "Taskforce Australia site"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
start_x = Inches(0.6)
card_top = Inches(2.4)
card_h = Inches(1.6)
for i, (num, lbl, src) in enumerate(cards):
    l = start_x + i * (card_w + gap)
    stat_card(s1, l, card_top, card_w, card_h, num, lbl, src)

# Revenue trend chart, left half
chart_top = Inches(4.35)
text(s1, "REVENUE, THREE VERIFIED YEARS", Inches(0.6), chart_top, Inches(5.6), Inches(0.3),
     size=11, bold=True, colour=TEAL, font=BODY_SANS)
bar_base_y = Inches(6.55)
bar_area_h = Inches(1.7)
bar_vals = [("FY23", 7.43), ("FY24", 9.7), ("FY25", 12.8)]
max_val = 12.8
bar_w = Inches(0.9)
bar_gap = Inches(0.75)
bx = Inches(0.8)
for label, val in bar_vals:
    bh = Inches(bar_area_h.inches * (val / max_val))
    by = bar_base_y - bh
    rect(s1, bx, by, bar_w, bh, TEAL)
    text(s1, f"${val:.2f}M", bx - Inches(0.15), by - Inches(0.32), Inches(1.2), Inches(0.28),
         size=11, bold=True, colour=BLACK, font=BODY_SANS)
    text(s1, label, bx, bar_base_y + Inches(0.05), bar_w, Inches(0.28),
         size=10, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.CENTER, font=BODY_SANS)
    bx += bar_w + bar_gap
text(s1, "Source: SmartCompany Smart50 award citations, 2023, 2024 and 2025",
     Inches(0.6), Inches(6.9), Inches(5.6), Inches(0.25),
     size=7.5, italic=True, colour=RGBColor(0x77, 0x77, 0x77), font=BODY_SANS)

# Key commercial signals, right half
sig_x = Inches(6.6)
sig_w = Inches(6.35)
text(s1, "KEY COMMERCIAL SIGNALS", sig_x, chart_top, sig_w, Inches(0.3),
     size=11, bold=True, colour=TEAL, font=BODY_SANS)
signals = [
    "Ranked 37th nationally in Smart50 2025, its third straight Smart50 year since 2023. (SmartCompany)",
    "Revenue rose from $7.43M in FY23 to $12.8M in FY25 while the team held near 19 to 20 people. (Smart50 2023 to 2025)",
    "Named Victorian State Winner for Outstanding Growth, 2024 Telstra Best of Business Awards. (Telstra)",
    "RentSafe, launched 2021, has completed over 140,000 compliance jobs across 300 real estate offices. (Company site)",
    "Runs three distinct lines: RentSafe compliance, RentRepair maintenance, and manufacturer warranty servicing. (Company site)",
]
sy = chart_top + Inches(0.38)
for sline in signals:
    text(s1, "•  " + sline, sig_x, sy, sig_w, Inches(0.42),
         size=10, colour=RGBColor(0x22, 0x22, 0x22), font=BODY_SANS, line_spacing=1.0)
    sy += Inches(0.46)

footer(s1)

# ═════════════════════════════════════════════════════════════════
# SLIDE 2
# ═════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, WHITE)
header_band(s2, "THE OPPORTUNITY", "PROFITPULSE")
text(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.12), Inches(12.6), Inches(0.35),
     size=13, bold=True, colour=BLACK, font=BODY_SANS)

col_top = Inches(1.65)
col_h = Inches(4.75)
col_w = Inches(4.02)
col_gap = Inches(0.12)
col_fills = [TEAL, BLACK, GOLD]
col_text_colours = [WHITE, WHITE, BLACK]
col_headers = [
    "Two products, one blended margin",
    "Two very different customers",
    "Flat headcount, doubling revenue",
]
col_bodies = [
    "RentSafe compliance checks, RentRepair maintenance subscriptions, and manufacturer warranty servicing are priced and resourced differently, yet revenue is reported as one blended figure. Without a line by line margin view it is hard to tell which line is funding the growth and which is diluting it.",
    "RentSafe and RentRepair serve real estate agencies and property managers, who chase compliance and tenant satisfaction. Warranty servicing serves manufacturers, who chase fault rates and callback costs. These groups pay, churn and cost to serve very differently, and that unevenness is worth mapping.",
    "Revenue rose from $7.43 million in FY23 to $12.8 million in FY25 while the team held at 19 to 20 people, a genuine operating leverage story built on a network of over 5,500 tradespeople rather than headcount. The next stage of growth deserves a costed plan for the platform and working capital.",
]

for i in range(3):
    cx = Inches(0.35) + i * (col_w + col_gap)
    rect(s2, cx, col_top, col_w, col_h, col_fills[i])
    tc = col_text_colours[i]
    text(s2, f"0{i+1}", cx + Inches(0.22), col_top + Inches(0.2), col_w - Inches(0.4), Inches(0.6),
         size=30, bold=True, colour=tc, font=HEAD_SERIF)
    text(s2, col_headers[i], cx + Inches(0.22), col_top + Inches(0.95), col_w - Inches(0.44), Inches(0.75),
         size=15, bold=True, colour=tc, font=HEAD_SERIF, line_spacing=1.0)
    text(s2, col_bodies[i], cx + Inches(0.22), col_top + Inches(1.75), col_w - Inches(0.44), Inches(2.85),
         size=10.5, colour=tc, font=BODY_SANS, line_spacing=1.12)

text(s2,
     "These observations are offered in good faith. Taskforce has built a genuinely scaled national platform. "
     "The question is simply whether the margin picture is as clear as the growth curve.",
     Inches(0.35), Inches(6.5), Inches(12.6), Inches(0.5),
     size=10.5, italic=True, colour=RGBColor(0x33, 0x33, 0x33), font=BODY_SANS, align=PP_ALIGN.CENTER)

footer(s2)

# ═════════════════════════════════════════════════════════════════
# SLIDE 3
# ═════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)
header_band(s3, "THE RECOMMENDATION", "PROFITPULSE")

left_x = Inches(0.35)
left_w = Inches(7.55)

text(s3, "Product and Service Line Profitability",
     left_x, Inches(1.25), left_w, Inches(0.55),
     size=22, bold=True, colour=BLACK, font=HEAD_SERIF)
text(s3, "$3,950 one off", left_x, Inches(1.83), left_w, Inches(0.4),
     size=16, bold=True, colour=RGBColor(0xB8, 0x7A, 0x02), font=BODY_SANS)
text(s3,
     "A three week review ranking RentSafe, RentRepair, and warranty servicing by gross margin, "
     "contribution margin, and operational drag, so growth capital follows the line that earns it.",
     left_x, Inches(2.28), left_w, Inches(0.75),
     size=11.5, colour=RGBColor(0x22, 0x22, 0x22), font=BODY_SANS, line_spacing=1.1)

rect(s3, left_x, Inches(3.15), left_w, Inches(1.45), RGBColor(0xF2, 0xFA, 0xF9), line_colour=TEAL, line_w=Pt(1))
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.3), left_w - Inches(0.4), Inches(0.35),
     size=13, bold=True, colour=TEAL, font=BODY_SANS)
text(s3, "See the solutions matched to your size and industry.",
     left_x + Inches(0.2), Inches(3.68), left_w - Inches(0.4), Inches(0.3),
     size=11, colour=RGBColor(0x22, 0x22, 0x22), font=BODY_SANS)
find_fit_box = text(s3, "profit-pulse.com.au/services/find-your-fit",
     left_x + Inches(0.2), Inches(4.0), left_w - Inches(0.4), Inches(0.35),
     size=13, bold=True, colour=TEAL, font=BODY_SANS)
find_fit_box.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://profit-pulse.com.au/services/find-your-fit/"

rect(s3, left_x, Inches(4.75), left_w, Inches(0.65), AMBER_B)
cta_box = text(s3, "Purchase the suggested product now to get started",
     left_x, Inches(4.75), left_w, Inches(0.65),
     size=13, bold=True, colour=BLACK, align=PP_ALIGN.CENTER, font=BODY_SANS,
     anchor=MSO_ANCHOR.MIDDLE)
cta_box.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D"

text(s3, "Prefer a conversation first?", left_x, Inches(5.6), left_w, Inches(0.3),
     size=11, colour=RGBColor(0x22, 0x22, 0x22), font=BODY_SANS)
book_box = text(s3, "Book a complimentary discovery call",
     left_x, Inches(5.92), left_w, Inches(0.35),
     size=12, bold=True, colour=TEAL, font=BODY_SANS)
book_box.text_frame.paragraphs[0].runs[0].hyperlink.address = (
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right column, credibility panel
right_x = Inches(8.15)
right_w = Inches(4.83)
rect(s3, right_x, Inches(1.25), right_w, Inches(5.35), BLACK)
text(s3, "NITESH ROOPA", right_x + Inches(0.25), Inches(1.45), right_w - Inches(0.5), Inches(0.4),
     size=17, bold=True, colour=AMBER_B, font=HEAD_SERIF)
text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.85), right_w - Inches(0.5), Inches(0.35),
     size=12, colour=WHITE, font=BODY_SANS)
rect(s3, right_x + Inches(0.25), Inches(2.28), right_w - Inches(0.5), Pt(1.5), TEAL)

creds = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3B, Cahora Bassa",
    "Total GRBT project value over AUD 10B",
]
cy = Inches(2.45)
for cline in creds:
    text(s3, "•  " + cline, right_x + Inches(0.25), cy, right_w - Inches(0.5), Inches(0.32),
         size=10.5, colour=OFF_WHITE, font=BODY_SANS)
    cy += Inches(0.34)

rect(s3, right_x + Inches(0.25), cy + Inches(0.06), right_w - Inches(0.5), Pt(1), RGBColor(0x33, 0x33, 0x33))
cy += Inches(0.24)
contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for cl in contact_lines:
    text(s3, cl, right_x + Inches(0.25), cy, right_w - Inches(0.5), Inches(0.3),
         size=10.5, colour=OFF_WHITE, font=BODY_SANS)
    cy += Inches(0.32)

footer(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Aug2026.pptx"
prs.save(out_path)
print("PPTX saved:", out_path)
