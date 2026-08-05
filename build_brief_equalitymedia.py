"""
ProfitPulse Brief Builder, Version 3.3 house style
Target: Equality Media + Marketing | Date: 06 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes in copy.
White body, black header band, amber left stripe, teal/black/gold columns.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# Brand colours, Rule 3
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STAMP = "06 Aug 2026"
COMPANY = "Equality Media + Marketing"
COMPANY_SHORT = "EQUALITY MEDIA + MARKETING"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, font_name="Calibri",
             font_size=18, bold=False, colour=BLACK, align=PP_ALIGN.LEFT,
             italic=False, link=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    if link:
        run.hyperlink.address = link
    return box


def add_multiline(slide, lines, left, top, width, height, font_name="Calibri",
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   space_after=4, leading=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    first = True
    for line in lines:
        cfg = {"text": line, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False, "font": font_name} if isinstance(line, str) else {
            "text": line.get("text", ""), "size": line.get("size", default_size),
            "colour": line.get("colour", default_colour), "bold": line.get("bold", False),
            "italic": line.get("italic", False), "font": line.get("font", font_name)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        if leading:
            p.line_spacing = leading
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = cfg["font"]
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def add_chrome(slide, eyebrow, right_label):
    """Header band, left stripe, footer line. Shared on every slide."""
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(7.5), Inches(0.4),
              font_size=13, bold=True, colour=OFF_WHITE)
    add_text(slide, right_label, Inches(5.5), Inches(0.32), Inches(7.5), Inches(0.4),
              font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
             "Profit-Pulse.com.au", Inches(0.35), Inches(7.05), Inches(9.5), Inches(0.3),
              font_size=8, colour=BLACK)
    add_text(slide, DATE_STAMP, Inches(10.5), Inches(7.05), Inches(2.5), Inches(0.3),
              font_size=8, colour=BLACK, align=PP_ALIGN.RIGHT)
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)


def stat_card(slide, x, top, w, h, number, label_lines, source):
    add_rect(slide, x, top, w, h, BLACK)
    add_rect(slide, x, top, w, Pt(4), TEAL)
    add_text(slide, number, x + Inches(0.12), top + Inches(0.12), w - Inches(0.24), Inches(0.5),
              font_size=28, bold=True, colour=AMBER_B)
    add_multiline(slide, label_lines, x + Inches(0.12), top + Inches(0.66), w - Inches(0.24), Inches(0.55),
                   default_size=10, default_colour=OFF_WHITE, space_after=0)
    add_text(slide, source, x + Inches(0.12), top + h - Inches(0.3), w - Inches(0.24), Inches(0.26),
              font_size=7, colour=OFF_WHITE, italic=True)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ============================================================================
# SLIDE 1, COMMERCIAL INTELLIGENCE BRIEF
# ============================================================================
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
add_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", COMPANY_SHORT)

add_text(s1, COMPANY, Inches(0.35), Inches(1.2), Inches(11.5), Inches(0.7),
          font_name="Georgia", font_size=40, bold=True, colour=BLACK)
add_text(s1, "Independent full service advertising agency, Richmond, Melbourne VIC",
          Inches(0.35), Inches(1.85), Inches(11.5), Inches(0.35),
          font_size=13, colour=BLACK, italic=True)

cards = [
    ("$13M", ["Annual revenue,", "rank 24 nationally"], "SmartCompany Smart50 2025"),
    ("49%", ["Three year revenue", "growth rate"], "SmartCompany Smart50 2025"),
    ("30", ["People on the", "team"], "Mumbrella, 2026"),
    ("#55", ["AFR Fast 100", "2025 national rank"], "AFR Fast 100 list, 2025"),
    ("2018", ["Year the agency", "was founded"], "Smart50 2024 and 2025"),
]
card_w = Inches(2.28)
gap = Inches(0.15)
start_x = Inches(0.35)
card_top = Inches(2.4)
card_h = Inches(1.6)
for i, (num, label, src) in enumerate(cards):
    x = start_x + i * (card_w + gap)
    stat_card(s1, x, card_top, card_w, card_h, num, label, src)

# Chart, left half of remaining space: two verified revenue data points
add_text(s1, "REVENUE REPORTED AT AWARD TIME", Inches(0.35), Inches(4.35), Inches(5.5), Inches(0.3),
          font_size=11, bold=True, colour=TEAL)
chart_data = CategoryChartData()
chart_data.categories = ["Smart50 2024", "Smart50 2025"]
chart_data.add_series("Revenue $M", (7.3, 13.0))
gframe = s1.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.35), Inches(4.7), Inches(5.5), Inches(1.95), chart_data
)
chart = gframe.chart
chart.has_legend = False
plot = chart.plots[0]
plot.has_data_labels = True
plot.data_labels.font.size = Pt(11)
plot.data_labels.font.bold = True
plot.data_labels.font.color.rgb = BLACK
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = TEAL
cat_ax = chart.category_axis
cat_ax.tick_labels.font.size = Pt(10)
cat_ax.format.line.color.rgb = TEAL
val_ax = chart.value_axis
val_ax.visible = False
val_ax.has_major_gridlines = False
add_text(s1, "Source: SmartCompany Smart50 2024 and 2025 award citations",
          Inches(0.35), Inches(6.68), Inches(5.5), Inches(0.25), font_size=7,
          colour=BLACK, italic=True)

# Key Commercial Signals, right half
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(6.1), Inches(4.35), Inches(6.85), Inches(0.3),
          font_size=11, bold=True, colour=TEAL)
signals = [
    "Ranked 24th on the 2025 Smart50 list, up from 45th in 2024. (SmartCompany)",
    "Also placed 55th nationally on the AFR Fast 100 2025 growth list. (AFR Fast 100)",
    "Appointed ANZ media agency of record for skincare brand BIODERMA in 2025. (AdNews, B&T)",
    "Named 2026 AFR BOSS Best Place to Work winner, Media and Marketing category. (Mumbrella)",
    "Team has grown from 17 people in 2024 toward 30 by 2026. (SmartCompany, Mumbrella)",
    "Runs a four day, full pay 32 hour work week introduced in 2022. (Mumbrella, AdNews)",
]
add_multiline(s1, [{"text": "•  " + s, "size": 11, "colour": BLACK} for s in signals],
               Inches(6.1), Inches(4.68), Inches(6.85), Inches(2.15), default_size=11,
               space_after=8, leading=1.05)

# ============================================================================
# SLIDE 2, THE OPPORTUNITY
# ============================================================================
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
add_chrome(s2, "THE OPPORTUNITY", COMPANY_SHORT)

add_text(s2, "Equality Media + Marketing: three commercial observations from ProfitPulse",
          Inches(0.35), Inches(1.08), Inches(12.3), Inches(0.35),
          font_size=13, bold=True, colour=TEAL)

col_top = Inches(1.55)
col_h = Inches(4.75)
col_w = Inches(4.03)
col_gap = Inches(0.12)
col_x = [Inches(0.35), Inches(0.35) + col_w + col_gap, Inches(0.35) + 2 * (col_w + col_gap)]
col_fills = [TEAL, BLACK, GOLD]
col_text = [BLACK, OFF_WHITE, BLACK]

observations = [
    ("01", "Growth is outrunning the finance function",
     "Revenue reported at Smart50 award time moved from 7.3 million dollars in 2024 to "
     "13 million in 2025, while the team has grown toward 30 people. That pace of hiring "
     "and client onboarding is a common point where cash timing, not profit, becomes the "
     "real constraint on how fast a business can safely take on new work."),
    ("02", "Media buying carries a working capital load",
     "Full service planning and buying, including the new ANZ mandate for BIODERMA, "
     "typically means paying media suppliers ahead of client settlement. As billings "
     "scale, the gap between paying media costs and collecting client revenue widens, "
     "and can quietly absorb cash the business assumes it still has on hand."),
    ("03", "A four day week raises the value of every hour",
     "Equality Time has delivered 85 percent staff retention, 14 percent growth in gross "
     "profit and a 23 percent rise in client work. Running full pay on a 32 hour week "
     "means each hour of capacity carries more weight, which makes a live view of where "
     "time turns into margin as important as the cash discipline needed to fund growth."),
]

for i, x in enumerate(col_x):
    add_rect(s2, x, col_top, col_w, col_h, col_fills[i])
    idx, header, para = observations[i]
    add_text(s2, idx, x + Inches(0.2), col_top + Inches(0.2), col_w - Inches(0.4), Inches(0.6),
              font_name="Georgia", font_size=30, bold=True, colour=col_text[i])
    add_text(s2, header, x + Inches(0.2), col_top + Inches(0.85), col_w - Inches(0.4), Inches(0.75),
              font_size=15, bold=True, colour=col_text[i])
    add_multiline(s2, [{"text": para, "size": 11, "colour": col_text[i]}],
                   x + Inches(0.2), col_top + Inches(1.7), col_w - Inches(0.4), Inches(2.9),
                   default_size=11, leading=1.15)

add_text(s2, "These observations are offered in good faith. Equality Media + Marketing has "
             "built something genuinely impressive. The question is simply whether the "
             "financial architecture keeps pace with the growth.",
          Inches(0.35), Inches(6.42), Inches(12.6), Inches(0.5),
          font_size=11, italic=True, colour=BLACK)

# ============================================================================
# SLIDE 3, THE RECOMMENDATION AND HOW TO START
# ============================================================================
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
add_chrome(s3, "THE RECOMMENDATION", COMPANY_SHORT)

left_x = Inches(0.35)
left_w = Inches(7.5)

add_text(s3, "Working Capital Unlock", left_x, Inches(1.2), left_w, Inches(0.55),
          font_name="Georgia", font_size=26, bold=True, colour=BLACK)
add_text(s3, "$4,450 one off, ProfitPulse verified price", left_x, Inches(1.78), left_w, Inches(0.35),
          font_size=14, bold=True, colour=TEAL)
add_multiline(s3, [
    {"text": "Maps cash trapped in debtors, supplier payment timing and stock or work in "
             "progress across the business, then delivers a prioritised action list to "
             "release cash within weeks.", "size": 12, "colour": BLACK},
], left_x, Inches(2.2), left_w, Inches(0.85), default_size=12, leading=1.15)

# Step one block
add_rect(s3, left_x, Inches(3.15), left_w, Inches(1.5), OFF_WHITE,
          line_colour=TEAL, line_width=Pt(1))
add_text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.32),
          left_w - Inches(0.4), Inches(0.35), font_size=13, bold=True, colour=BLACK)
add_text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.2),
          Inches(3.68), left_w - Inches(0.4), Inches(0.35), font_size=11, colour=BLACK)
add_text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.2), Inches(4.05),
          left_w - Inches(0.4), Inches(0.4), font_size=13, bold=True, colour=TEAL,
          link="https://profit-pulse.com.au/services/find-your-fit/")

# Direct purchase CTA
add_rect(s3, left_x, Inches(4.85), left_w, Inches(0.7), AMBER_D)
add_text(s3, "Purchase the suggested product now to get started", left_x + Inches(0.2), Inches(5.06),
          left_w - Inches(0.4), Inches(0.35), font_size=13, bold=True, colour=BLACK,
          align=PP_ALIGN.CENTER, link="https://buy.stripe.com/14AaEYb163Y4a9q4c73ks0y")

add_text(s3, "Prefer a conversation first?", left_x, Inches(5.75), left_w, Inches(0.3),
          font_size=12, bold=True, colour=BLACK)
add_text(s3, "Book a complimentary discovery call", left_x, Inches(6.08), left_w, Inches(0.35),
          font_size=12, colour=TEAL,
          link="https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right column, credibility panel
right_x = Inches(8.15)
right_w = Inches(4.85)
add_rect(s3, right_x, Inches(1.2), right_w, Inches(5.35), BLACK)
add_rect(s3, right_x, Inches(1.2), right_w, Pt(4), TEAL)
add_text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.42), right_w - Inches(0.5), Inches(0.45),
          font_name="Georgia", font_size=19, bold=True, colour=AMBER_B)
add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.85),
          right_w - Inches(0.5), Inches(0.35), font_size=12, colour=WHITE)

cred_lines = [
    {"text": "16 years across 4 countries", "size": 11, "colour": OFF_WHITE},
    {"text": "52 deals executed and managed", "size": 11, "colour": OFF_WHITE},
    {"text": "Largest single deal, USD 1.3 billion,", "size": 11, "colour": OFF_WHITE},
    {"text": "Cahora Bassa, Mozambique Government", "size": 11, "colour": OFF_WHITE},
    {"text": "Total GRBT project value over AUD 10 billion", "size": 11, "colour": OFF_WHITE},
]
add_multiline(s3, cred_lines, right_x + Inches(0.25), Inches(2.35), right_w - Inches(0.5), Inches(1.5),
               default_size=11, space_after=5, leading=1.1)

add_rect(s3, right_x + Inches(0.25), Inches(4.0), right_w - Inches(0.5), Pt(1.5), TEAL)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 12, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 12, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10, "colour": OFF_WHITE},
]
add_multiline(s3, contact_lines, right_x + Inches(0.25), Inches(4.2), right_w - Inches(0.5), Inches(1.6),
               default_size=12, space_after=8, leading=1.1)

# ============================================================================
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMediaMarketing_06Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
