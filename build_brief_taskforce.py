"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 27 Jun 2026
Three-slide prospect-facing deck. White background house style per Section 6.
Brand colours only. Zero dashes. No tier names.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import copy

# ── Brand colours ─────────────────────────────────────────────────────────────
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

# Slide dimensions: widescreen 13.333" x 7.5"
W = Inches(13.333)
H = Inches(7.5)

# Header height and amber stripe width (house style)
HEADER_H = Inches(1.0)
STRIPE_W = Inches(0.10)
FOOTER_Y = Inches(7.15)


def set_background(slide, colour):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width_pt=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        if line_width_pt:
            shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height,
             font_name="Calibri", font_size=12, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT,
             wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return txBox


def add_multiline(slide, lines, left, top, width, height,
                  font_name="Calibri", default_size=12,
                  default_colour=BLACK, default_bold=False,
                  align=PP_ALIGN.LEFT, spacing_after=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size,
                   "colour": default_colour, "bold": default_bold, "italic": False}
        else:
            cfg = {
                "text":   line.get("text", ""),
                "size":   line.get("size", default_size),
                "colour": line.get("colour", default_colour),
                "bold":   line.get("bold", default_bold),
                "italic": line.get("italic", False),
            }
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if spacing_after:
            p.space_after = Pt(spacing_after)
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return txBox


def add_hyperlink_text(slide, link_text, url, left, top, width, height,
                       font_name="Calibri", font_size=12, bold=False,
                       colour=TEAL, align=PP_ALIGN.LEFT):
    """Add a text box with a hyperlink run."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = link_text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.hyperlink.address = url
    return txBox


def house_header(slide, eyebrow_label):
    """Draw the standard house style header on a slide."""
    # Left amber stripe
    add_rect(slide, Inches(0), Inches(0), STRIPE_W, H, AMBER_D)
    # Black header band
    add_rect(slide, STRIPE_W, Inches(0), W - STRIPE_W, HEADER_H, BLACK)
    # Eyebrow label
    add_text(slide, eyebrow_label,
             left=Inches(0.22), top=Inches(0.18), width=Inches(8), height=Inches(0.55),
             font_name="Calibri", font_size=13, bold=True,
             colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    # PROFITPULSE on right
    add_text(slide, "PROFITPULSE",
             left=Inches(10.5), top=Inches(0.18), width=Inches(2.7), height=Inches(0.55),
             font_name="Calibri", font_size=13, bold=True,
             colour=TEAL, align=PP_ALIGN.RIGHT)


def house_footer(slide, date_str="27 Jun 2026"):
    """Draw the standard house style footer."""
    # Thin separator line in teal
    add_rect(slide, STRIPE_W, Inches(7.05), W - STRIPE_W, Pt(1), TEAL)
    add_text(slide,
             "Prepared by Nitesh Roopa CA  |  Managing Partner  |  ProfitPulse  |  Profit-Pulse.com.au",
             left=Inches(0.22), top=FOOTER_Y, width=Inches(9.5), height=Inches(0.3),
             font_name="Calibri", font_size=8, bold=False,
             colour=TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, date_str,
             left=Inches(10.5), top=FOOTER_Y, width=Inches(2.7), height=Inches(0.3),
             font_name="Calibri", font_size=8, bold=False,
             colour=TEAL, align=PP_ALIGN.RIGHT)


def stat_card(slide, left, top, width, height, number, label_line1, label_line2, source):
    """Draw a stat card: black tile, teal top stripe, large number, label, source."""
    # Black tile
    add_rect(slide, left, top, width, height, BLACK)
    # Teal top stripe (thin)
    add_rect(slide, left, top, width, Pt(4), TEAL)
    # Large number
    add_text(slide, number,
             left=left + Inches(0.07), top=top + Inches(0.1), width=width - Inches(0.14), height=Inches(0.6),
             font_name="Georgia", font_size=26, bold=True,
             colour=AMBER_B, align=PP_ALIGN.LEFT)
    # Label lines
    label = label_line1 + ("\n" + label_line2 if label_line2 else "")
    add_text(slide, label,
             left=left + Inches(0.07), top=top + Inches(0.72), width=width - Inches(0.14), height=Inches(0.45),
             font_name="Calibri", font_size=9, bold=False,
             colour=OFF_WHITE, align=PP_ALIGN.LEFT, wrap=True)
    # Source line
    add_text(slide, source,
             left=left + Inches(0.07), top=top + Inches(1.18), width=width - Inches(0.14), height=Inches(0.22),
             font_name="Calibri", font_size=7, bold=False,
             colour=TEAL, align=PP_ALIGN.LEFT, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD PRESENTATION
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank_layout = prs.slide_layouts[6]


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank_layout)
set_background(s1, WHITE)
house_header(s1, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name (large black serif)
add_text(s1, "TASKFORCE AUSTRALIA",
         left=Inches(0.22), top=Inches(1.08), width=Inches(10), height=Inches(0.8),
         font_name="Georgia", font_size=44, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

# Descriptor
add_text(s1, "Field service management platform and property safety compliance, Burnley, Melbourne VIC",
         left=Inches(0.22), top=Inches(1.85), width=Inches(10.5), height=Inches(0.35),
         font_name="Calibri", font_size=12, bold=False,
         colour=BLACK, align=PP_ALIGN.LEFT)

# ── Stat cards row ──────────────────────────────────────────────────────────
CARD_W = Inches(2.1)
CARD_H = Inches(1.5)
CARD_Y = Inches(2.28)
CARD_GAP = Inches(0.07)
card_left = Inches(0.22)

cards = [
    ("$12.8M",  "Revenue",         "FY2025",          "SmartCompany Smart50 2025, rank 37"),
    ("31%",     "Three year average",  "revenue growth",  "SmartCompany Smart50 2025"),
    ("19",      "Team members",    "",                "SmartCompany Smart50 2025"),
    ("5,500",   "Tradespeople",    "in platform network", "SmartCompany Smart50 2024"),
    ("140K+",   "Jobs delivered",  "via RentSafe since 2021", "SmartCompany Smart50 2025"),
    ("2014",    "Year founded",    "",                "Company records"),
]

for i, (num, l1, l2, src) in enumerate(cards):
    x = card_left + i * (CARD_W + CARD_GAP)
    stat_card(s1, x, CARD_Y, CARD_W, CARD_H, num, l1, l2, src)

# ── Revenue bar chart (manual bars, two verified data points) ──────────────
chart_label_y = Inches(3.95)
chart_left = Inches(0.22)
chart_w_total = Inches(3.5)
chart_h = Inches(0.9)
chart_y = Inches(4.12)
bar_w = Inches(0.9)
gap_bars = Inches(0.15)

# Teal section label
add_text(s1, "REVENUE GROWTH  (AUD, verified)",
         left=chart_left, top=Inches(3.87), width=Inches(5), height=Inches(0.22),
         font_name="Calibri", font_size=9, bold=True,
         colour=TEAL, align=PP_ALIGN.LEFT)

# 2024 bar (height proportional: 9.7 / 12.8)
bar2024_h = chart_h * (9.7 / 12.8)
bar2024_y = chart_y + (chart_h - bar2024_h)
add_rect(s1, chart_left, bar2024_y, bar_w, bar2024_h, TEAL)
add_text(s1, "$9.7M",
         left=chart_left, top=bar2024_y - Inches(0.22),
         width=bar_w, height=Inches(0.2),
         font_name="Calibri", font_size=8, bold=True,
         colour=BLACK, align=PP_ALIGN.CENTER)
add_text(s1, "FY2024",
         left=chart_left, top=chart_y + chart_h + Inches(0.02),
         width=bar_w, height=Inches(0.18),
         font_name="Calibri", font_size=8, bold=False,
         colour=BLACK, align=PP_ALIGN.CENTER)

# 2025 bar (full height)
bar2025_x = chart_left + bar_w + gap_bars
add_rect(s1, bar2025_x, chart_y, bar_w, chart_h, AMBER_D)
add_text(s1, "$12.8M",
         left=bar2025_x, top=chart_y - Inches(0.22),
         width=bar_w, height=Inches(0.2),
         font_name="Calibri", font_size=8, bold=True,
         colour=BLACK, align=PP_ALIGN.CENTER)
add_text(s1, "FY2025",
         left=bar2025_x, top=chart_y + chart_h + Inches(0.02),
         width=bar_w, height=Inches(0.18),
         font_name="Calibri", font_size=8, bold=False,
         colour=BLACK, align=PP_ALIGN.CENTER)

add_text(s1, "Source: SmartCompany Smart50 2024 and 2025 award citations",
         left=chart_left, top=chart_y + chart_h + Inches(0.23),
         width=Inches(3.5), height=Inches(0.2),
         font_name="Calibri", font_size=7, bold=False,
         colour=TEAL, align=PP_ALIGN.LEFT, italic=True)

# ── Key Commercial Signals ──────────────────────────────────────────────────
sig_x = Inches(4.0)
sig_w = Inches(9.0)

add_text(s1, "KEY COMMERCIAL SIGNALS",
         left=sig_x, top=Inches(3.87), width=sig_w, height=Inches(0.22),
         font_name="Calibri", font_size=9, bold=True,
         colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    ("RentRepair launched as Australia's only complete subscription property maintenance service. Two tiers (Essential and Comprehensive). (taskforce.com.au/rentrepair)", True),
    ("140,000+ jobs delivered through RentSafe across 20 major consumer brands since 2021. Brands include Brilliant Lighting, Nero, Universal Fans. (SmartCompany Smart50 2025)", False),
    ("Agency clients doubled in 12 months. Over 300 real estate offices and 180 major real estate brands now using the platform. (SmartCompany Smart50 2025)", False),
    ("Three consecutive Smart50 appearances: 2023, 2024 rank 46, 2025 rank 37. Revenue climbed from $9.7M to $12.8M year on year. (SmartCompany)", True),
    ("Won Most Innovative Proptech at the 2024 Proptech Awards, Sydney. (PropTech Australia 2024 citation)", False),
]

sig_line_y = Inches(4.15)
for text, bold in signals:
    add_text(s1, "•  " + text,
             left=sig_x, top=sig_line_y, width=sig_w, height=Inches(0.28),
             font_name="Calibri", font_size=9.5, bold=bold,
             colour=BLACK, align=PP_ALIGN.LEFT, wrap=True)
    sig_line_y += Inches(0.32)

house_footer(s1)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE OPPORTUNITY
# ────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank_layout)
set_background(s2, WHITE)
house_header(s2, "THE OPPORTUNITY")

# Subtitle
add_text(s2, "Taskforce Australia: Three commercial observations from ProfitPulse",
         left=Inches(0.22), top=HEADER_H + Inches(0.08), width=Inches(12.8), height=Inches(0.33),
         font_name="Calibri", font_size=13, bold=False,
         colour=BLACK, align=PP_ALIGN.LEFT)

# Three observation columns
COL_Y = HEADER_H + Inches(0.47)
COL_H = Inches(5.4)
COL_W = Inches(4.22)
COL_GAP = Inches(0.04)

obs_data = [
    {
        "bg": TEAL,
        "text_colour": WHITE,
        "num_colour": WHITE,
        "index": "01",
        "header": "Revenue per person is exceptional. Margin per client is unknown.",
        "body": (
            "At $12.8 million across 19 employees, Taskforce generates "
            "$674,000 in revenue per person, placing it among the most "
            "efficient field service businesses in Australia. But the "
            "platform coordinates more than 5,500 external tradespeople, "
            "meaning a significant share of revenue flows through to "
            "subcontractor fees. Until each manufacturer client and each "
            "real estate office is ranked by gross margin contribution "
            "and coordination cost, growth is being managed by revenue "
            "alone. A Customer Concentration and Profitability Map reveals "
            "which accounts to grow, which to reprice, and which consume "
            "more than they return."
        ),
    },
    {
        "bg": BLACK,
        "text_colour": OFF_WHITE,
        "num_colour": AMBER_B,
        "index": "02",
        "header": "Two channels with different economics need separate financial discipline.",
        "body": (
            "RentSafe manages compliance for manufacturers and real estate "
            "offices on a per job model. RentRepair offers subscription "
            "maintenance at a fixed monthly fee. These are structurally "
            "different businesses within one entity: one is volume and "
            "frequency driven, the other is recurring revenue with customer "
            "lifetime value mechanics. A business that prices both against "
            "the same cost base will misread its margins in both directions. "
            "Understanding which channel funds growth and which consumes it "
            "is the foundational question before the next expansion phase."
        ),
    },
    {
        "bg": GOLD,
        "text_colour": BLACK,
        "num_colour": BLACK,
        "index": "03",
        "header": "Doubling agency clients signals growth. Cash timing needs matching attention.",
        "body": (
            "Doubling agency clients in a year is a genuine commercial "
            "achievement. It also means more work in progress, more debtor "
            "cycles, and more subcontractor payments ahead of client "
            "settlement. As Taskforce moves into subscription billing "
            "through RentRepair, the cash conversion model changes again. "
            "A disciplined working capital review and a 13-week cash flow "
            "build are the natural next steps for a business that has "
            "outgrown its original revenue model and is building a second "
            "engine alongside the first."
        ),
    },
]

for i, obs in enumerate(obs_data):
    cx = Inches(0.13) + i * (COL_W + COL_GAP)
    add_rect(s2, cx, COL_Y, COL_W, COL_H, obs["bg"])

    # Index number
    add_text(s2, obs["index"],
             left=cx + Inches(0.18), top=COL_Y + Inches(0.15),
             width=COL_W - Inches(0.36), height=Inches(0.6),
             font_name="Georgia", font_size=36, bold=True,
             colour=obs["num_colour"], align=PP_ALIGN.LEFT)

    # Observation header
    add_text(s2, obs["header"],
             left=cx + Inches(0.18), top=COL_Y + Inches(0.78),
             width=COL_W - Inches(0.36), height=Inches(0.75),
             font_name="Georgia", font_size=12, bold=True,
             colour=obs["text_colour"], align=PP_ALIGN.LEFT)

    # Body text
    add_text(s2, obs["body"],
             left=cx + Inches(0.18), top=COL_Y + Inches(1.57),
             width=COL_W - Inches(0.36), height=Inches(3.6),
             font_name="Calibri", font_size=10.5, bold=False,
             colour=obs["text_colour"], align=PP_ALIGN.LEFT, wrap=True)

# Warm closing line
add_text(s2,
         "These observations are offered in good faith. Taskforce Australia "
         "has built a remarkable platform. The question is simply whether the "
         "financial architecture is keeping pace with the commercial ambition.",
         left=Inches(0.22), top=COL_Y + COL_H + Inches(0.1),
         width=Inches(12.8), height=Inches(0.32),
         font_name="Calibri", font_size=10, bold=False, italic=True,
         colour=BLACK, align=PP_ALIGN.CENTER)

house_footer(s2)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE RECOMMENDATION
# ────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank_layout)
set_background(s3, WHITE)
house_header(s3, "THE RECOMMENDATION")

LEFT_COL_X = Inches(0.22)
LEFT_COL_W = Inches(7.5)
RIGHT_COL_X = Inches(8.0)
RIGHT_COL_W = Inches(5.1)
CONTENT_Y_START = HEADER_H + Inches(0.18)

# ── Left column ──────────────────────────────────────────────────────────────
# Service name and price
add_text(s3, "Customer Concentration and Profitability Map",
         left=LEFT_COL_X, top=CONTENT_Y_START,
         width=LEFT_COL_W, height=Inches(0.65),
         font_name="Georgia", font_size=20, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

add_text(s3, "$3,950  one off  |  ProfitPulse verified price",
         left=LEFT_COL_X, top=CONTENT_Y_START + Inches(0.65),
         width=LEFT_COL_W, height=Inches(0.3),
         font_name="Calibri", font_size=13, bold=False,
         colour=TEAL, align=PP_ALIGN.LEFT)

# What it does
add_text(s3,
         "A three week project ranking every customer by revenue, gross margin "
         "contribution, and effort to serve. For Taskforce Australia, this maps "
         "both manufacturer clients and real estate agency accounts to reveal "
         "which drive the real margin after tradie fees, platform costs, and "
         "coordination effort are allocated. Output: a clear action list of "
         "which accounts to grow, reprice, or reset.",
         left=LEFT_COL_X, top=CONTENT_Y_START + Inches(1.03),
         width=LEFT_COL_W, height=Inches(1.1),
         font_name="Calibri", font_size=10.5, bold=False,
         colour=BLACK, align=PP_ALIGN.LEFT, wrap=True)

# Teal separator
add_rect(s3, LEFT_COL_X, CONTENT_Y_START + Inches(2.2), LEFT_COL_W, Pt(2), TEAL)

# Step one block
STEP_Y = CONTENT_Y_START + Inches(2.28)
add_text(s3, "STEP ONE",
         left=LEFT_COL_X, top=STEP_Y,
         width=LEFT_COL_W, height=Inches(0.22),
         font_name="Calibri", font_size=9, bold=True,
         colour=TEAL, align=PP_ALIGN.LEFT)

add_text(s3, "Answer a few quick questions",
         left=LEFT_COL_X, top=STEP_Y + Inches(0.22),
         width=LEFT_COL_W, height=Inches(0.32),
         font_name="Calibri", font_size=14, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

add_text(s3, "See the solutions matched to your size and industry",
         left=LEFT_COL_X, top=STEP_Y + Inches(0.55),
         width=LEFT_COL_W, height=Inches(0.25),
         font_name="Calibri", font_size=11, bold=False,
         colour=BLACK, align=PP_ALIGN.LEFT)

# Questionnaire address shown as clean text with hyperlink to clean address (no UTM on brief)
add_hyperlink_text(
    s3,
    "profit-pulse.com.au/full-suite-of-products",
    "https://profit-pulse.com.au/full-suite-of-products",
    left=LEFT_COL_X, top=STEP_Y + Inches(0.82),
    width=LEFT_COL_W, height=Inches(0.3),
    font_name="Calibri", font_size=11, bold=False,
    colour=TEAL
)

# Amber separator
add_rect(s3, LEFT_COL_X, STEP_Y + Inches(1.17), LEFT_COL_W, Pt(1.5), AMBER_D)

# Direct CTA (Stripe link hidden behind words)
add_hyperlink_text(
    s3,
    "Purchase the suggested product now to get started",
    "https://buy.stripe.com/14AbJ21qw2U0ftK0ZV3ks1A",
    left=LEFT_COL_X, top=STEP_Y + Inches(1.22),
    width=LEFT_COL_W, height=Inches(0.32),
    font_name="Calibri", font_size=12, bold=True,
    colour=AMBER_D
)

# Booking link
add_text(s3, "Prefer a conversation first?",
         left=LEFT_COL_X, top=STEP_Y + Inches(1.63),
         width=LEFT_COL_W, height=Inches(0.25),
         font_name="Calibri", font_size=10, bold=False,
         colour=BLACK, align=PP_ALIGN.LEFT)

add_hyperlink_text(
    s3,
    "Book a complimentary discovery call",
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
    left=LEFT_COL_X, top=STEP_Y + Inches(1.9),
    width=LEFT_COL_W, height=Inches(0.28),
    font_name="Calibri", font_size=11, bold=False,
    colour=TEAL
)

# ── Right column: credibility panel ─────────────────────────────────────────
# Off-white background for the right panel with teal border
add_rect(s3, RIGHT_COL_X, CONTENT_Y_START, RIGHT_COL_W, Inches(5.62),
         OFF_WHITE, line_colour=TEAL, line_width_pt=0.75)

# Teal top accent on panel
add_rect(s3, RIGHT_COL_X, CONTENT_Y_START, RIGHT_COL_W, Pt(4), TEAL)

PX = RIGHT_COL_X + Inches(0.22)
PW = RIGHT_COL_W - Inches(0.44)

add_text(s3, "Nitesh Roopa",
         left=PX, top=CONTENT_Y_START + Inches(0.18),
         width=PW, height=Inches(0.5),
         font_name="Georgia", font_size=22, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

add_text(s3, "CA, Managing Partner  |  ProfitPulse",
         left=PX, top=CONTENT_Y_START + Inches(0.68),
         width=PW, height=Inches(0.3),
         font_name="Calibri", font_size=12, bold=False,
         colour=TEAL, align=PP_ALIGN.LEFT)

add_rect(s3, PX, CONTENT_Y_START + Inches(1.03), PW, Pt(1.5), TEAL)

cred_points = [
    "16 years across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion (Cahora Bassa, Mozambique Government, Hydro)",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "CA qualification: SAICA, South Africa",
    "Practice: PwC South Africa, Nedbank CIB, QIC, ProfitPulse",
]

cy = CONTENT_Y_START + Inches(1.12)
for pt_text in cred_points:
    add_text(s3, "•  " + pt_text,
             left=PX, top=cy, width=PW, height=Inches(0.32),
             font_name="Calibri", font_size=9.5, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, wrap=True)
    cy += Inches(0.34)

# Contact block
add_rect(s3, PX, cy + Inches(0.06), PW, Pt(1.5), AMBER_D)
add_multiline(s3, [
    {"text": "Profit-Pulse.com.au",         "size": 10, "colour": BLACK, "bold": False},
    {"text": "Nitesh@Profit-Pulse.com.au",  "size": 10, "colour": TEAL,  "bold": False},
    {"text": "+61 411 876 267",             "size": 10, "colour": BLACK, "bold": False},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 9, "colour": BLACK, "bold": False},
],
    left=PX, top=cy + Inches(0.14), width=PW, height=Inches(1.1),
    default_size=10, default_colour=BLACK, spacing_after=1)

house_footer(s3)


# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_27Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
