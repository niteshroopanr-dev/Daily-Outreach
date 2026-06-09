"""
ProfitPulse Brief Builder
Target: Paire | Date: 10 Jun 2026
Three-slide prospect-facing deck. Brand colours only. Zero dashes.
House style: white body background, black header band, amber left stripe.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Brand colours ──────────────────────────────────────────────────────────────
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY  = RGBColor(0x88, 0x88, 0x88)
DRK_GREY  = RGBColor(0x44, 0x44, 0x44)

W = Inches(13.333)
H = Inches(7.5)


def set_background(slide, colour):
    bg = slide.background
    fill = bg.fill
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
    return shape


def add_text(slide, text, left, top, width, height,
             font_name="Arial", font_size=12, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
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


def add_text_with_link(slide, text, url, left, top, width, height,
                       font_name="Arial", font_size=12, bold=False,
                       colour=TEAL, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    rId = slide.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    rPr = run._r.get_or_add_rPr()
    hlinkClick = etree.SubElement(rPr, qn("a:hlinkClick"))
    hlinkClick.set(qn("r:id"), rId)
    return txBox


def add_multiline(slide, lines, left, top, width, height,
                  font_name="Arial", default_size=12,
                  default_colour=BLACK, align=PP_ALIGN.LEFT, spacing_after=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size, "colour": default_colour,
                   "bold": False, "italic": False}
        else:
            cfg = {
                "text":   line.get("text", ""),
                "size":   line.get("size", default_size),
                "colour": line.get("colour", default_colour),
                "bold":   line.get("bold", False),
                "italic": line.get("italic", False),
            }
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
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


def add_stat_card(slide, x, y, w, h, number, label1, label2, source):
    """Black stat card with thin teal top stripe, white number, off-white label, muted source."""
    add_rect(slide, x, y, w, h, BLACK)
    add_rect(slide, x, y, w, Inches(0.06), TEAL)
    add_text(slide, number,
             x + Inches(0.12), y + Inches(0.1), w - Inches(0.15), Inches(0.55),
             font_size=26, bold=True, colour=WHITE)
    add_multiline(slide,
                  [{"text": label1, "size": 10, "colour": OFF_WHITE, "bold": False},
                   {"text": label2, "size": 10, "colour": OFF_WHITE, "bold": False}],
                  x + Inches(0.12), y + Inches(0.66), w - Inches(0.15), Inches(0.5),
                  default_size=10, default_colour=OFF_WHITE, spacing_after=1)
    add_text(slide, source,
             x + Inches(0.12), y + Inches(1.18), w - Inches(0.15), Inches(0.22),
             font_size=7, bold=False, colour=MID_GREY, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
# PRESENTATION SETUP
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]

DATE_LABEL  = "10 Jun 2026"
FOOTER_LEFT = "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au"
COMPANY     = "Paire"


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)

# Left amber accent stripe
add_rect(s1, Inches(0), Inches(0), Inches(0.08), H, AMBER_B)

# Black header band
add_rect(s1, Inches(0.08), Inches(0), W - Inches(0.08), Inches(0.95), BLACK)
add_text(s1, "COMMERCIAL INTELLIGENCE BRIEF",
         Inches(0.22), Inches(0.22), Inches(8), Inches(0.5),
         font_size=11, bold=True, colour=OFF_WHITE)
add_text(s1, COMPANY,
         Inches(0.22), Inches(0.22), W - Inches(0.44), Inches(0.5),
         font_size=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Company name and descriptor
add_text(s1, "Paire",
         Inches(0.22), Inches(1.05), Inches(9), Inches(0.75),
         font_name="Georgia", font_size=44, bold=True, colour=BLACK)
add_text(s1, "Sustainable direct to consumer apparel brand, South Melbourne VIC",
         Inches(0.22), Inches(1.83), Inches(10), Inches(0.4),
         font_size=12, colour=TEAL)

# Section label
add_text(s1, "KEY METRICS",
         Inches(0.22), Inches(2.3), Inches(4), Inches(0.3),
         font_size=11, bold=True, colour=TEAL)

# 5 Stat cards (card w=2.46, gap=0.13, start x=0.22)
CW = Inches(2.46)
CH = Inches(1.42)
CY = Inches(2.64)
CX = [Inches(0.22), Inches(2.81), Inches(5.40), Inches(7.99), Inches(10.58)]

add_stat_card(s1, CX[0], CY, CW, CH,
              "$10M", "Revenue FY2025", "",
              "SmartCompany Smart50 2025")

add_stat_card(s1, CX[1], CY, CW, CH,
              "64%", "Revenue growth", "three year average",
              "SmartCompany Smart50 2025")

add_stat_card(s1, CX[2], CY, CW, CH,
              "26", "Employees", "16 FTE plus 10 casual",
              "SmartCompany Jul 2025")

add_stat_card(s1, CX[3], CY, CW, CH,
              "15", "International wholesale", "stores: Singapore and Malaysia",
              "businessnewsaustralia.com 2025")

add_stat_card(s1, CX[4], CY, CW, CH,
              "70%", "Repeat customer rate", "brand loyalty metric",
              "SmartCompany 2025")

# ── Revenue trajectory chart (left column below cards) ───────────────────────
add_text(s1, "REVENUE TRAJECTORY",
         Inches(0.22), Inches(4.2), Inches(5.0), Inches(0.3),
         font_size=11, bold=True, colour=TEAL)

# Chart baseline
CHART_X     = Inches(0.55)
CHART_Y_BOT = Inches(6.4)
CHART_H_MAX = Inches(1.75)
BAR_W       = Inches(0.65)

bars = [
    ("FY21", 1.0,  "$1M",   "SmartCompany"),
    ("FY22", 1.7,  "$1.7M", "SmartCompany"),
    ("FY24", 6.0,  "$6M",   "SmartCompany Shark Tank recap Nov 2024"),
    ("FY25", 10.0, "$10M",  "Smart50 2025"),
]
BAR_SPACING = Inches(1.10)

for i, (yr, val, label, src) in enumerate(bars):
    bx = CHART_X + Inches(i * 1.1)
    bh = Inches(val / 10.0 * 1.75)
    by = CHART_Y_BOT - bh
    add_rect(s1, bx, by, BAR_W, bh, TEAL)
    # Value label above bar
    add_text(s1, label,
             bx - Inches(0.05), by - Inches(0.32), Inches(0.75), Inches(0.28),
             font_size=9, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
    # Year label below bar
    add_text(s1, yr,
             bx - Inches(0.05), CHART_Y_BOT + Inches(0.05), Inches(0.75), Inches(0.25),
             font_size=9, colour=DRK_GREY, align=PP_ALIGN.CENTER)

# Baseline rule
add_rect(s1, Inches(0.5), CHART_Y_BOT, Inches(5.0), Inches(0.015), DRK_GREY)

# Source note for chart
add_text(s1, "Sources: SmartCompany Smart50 2025, SmartCompany articles 2021 to 2025",
         Inches(0.22), Inches(6.72), Inches(5.0), Inches(0.28),
         font_size=7, colour=MID_GREY, italic=True)

# ── Key Commercial Signals (right column) ────────────────────────────────────
add_text(s1, "KEY COMMERCIAL SIGNALS",
         Inches(5.8), Inches(4.2), Inches(7.3), Inches(0.3),
         font_size=11, bold=True, colour=TEAL)

signals = [
    ("Smart50 2025 rank 15 confirms revenue above AUD 5M eligibility threshold",
     "SmartCompany Smart50 2025"),
    ("Permanent bricks and mortar retail: QV Melbourne CBD flagship and South Melbourne store",
     "Ragtrader 2025 / Inside Retail Australia 2025"),
    ("International wholesale entry: 12 Boarding Gate stores in Singapore and 3 stores in Malaysia",
     "businessnewsaustralia.com 2025"),
    ("Raised AUD 500,000 for 4 percent equity on Shark Tank Australia, implied valuation AUD 12.5M",
     "SmartCompany Shark Tank recap Nov 2024"),
    ("CEO Nathan Yun has publicly stated the focus is now profitability over hyper growth",
     "SmartCompany 2025"),
    ("20,000 units of underwear sold monthly, confirming category scale in a single product line",
     "SmartCompany 2025"),
]

sy = Inches(4.55)
for signal_text, signal_src in signals:
    # Teal bullet dot
    add_rect(s1, Inches(5.8), sy + Inches(0.1), Inches(0.08), Inches(0.08), TEAL)
    add_multiline(s1,
                  [{"text": signal_text, "size": 11, "colour": BLACK, "bold": False},
                   {"text": signal_src, "size": 7, "colour": MID_GREY, "bold": False, "italic": True}],
                  Inches(6.0), sy, Inches(7.1), Inches(0.38),
                  spacing_after=1)
    sy += Inches(0.42)

# Footer
add_rect(s1, Inches(0.08), Inches(7.05), W - Inches(0.08), Inches(0.015), DRK_GREY)
add_text(s1, FOOTER_LEFT,
         Inches(0.22), Inches(7.12), Inches(9), Inches(0.3),
         font_size=8, colour=DRK_GREY)
add_text(s1, DATE_LABEL,
         Inches(0.22), Inches(7.12), W - Inches(0.44), Inches(0.3),
         font_size=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE OPPORTUNITY
# ────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)

# Left amber stripe
add_rect(s2, Inches(0), Inches(0), Inches(0.08), H, AMBER_B)

# Header band
add_rect(s2, Inches(0.08), Inches(0), W - Inches(0.08), Inches(0.95), BLACK)
add_text(s2, "THE OPPORTUNITY",
         Inches(0.22), Inches(0.1), Inches(6), Inches(0.38),
         font_size=11, bold=True, colour=OFF_WHITE)
add_text(s2, "Paire: three commercial observations from ProfitPulse",
         Inches(0.22), Inches(0.5), Inches(8), Inches(0.38),
         font_size=11, colour=MID_GREY)
add_text(s2, COMPANY,
         Inches(0.22), Inches(0.1), W - Inches(0.44), Inches(0.38),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# Three observation columns
COL_W = Inches(4.41)
COL_Y = Inches(0.95)
COL_H = Inches(5.35)

OBS = [
    {
        "idx": "01",
        "bg": TEAL,
        "text_col": BLACK,
        "accent_col": BLACK,
        "header": "Four channels, one P and L: where is the margin?",
        "body": (
            "Paire now earns revenue across four distinct commercial channels: "
            "DTC online, QV Melbourne CBD store, South Melbourne store, and "
            "wholesale export across 15 retail doors in Singapore and Malaysia. "
            "Each channel carries a materially different cost and margin structure. "
            "DTC online captures the highest margin per unit with no intermediary. "
            "Physical retail adds fixed lease, staff, and fit-out costs that absorb "
            "contribution margin. Wholesale reduces per-unit margin but opens volume "
            "at scale. Without a channel by channel profitability view, growth capital "
            "flows to the highest revenue source rather than the highest return source. "
            "The four channel question needs a financial answer, not a revenue answer."
        ),
    },
    {
        "idx": "02",
        "bg": BLACK,
        "text_col": OFF_WHITE,
        "accent_col": AMBER_B,
        "header": "Six product categories and one question: which ones win?",
        "body": (
            "Paire launched in 2020 with a single product and now sells socks, "
            "underwear, bras, tops, leggings, activewear, and loungewear. Selling "
            "20,000 units of underwear each month signals genuine category scale "
            "in at least one line. In DTC apparel brands at this stage, a small "
            "share of product categories typically drives the majority of gross "
            "margin while others generate volume at low or zero contribution. "
            "Identifying which categories to scale, which to reprice, and which "
            "are diluting return on working capital is the foundation for every "
            "inventory, marketing, and expansion decision the business faces now. "
            "The Shark Tank capital deserves to flow to the proven winners."
        ),
    },
    {
        "idx": "03",
        "bg": GOLD,
        "text_col": BLACK,
        "accent_col": BLACK,
        "header": "Investor backed business needs a margin map, not just a revenue line",
        "body": (
            "Raising AUD 500,000 for 4 percent equity on Shark Tank Australia "
            "created an implied valuation of AUD 12.5 million and attached investor "
            "expectations. The sharks described it as the most successful company "
            "to appear on Shark Tank Australia. CEO Nathan Yun has publicly stated "
            "the next chapter is profitability over hyper growth. That pivot, stated "
            "publicly, requires a financial architecture that can answer: which "
            "products carry the business, which channels deliver real return, and "
            "where does growth capital generate the highest ROI. A product and "
            "channel profitability map is the prerequisite for that conversation, "
            "and for the next investor discussion."
        ),
    },
]

for i, obs in enumerate(OBS):
    cx = Inches(0.08) + Inches(i * 4.42)
    add_rect(s2, cx, COL_Y, COL_W, COL_H, obs["bg"])
    # Index number
    add_text(s2, obs["idx"],
             cx + Inches(0.2), COL_Y + Inches(0.2), COL_W - Inches(0.3), Inches(0.65),
             font_name="Georgia", font_size=40, bold=True, colour=obs["accent_col"])
    # Header
    add_text(s2, obs["header"],
             cx + Inches(0.2), COL_Y + Inches(0.95), COL_W - Inches(0.3), Inches(0.7),
             font_size=12, bold=True, colour=obs["text_col"], wrap=True)
    # Body
    add_text(s2, obs["body"],
             cx + Inches(0.2), COL_Y + Inches(1.72), COL_W - Inches(0.3), Inches(3.4),
             font_size=10, colour=obs["text_col"], wrap=True)

# Warm line below columns
add_rect(s2, Inches(0.08), Inches(6.36), W - Inches(0.08), Inches(0.015), DRK_GREY)
add_text(s2,
         "These are observations offered in good faith. Paire has built something genuinely impressive. "
         "The question is simply whether the financial architecture now matches the ambition.",
         Inches(0.22), Inches(6.44), W - Inches(0.44), Inches(0.5),
         font_size=10, colour=DRK_GREY, italic=True)

# Footer
add_rect(s2, Inches(0.08), Inches(7.05), W - Inches(0.08), Inches(0.015), DRK_GREY)
add_text(s2, FOOTER_LEFT,
         Inches(0.22), Inches(7.12), Inches(9), Inches(0.3),
         font_size=8, colour=DRK_GREY)
add_text(s2, DATE_LABEL,
         Inches(0.22), Inches(7.12), W - Inches(0.44), Inches(0.3),
         font_size=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE RECOMMENDATION
# ────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)

# Left amber stripe
add_rect(s3, Inches(0), Inches(0), Inches(0.08), H, AMBER_B)

# Header band
add_rect(s3, Inches(0.08), Inches(0), W - Inches(0.08), Inches(0.95), BLACK)
add_text(s3, "THE RECOMMENDATION",
         Inches(0.22), Inches(0.22), Inches(8), Inches(0.5),
         font_size=11, bold=True, colour=OFF_WHITE)
add_text(s3, COMPANY,
         Inches(0.22), Inches(0.22), W - Inches(0.44), Inches(0.5),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# Vertical divider between columns
DIVIDER_X = Inches(8.2)
add_rect(s3, DIVIDER_X, Inches(1.05), Inches(0.012), Inches(5.85), AMBER_D)

# ── LEFT COLUMN: Recommendation and CTAs ────────────────────────────────────
LEFT_W = Inches(7.8)

# Service name and price
add_text(s3, "Product and Service Line Profitability",
         Inches(0.22), Inches(1.1), LEFT_W, Inches(0.65),
         font_name="Georgia", font_size=22, bold=True, colour=BLACK)
add_text(s3, "$3,950 one off",
         Inches(0.22), Inches(1.8), LEFT_W, Inches(0.35),
         font_size=14, bold=True, colour=TEAL)

# One line description
add_text(s3,
         "A three week project ranking every product line and revenue channel by gross margin, "
         "contribution margin, and operational drag, with a kill, fix, or scale decision on each. "
         "For Paire: maps six product categories across four channels and shows where the Shark Tank "
         "capital generates the highest return.",
         Inches(0.22), Inches(2.2), LEFT_W, Inches(0.8),
         font_size=11, colour=BLACK, wrap=True)

# Separator rule
add_rect(s3, Inches(0.22), Inches(3.1), Inches(7.7), Inches(0.012), TEAL)

# Step one block
add_text(s3, "STEP ONE",
         Inches(0.22), Inches(3.2), Inches(3), Inches(0.3),
         font_size=10, bold=True, colour=TEAL)
add_text(s3, "Answer a few quick questions",
         Inches(0.22), Inches(3.52), LEFT_W, Inches(0.35),
         font_size=13, bold=True, colour=BLACK)
add_text(s3, "See the solutions matched to your size and industry",
         Inches(0.22), Inches(3.9), LEFT_W, Inches(0.3),
         font_size=11, colour=DRK_GREY)

# Questionnaire address (clean text, hyperlinked to clean URL, no UTM on brief)
add_text_with_link(s3,
                   "profit-pulse.com.au/full-suite-of-products",
                   "https://profit-pulse.com.au/full-suite-of-products",
                   Inches(0.22), Inches(4.23), LEFT_W, Inches(0.35),
                   font_size=12, bold=False, colour=TEAL)

# Separator
add_rect(s3, Inches(0.22), Inches(4.66), Inches(7.7), Inches(0.012), DRK_GREY)

# Direct CTA (Stripe link hidden behind text)
add_text_with_link(s3,
                   "Purchase the suggested product now to get started",
                   "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
                   Inches(0.22), Inches(4.75), LEFT_W, Inches(0.38),
                   font_size=12, bold=True, colour=AMBER_D)

# Booking link
add_rect(s3, Inches(0.22), Inches(5.25), Inches(7.7), Inches(0.012), DRK_GREY)
add_text(s3, "Prefer a conversation first?",
         Inches(0.22), Inches(5.35), LEFT_W, Inches(0.3),
         font_size=11, colour=DRK_GREY)
add_text_with_link(s3,
                   "Book a complimentary discovery call",
                   "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
                   Inches(0.22), Inches(5.67), LEFT_W, Inches(0.35),
                   font_size=12, bold=True, colour=TEAL)

# Supporting services note
add_text(s3,
         "Supporting services: C2 Working Capital Unlock  |  D2 Pricing Reset",
         Inches(0.22), Inches(6.18), LEFT_W, Inches(0.3),
         font_size=9, colour=MID_GREY)

# ── RIGHT COLUMN: Credibility panel ─────────────────────────────────────────
RIGHT_X = DIVIDER_X + Inches(0.25)
RIGHT_W = W - DIVIDER_X - Inches(0.45)

add_text(s3, "Nitesh Roopa",
         RIGHT_X, Inches(1.1), RIGHT_W, Inches(0.55),
         font_name="Georgia", font_size=20, bold=True, colour=BLACK)
add_text(s3, "CA, Managing Partner",
         RIGHT_X, Inches(1.68), RIGHT_W, Inches(0.3),
         font_size=12, colour=DRK_GREY)
add_text(s3, "ProfitPulse",
         RIGHT_X, Inches(2.0), RIGHT_W, Inches(0.4),
         font_size=16, bold=True, colour=TEAL)

add_rect(s3, RIGHT_X, Inches(2.45), RIGHT_W - Inches(0.2), Inches(0.012), TEAL)

cred_items = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Total GRBT project value: over AUD 10 billion",
    "Largest single deal: USD 1.3 billion (Cahora Bassa)",
]
cy = Inches(2.6)
for item in cred_items:
    add_rect(s3, RIGHT_X, cy + Inches(0.1), Inches(0.07), Inches(0.07), TEAL)
    add_text(s3, item,
             RIGHT_X + Inches(0.18), cy, RIGHT_W - Inches(0.2), Inches(0.35),
             font_size=10, colour=BLACK)
    cy += Inches(0.38)

add_rect(s3, RIGHT_X, cy + Inches(0.08), RIGHT_W - Inches(0.2), Inches(0.012), DRK_GREY)
cy += Inches(0.25)

contact_items = [
    ("Profit-Pulse.com.au", BLACK),
    ("Nitesh@Profit-Pulse.com.au", TEAL),
    ("+61 411 876 267", BLACK),
    ("linkedin.com/in/nitesh-roopa-77594163", DRK_GREY),
]
for ctext, ccol in contact_items:
    add_text(s3, ctext, RIGHT_X, cy, RIGHT_W, Inches(0.3),
             font_size=10, colour=ccol)
    cy += Inches(0.33)

# Footer
add_rect(s3, Inches(0.08), Inches(7.05), W - Inches(0.08), Inches(0.015), DRK_GREY)
add_text(s3, FOOTER_LEFT,
         Inches(0.22), Inches(7.12), Inches(9), Inches(0.3),
         font_size=8, colour=DRK_GREY)
add_text(s3, DATE_LABEL,
         Inches(0.22), Inches(7.12), W - Inches(0.44), Inches(0.3),
         font_size=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_10Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
