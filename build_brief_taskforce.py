"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date for: 18 Jun 2026
Three-slide prospect-facing deck. White background house style.
Brand colours only. Zero dashes. No tier names.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Brand colours ──────────────────────────────────────────────────────────────
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
GREY_MID  = OFF_WHITE    # muted text on dark backgrounds (brand-compliant alias)
GREY_DRK  = BLACK       # muted text on white backgrounds (brand-compliant alias)

W = Inches(13.333)
H = Inches(7.5)


def set_background(slide, colour):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour,
             line_colour=None, line_width=None):
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
                  font_name="Arial", default_size=12,
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
            cfg = {"text":   line.get("text", ""),
                   "size":   line.get("size", default_size),
                   "colour": line.get("colour", default_colour),
                   "bold":   line.get("bold", default_bold),
                   "italic": line.get("italic", False)}
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


def add_hyperlink_text(slide, text, url, left, top, width, height,
                       font_name="Arial", font_size=11, bold=False,
                       colour=TEAL, align=PP_ALIGN.LEFT):
    """Add a text box whose single run is hyperlinked to url."""
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
    # Add hyperlink via relationship
    rId = slide.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    rPr = run._r.get_or_add_rPr()
    hlinkClick = etree.SubElement(rPr, qn("a:hlinkClick"))
    hlinkClick.set(qn("r:id"), rId)
    return txBox


def add_stat_card(slide, left, top, width, height, number, label, source):
    """Black tile with teal top stripe, large number, label, source."""
    add_rect(slide, left, top, width, Pt(4), TEAL)
    add_rect(slide, left, top + Pt(4), width, height - Pt(4), BLACK)
    add_text(slide, number,
             left + Inches(0.12), top + Inches(0.08), width - Inches(0.24), Inches(0.45),
             font_size=26, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_text(slide, label,
             left + Inches(0.12), top + Inches(0.52), width - Inches(0.24), Inches(0.45),
             font_size=10, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT, wrap=True)
    add_text(slide, source,
             left + Inches(0.12), top + Inches(0.96), width - Inches(0.24), Inches(0.28),
             font_size=7, bold=False, colour=GREY_MID, align=PP_ALIGN.LEFT, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD PRESENTATION
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]

STRIPE_W = Inches(0.09)   # left amber stripe width
HEADER_H = Inches(0.95)   # black header band height
CONTENT_L = STRIPE_W + Inches(0.18)  # left edge of content
CONTENT_W = W - STRIPE_W - Inches(0.22)
DATE_STR  = "18 Jun 2026"
FOOTER_Y  = Inches(7.12)


def add_standard_header(slide, eyebrow_text, right_text="PROFITPULSE"):
    """Amber left stripe + black header band."""
    add_rect(slide, Inches(0), Inches(0), STRIPE_W, H, AMBER_D)
    add_rect(slide, STRIPE_W, Inches(0), W - STRIPE_W, HEADER_H, BLACK)
    add_text(slide, eyebrow_text,
             CONTENT_L, Inches(0.18), Inches(8), Inches(0.55),
             font_size=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, right_text,
             Inches(9.5), Inches(0.18), Inches(3.6), Inches(0.55),
             font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)


def add_standard_footer(slide, left_text="Nitesh Roopa CA  |  Managing Partner  |  ProfitPulse  |  Profit-Pulse.com.au"):
    add_text(slide, left_text,
             CONTENT_L, FOOTER_Y, Inches(9), Inches(0.3),
             font_size=8, bold=False, colour=GREY_DRK, align=PP_ALIGN.LEFT)
    add_text(slide, DATE_STR,
             Inches(10.8), FOOTER_Y, Inches(2.3), Inches(0.3),
             font_size=8, bold=False, colour=GREY_DRK, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ─────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
add_standard_header(s1, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
add_text(s1, "Taskforce Australia",
         CONTENT_L, Inches(1.05), Inches(10), Inches(0.7),
         font_name="Georgia", font_size=40, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

# Descriptor
add_text(s1, "Property maintenance, safety compliance and PropTech platform  |  Burnley, Melbourne VIC",
         CONTENT_L, Inches(1.75), Inches(12), Inches(0.35),
         font_size=12, bold=False, colour=TEAL, align=PP_ALIGN.LEFT)

# Thin teal divider
add_rect(s1, CONTENT_L, Inches(2.15), Inches(11.5), Pt(1.5), TEAL)

# ── Stat cards row ────────────────────────────────────────────────────────────
card_top  = Inches(2.25)
card_h    = Inches(1.35)
card_w    = Inches(2.13)
card_gap  = Inches(0.08)
cards = [
    ("$12.8M",   "Revenue FY2025",       "Smart50 2025 rank 37"),
    ("31%",      "Revenue growth FY2025","Smart50 2025"),
    ("19",       "Employees",            "Smart50 2025"),
    ("5,000+",   "National tradespeople","taskforce.com.au"),
    ("140K+",    "Jobs via RentSafe",    "taskforce.com.au"),
    ("Rank 37",  "Smart50 2025",         "Up from rank 46 in 2024"),
]
cx = CONTENT_L
for num, lbl, src in cards:
    add_stat_card(s1, cx, card_top, card_w, card_h, num, lbl, src)
    cx += card_w + card_gap

# ── Revenue chart (left) and Key Commercial Signals (right) ──────────────────
chart_top = Inches(3.72)
chart_h   = Inches(2.9)
chart_l   = CONTENT_L
chart_w   = Inches(4.0)

# Chart title
add_text(s1, "REVENUE GROWTH (VERIFIED)",
         chart_l, chart_top, chart_w, Inches(0.3),
         font_size=9, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

# Chart background
add_rect(s1, chart_l, chart_top + Inches(0.32), chart_w, chart_h - Inches(0.32), WHITE)

# Bar dimensions: max height for $12.8M bar = chart_h - 0.32 - 0.5 = ~2.08"
bar_area_h = Inches(2.1)
bar_base_y = chart_top + Inches(0.32) + bar_area_h
bar_w_each = Inches(1.0)

# FY2024: $9.7M bar
bar_24_h = bar_area_h * (9.7 / 13.5)
bar_24_l = chart_l + Inches(0.5)
add_rect(s1, bar_24_l, bar_base_y - bar_24_h, bar_w_each, bar_24_h,
         TEAL, line_colour=None)
add_text(s1, "$9.7M",
         bar_24_l, bar_base_y - bar_24_h - Inches(0.28), bar_w_each, Inches(0.28),
         font_size=10, bold=True, colour=TEAL, align=PP_ALIGN.CENTER)
add_text(s1, "FY2024",
         bar_24_l, bar_base_y + Inches(0.04), bar_w_each, Inches(0.22),
         font_size=9, bold=False, colour=GREY_DRK, align=PP_ALIGN.CENTER)

# FY2025: $12.8M bar
bar_25_h = bar_area_h * (12.8 / 13.5)
bar_25_l = chart_l + Inches(2.1)
add_rect(s1, bar_25_l, bar_base_y - bar_25_h, bar_w_each, bar_25_h,
         AMBER_D, line_colour=None)
add_text(s1, "$12.8M",
         bar_25_l, bar_base_y - bar_25_h - Inches(0.28), bar_w_each + Inches(0.2), Inches(0.28),
         font_size=10, bold=True, colour=AMBER_D, align=PP_ALIGN.CENTER)
add_text(s1, "FY2025",
         bar_25_l, bar_base_y + Inches(0.04), bar_w_each, Inches(0.22),
         font_size=9, bold=False, colour=GREY_DRK, align=PP_ALIGN.CENTER)

# Chart source note
add_text(s1, "Source: SmartCompany Smart50 2024 and Smart50 2025 award citations",
         chart_l, chart_top + chart_h, chart_w, Inches(0.28),
         font_size=7, bold=False, colour=GREY_MID, italic=True, align=PP_ALIGN.LEFT)

# ── Key Commercial Signals (right panel) ─────────────────────────────────────
sig_l = chart_l + chart_w + Inches(0.35)
sig_w = W - sig_l - Inches(0.25)
sig_top = chart_top

add_text(s1, "KEY COMMERCIAL SIGNALS",
         sig_l, sig_top, sig_w, Inches(0.3),
         font_size=9, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    ("Housing division publicly described as strongest and most profitable line; new housing provider clients onboarded in last 12 months.",
     "Smart50 2025, Nov 2025"),
    ("Partnership with Real+ announced March 2025; CEO Jason Bright co-presented webinar on rental compliance challenges.",
     "taskforce.com.au, Mar 2025"),
    ("Smart50 rank improved from 46 (2024) to 37 (2025), confirming accelerating growth trajectory.",
     "SmartCompany Smart50 2024 and 2025"),
    ("Telstra Best of Business VIC State Winner, Outstanding Growth category. PropTech Award winner.",
     "eliteagent.com; proptechaustralia.com.au"),
    ("RentSafe platform: 140,000 plus jobs across 20 consumer brands, 300 real estate offices, 180 real estate companies.",
     "taskforce.com.au (public)"),
]

sy = sig_top + Inches(0.36)
for sig_text, sig_src in signals:
    add_rect(s1, sig_l, sy, Inches(0.04), Inches(0.12), TEAL)
    add_multiline(s1,
        [{"text": sig_text, "size": 9, "colour": BLACK, "bold": False},
         {"text": sig_src,  "size": 7, "colour": GREY_MID, "bold": False, "italic": True}],
        sig_l + Inches(0.1), sy, sig_w - Inches(0.1), Inches(0.55),
        default_size=9, default_colour=BLACK)
    sy += Inches(0.59)

add_standard_footer(s1)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE OPPORTUNITY
# ─────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
add_standard_header(s2, "THE OPPORTUNITY", "Taskforce Australia")

# Subtitle
add_text(s2, "Three commercial observations from ProfitPulse",
         CONTENT_L, Inches(1.05), Inches(12), Inches(0.35),
         font_size=13, bold=False, colour=GREY_DRK, align=PP_ALIGN.LEFT)

# Three columns
col_top = Inches(1.52)
col_h   = Inches(5.35)
col_gap = Inches(0.05)
available_w = W - STRIPE_W - Inches(0.22)
col_w   = (available_w - col_gap * 2) / 3
c1_l = CONTENT_L
c2_l = c1_l + col_w + col_gap
c3_l = c2_l + col_w + col_gap

col_fills  = [TEAL, BLACK, GOLD]
num_cols   = ["01", "02", "03"]
num_cols_c = [WHITE, AMBER_B, BLACK]
hdr_cols   = [WHITE, AMBER_B, BLACK]
body_cols  = [WHITE, OFF_WHITE, BLACK]

headers = [
    "Revenue velocity needs a financial map to match it",
    "The housing division is the edge. A financial map confirms it.",
    "High revenue per head signals a coming capacity inflection",
]

bodies = [
    ("Taskforce grew revenue from $9.7 million to $12.8 million in one year, "
     "a 32 percent increase on top of 26 percent the year before. That "
     "trajectory is verifiably compounding across two independent Smart50 "
     "citations. At this pace, every resourcing and client onboarding decision "
     "compounds financially. The architecture that should sit behind "
     "compounding growth includes a margin model by service line, a capacity "
     "analysis for the team, and a capital allocation plan that maps each "
     "growth scenario in dollar terms."),
    ("The housing division has been publicly identified as the strongest and "
     "most profitable segment of the business. Yet without a line-level "
     "profitability analysis, the capital allocation case for doubling down on "
     "housing sits on instinct rather than numbers. Understanding the true "
     "margin by service line changes every resourcing, pricing, and sales focus "
     "decision. This is the financial work that turns a strong instinct into a "
     "deliberate and defensible strategy with a clear financial outcome."),
    ("With 19 people and $12.8 million in revenue, the revenue per head ratio "
     "is well above average for a services business. That is a mark of an "
     "exceptional technology platform. It is also a signal that a ceiling on "
     "further growth without structural change is approaching. A strategic "
     "growth diagnostic maps exactly where that ceiling sits, what the next "
     "team or technology investment needs to look like, and what three "
     "scenarios for the next 12 months look like in dollar terms so that "
     "the decision to grow is made deliberately rather than reactively."),
]

src_notes = [
    "Sources: Smart50 2024 and Smart50 2025 (revenue and growth data)",
    "Source: Smart50 2025 (housing division public statement)",
    "Sources: Smart50 2025 (revenue $12.8M, employees 19)",
]

for i, (cl, fill_c, num_txt, nc_c, hc, bc, hdr, bdy, src) in enumerate(
        zip([c1_l, c2_l, c3_l], col_fills, num_cols, num_cols_c, hdr_cols, body_cols, headers, bodies, src_notes)):
    add_rect(s2, cl, col_top, col_w, col_h, fill_c)
    add_text(s2, num_txt,
             cl + Inches(0.2), col_top + Inches(0.18), Inches(0.8), Inches(0.6),
             font_name="Georgia", font_size=38, bold=True,
             colour=nc_c, align=PP_ALIGN.LEFT)
    add_text(s2, hdr,
             cl + Inches(0.2), col_top + Inches(0.78), col_w - Inches(0.4), Inches(0.7),
             font_size=12, bold=True, colour=hc, align=PP_ALIGN.LEFT, wrap=True)
    add_rect(s2, cl + Inches(0.2), col_top + Inches(1.5), col_w - Inches(0.4), Pt(1.5),
             nc_c)
    add_text(s2, bdy,
             cl + Inches(0.2), col_top + Inches(1.62), col_w - Inches(0.4), Inches(3.3),
             font_size=10, bold=False, colour=bc, align=PP_ALIGN.LEFT, wrap=True)
    add_text(s2, src,
             cl + Inches(0.2), col_top + col_h - Inches(0.32), col_w - Inches(0.4), Inches(0.3),
             font_size=7, bold=False, colour=nc_c, italic=True, align=PP_ALIGN.LEFT)

# Warm closing line
add_text(s2,
         "These observations are offered in good faith. Taskforce has built something genuinely impressive. "
         "The question is simply whether the financial architecture now matches the commercial ambition.",
         CONTENT_L, Inches(6.97), Inches(11.5), Inches(0.35),
         font_size=10, bold=False, colour=GREY_DRK, italic=True, align=PP_ALIGN.CENTER)

add_standard_footer(s2)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE RECOMMENDATION
# ─────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
add_standard_header(s3, "THE RECOMMENDATION", "Taskforce Australia")

LEFT_COL_L = CONTENT_L
LEFT_COL_W = Inches(7.5)
RIGHT_COL_L = CONTENT_L + LEFT_COL_W + Inches(0.25)
RIGHT_COL_W = W - RIGHT_COL_L - Inches(0.18)

# ── Left column ──────────────────────────────────────────────────────────────
# Service name heading
add_text(s3, "Strategic Growth Diagnostic",
         LEFT_COL_L, Inches(1.08), LEFT_COL_W, Inches(0.65),
         font_name="Georgia", font_size=28, bold=True,
         colour=BLACK, align=PP_ALIGN.LEFT)

# Price
add_multiline(s3,
    [{"text": "$5,000 one off", "size": 20, "colour": TEAL, "bold": True},
     {"text": "ProfitPulse verified price. The questionnaire confirms the exact figure for your size and industry.",
      "size": 10, "colour": GREY_DRK, "bold": False}],
    LEFT_COL_L, Inches(1.73), LEFT_COL_W, Inches(0.7),
    default_size=10, default_colour=GREY_DRK, spacing_after=2)

# Description
add_text(s3,
    "A six-week engagement that maps revenue, capacity, and margin headroom, "
    "then produces a 12-month growth plan with funding and capital allocation "
    "steps spelled out in three scenarios. Exactly matched to a business at "
    "Taskforce's growth rate and stage.",
    LEFT_COL_L, Inches(2.48), LEFT_COL_W, Inches(0.72),
    font_size=11, bold=False, colour=BLACK, align=PP_ALIGN.LEFT, wrap=True)

# Divider
add_rect(s3, LEFT_COL_L, Inches(3.28), LEFT_COL_W, Pt(1.5), TEAL)

# Step one block
add_rect(s3, LEFT_COL_L, Inches(3.35), LEFT_COL_W, Inches(1.22),
         WHITE, line_colour=TEAL, line_width=Pt(1))

add_text(s3, "STEP ONE",
         LEFT_COL_L + Inches(0.15), Inches(3.42), Inches(3), Inches(0.28),
         font_size=9, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, "Answer a few quick questions",
         LEFT_COL_L + Inches(0.15), Inches(3.68), LEFT_COL_W - Inches(0.3), Inches(0.3),
         font_size=13, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry.",
         LEFT_COL_L + Inches(0.15), Inches(3.98), LEFT_COL_W - Inches(0.3), Inches(0.24),
         font_size=10, bold=False, colour=GREY_DRK, align=PP_ALIGN.LEFT)

# Clean questionnaire URL (no tags, per brief link rules)
add_hyperlink_text(s3,
    "profit-pulse.com.au/full-suite-of-products",
    "https://profit-pulse.com.au/full-suite-of-products",
    LEFT_COL_L + Inches(0.15), Inches(4.22),
    LEFT_COL_W - Inches(0.3), Inches(0.28),
    font_size=10, bold=True, colour=TEAL)

# Direct Stripe CTA (text only, Stripe URL hidden behind it)
add_rect(s3, LEFT_COL_L, Inches(4.68), LEFT_COL_W, Inches(0.5),
         AMBER_D, line_colour=None)
add_hyperlink_text(s3,
    "Purchase the suggested product now to get started",
    "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
    LEFT_COL_L + Inches(0.15), Inches(4.72),
    LEFT_COL_W - Inches(0.2), Inches(0.42),
    font_size=11, bold=True, colour=WHITE)

# Prefer conversation
add_rect(s3, LEFT_COL_L, Inches(5.28), LEFT_COL_W, Pt(1), GREY_DRK)
add_text(s3, "Prefer a conversation first?",
         LEFT_COL_L, Inches(5.35), LEFT_COL_W, Inches(0.28),
         font_size=10, bold=False, colour=GREY_DRK, align=PP_ALIGN.LEFT)
add_hyperlink_text(s3,
    "Book a complimentary discovery call",
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
    LEFT_COL_L, Inches(5.63),
    LEFT_COL_W, Inches(0.28),
    font_size=10, bold=True, colour=TEAL)

# ── Right column: Nitesh credibility panel ───────────────────────────────────
add_rect(s3, RIGHT_COL_L, Inches(1.05), RIGHT_COL_W, Inches(5.9),
         BLACK, line_colour=None)

add_text(s3, "Nitesh Roopa",
         RIGHT_COL_L + Inches(0.2), Inches(1.2), RIGHT_COL_W - Inches(0.35), Inches(0.5),
         font_name="Georgia", font_size=20, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

add_text(s3, "CA, Managing Partner",
         RIGHT_COL_L + Inches(0.2), Inches(1.7), RIGHT_COL_W - Inches(0.35), Inches(0.3),
         font_size=11, bold=False, colour=WHITE, align=PP_ALIGN.LEFT)

add_text(s3, "ProfitPulse",
         RIGHT_COL_L + Inches(0.2), Inches(2.0), RIGHT_COL_W - Inches(0.35), Inches(0.38),
         font_size=16, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

add_rect(s3, RIGHT_COL_L + Inches(0.2), Inches(2.42),
         RIGHT_COL_W - Inches(0.4), Pt(1.5), TEAL)

cred_lines = [
    {"text": "16 years across 4 countries",          "size": 10, "colour": OFF_WHITE, "bold": False},
    {"text": "52 deals executed and managed",         "size": 10, "colour": OFF_WHITE, "bold": False},
    {"text": "Largest deal: USD 1.3B, Cahora Bassa", "size": 10, "colour": OFF_WHITE, "bold": False},
    {"text": "Total GRBT value in QLD: AUD 10B",     "size": 10, "colour": OFF_WHITE, "bold": False},
]
add_multiline(s3, cred_lines,
              RIGHT_COL_L + Inches(0.2), Inches(2.52),
              RIGHT_COL_W - Inches(0.35), Inches(1.2),
              default_size=10, default_colour=OFF_WHITE, spacing_after=3)

add_rect(s3, RIGHT_COL_L + Inches(0.2), Inches(3.8),
         RIGHT_COL_W - Inches(0.4), Pt(1), GREY_DRK)

contact_lines = [
    {"text": "Profit-Pulse.com.au",           "size": 10, "colour": OFF_WHITE, "bold": False},
    {"text": "Nitesh@Profit-Pulse.com.au",    "size": 10, "colour": TEAL,      "bold": False},
    {"text": "+61 411 876 267",               "size": 10, "colour": OFF_WHITE, "bold": False},
    {"text": "linkedin.com/in/nitesh-roopa-77594163",
                                              "size": 9,  "colour": GREY_MID,  "bold": False},
]
add_multiline(s3, contact_lines,
              RIGHT_COL_L + Inches(0.2), Inches(3.9),
              RIGHT_COL_W - Inches(0.35), Inches(1.5),
              default_size=10, default_colour=OFF_WHITE, spacing_after=3)

add_standard_footer(s3)


# ── Save ───────────────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_18Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
