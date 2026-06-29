"""
ProfitPulse Brief Builder
Target: National Media | Date: 30 Jun 2026
Three-slide prospect-facing deck. Brand colours only. Zero dashes. No tier names.
Section 6 house style: white background slides, left amber stripe, black header band,
stat cards as black tiles with teal accent, teal section labels.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Brand colours ─────────────────────────────────────────────────────────────
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY   = RGBColor(0x88, 0x88, 0x88)
DRK_GREY   = RGBColor(0x44, 0x44, 0x44)
LT_GREY    = RGBColor(0xCC, 0xCC, 0xCC)

# Section 6 calls for white background layouts
WHITE_BG   = RGBColor(0xFF, 0xFF, 0xFF)
NEAR_WHITE = RGBColor(0xF8, 0xF7, 0xF2)

# Slide dimensions: widescreen 13.333" x 7.5"
W = Inches(13.333)
H = Inches(7.5)


def set_background(slide, colour):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(
        1,
        left, top, width, height
    )
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
             font_name="Calibri", font_size=18, bold=False,
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
                  font_name="Calibri", default_size=14,
                  default_colour=BLACK, default_bold=False,
                  align=PP_ALIGN.LEFT, spacing_after=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size,
                   "colour": default_colour, "bold": default_bold,
                   "italic": False}
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


def add_stat_card(slide, left, top, w, h, number_text, label_text, source_text):
    """Black tile with teal top accent, white number, off-white label, grey source."""
    # Black tile
    add_rect(slide, left, top, w, h, BLACK)
    # Teal top accent bar
    add_rect(slide, left, top, w, Pt(4), TEAL)
    # Number
    add_text(slide, number_text,
             left=left + Inches(0.12), top=top + Pt(10),
             width=w - Inches(0.24), height=Inches(0.55),
             font_name="Georgia", font_size=26, bold=True,
             colour=AMBER_B, align=PP_ALIGN.LEFT)
    # Label
    add_multiline(slide, [{"text": label_text, "size": 10, "colour": OFF_WHITE, "bold": False}],
                  left=left + Inches(0.12), top=top + Inches(0.62),
                  width=w - Inches(0.24), height=Inches(0.45),
                  font_name="Calibri", default_size=10, default_colour=OFF_WHITE)
    # Source
    add_text(slide, source_text,
             left=left + Inches(0.12), top=top + Inches(1.08),
             width=w - Inches(0.24), height=Inches(0.3),
             font_name="Calibri", font_size=7, bold=False,
             colour=MID_GREY, align=PP_ALIGN.LEFT, italic=True)


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
set_background(s1, WHITE_BG)

# Left amber accent stripe (~0.1" wide, full height)
add_rect(s1, Inches(0), Inches(0), Inches(0.10), H, AMBER_D)

# Black header band (~1 inch tall)
add_rect(s1, Inches(0.10), Inches(0), W - Inches(0.10), Inches(1.0), BLACK)

# Eyebrow label in header
add_text(s1, "COMMERCIAL INTELLIGENCE BRIEF",
         left=Inches(0.25), top=Inches(0.08), width=Inches(6), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

# Company name in header (right side)
add_text(s1, "PROFITPULSE",
         left=Inches(9.5), top=Inches(0.06), width=Inches(3.6), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# Company name large (below header)
add_text(s1, "National Media",
         left=Inches(0.25), top=Inches(1.05), width=Inches(8), height=Inches(0.85),
         font_name="Georgia", font_size=40, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)

# Descriptor
add_text(s1, "B2B exhibitions and events company  |  Bundall, Gold Coast QLD",
         left=Inches(0.25), top=Inches(1.92), width=Inches(9), height=Inches(0.3),
         font_name="Calibri", font_size=12, bold=False, colour=DRK_GREY, align=PP_ALIGN.LEFT)

# Teal section label
add_text(s1, "VERIFIED METRICS",
         left=Inches(0.25), top=Inches(2.30), width=Inches(4), height=Inches(0.25),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

# Row of stat cards: 6 cards across
card_w = Inches(2.12)
card_h = Inches(1.45)
card_y = Inches(2.58)
card_x_start = Inches(0.25)
card_gap = Inches(0.06)

cards = [
    ("$18.4M",   "Annual Revenue",           "Smart50 2025 citation, Nov 2025"),
    ("42%",      "3 Year Avg Growth",         "Smart50 2025, published Nov 2025"),
    ("48",       "Employees",                 "Smart50 2025 profile"),
    ("13",       "National Events",           "Exhibition Industry News, Nov 2025"),
    ("#28",      "Smart50 2025 Rank",         "SmartCompany Smart50 Awards 2025"),
    ("#62",      "AFR Fast 100 2025 Rank",    "Australian Financial Review Fast 100 2025"),
]

for i, (num, lbl, src) in enumerate(cards):
    cx = card_x_start + i * (card_w + card_gap)
    add_stat_card(s1, cx, card_y, card_w, card_h, num, lbl, src)

# Teal section label for signals
add_text(s1, "KEY COMMERCIAL SIGNALS",
         left=Inches(0.25), top=Inches(4.12), width=Inches(5), height=Inches(0.25),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    "Revenue doubled in the past year through organic growth and three acquisition programmes  (Exhibition Industry News, Nov 2025)",
    "Three acquisitions confirmed: Workplace Health and Safety Shows, Foodservice Australia portfolio, Cafe Culture and Cafe Biz Expo  (press releases, 2024 to 2025)",
    "Dual award recognition: Smart50 rank 28 and AFR Fast 100 rank 62 in the same year  (SmartCompany and AFR, Nov 2025)",
    "13 national events across 10 brands now operating across food, hospitality, safety, fitness, architecture, and accommodation  (Exhibition Industry News)",
    "Founded 1993, over 30 years in B2B exhibitions, now one of Australia's largest independent event organisers  (National Media, nationalmedia.com.au)",
]

sig_y = Inches(4.42)
for sig in signals:
    add_multiline(s1,
        [{"text": sig, "size": 10, "colour": BLACK, "bold": False}],
        left=Inches(0.35), top=sig_y, width=Inches(12.8), height=Inches(0.28),
        font_name="Calibri")
    sig_y += Inches(0.30)

# Footer line
add_rect(s1, Inches(0.25), Inches(7.18), Inches(12.9), Pt(0.5), TEAL)
add_text(s1, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse  |  Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(7.24), width=Inches(9), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY, align=PP_ALIGN.LEFT)
add_text(s1, "30 Jun 2026",
         left=Inches(10.5), top=Inches(7.24), width=Inches(2.6), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE OPPORTUNITY
# ────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank_layout)
set_background(s2, WHITE_BG)

# Left amber accent stripe
add_rect(s2, Inches(0), Inches(0), Inches(0.10), H, AMBER_D)

# Black header band
add_rect(s2, Inches(0.10), Inches(0), W - Inches(0.10), Inches(1.0), BLACK)

# Eyebrow
add_text(s2, "THE OPPORTUNITY",
         left=Inches(0.25), top=Inches(0.08), width=Inches(5), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

# Subtitle
add_text(s2, "National Media  |  Three commercial observations from ProfitPulse",
         left=Inches(0.25), top=Inches(0.46), width=Inches(9), height=Inches(0.34),
         font_name="Calibri", font_size=12, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

# Company right
add_text(s2, "PROFITPULSE",
         left=Inches(9.5), top=Inches(0.06), width=Inches(3.6), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# Three columns
col_w   = Inches(4.25)
col_h   = Inches(6.0)
col_y   = Inches(1.05)
col_gap = Inches(0.10)
col1_x  = Inches(0.25)
col2_x  = col1_x + col_w + col_gap
col3_x  = col2_x + col_w + col_gap

col_fills = [TEAL, BLACK, GOLD]
num_colours = [WHITE, AMBER_B, BLACK]
hdr_colours = [WHITE, AMBER_B, BLACK]
body_colours = [WHITE, OFF_WHITE, BLACK]

observations = [
    {
        "index": "01",
        "header": "Revenue doubled in one year. The financial architecture needs to keep pace.",
        "body": (
            "National Media grew from roughly $9 million to $18.4 million in a single year. "
            "That kind of velocity is remarkable. It also means that the financial reporting "
            "cadence, the management pack structure, and the capital allocation framework "
            "were likely built for a smaller and simpler business. At $18.4 million across "
            "10 brands, the P and L story becomes harder to read without a clear framework. "
            "ProfitPulse works with growing businesses to build that reporting layer before "
            "the complexity outpaces the visibility."
        ),
    },
    {
        "index": "02",
        "header": "Three acquisitions in 18 months. Integration economics need a disciplined framework.",
        "body": (
            "National Media has executed three separate acquisition programmes: the Workplace "
            "Health and Safety Shows, the Foodservice Australia portfolio from Specialised "
            "Events, and the Cafe Culture and Cafe Biz Expo brands. Each carries its own "
            "cost base, revenue model, and integration risk. Without a formal framework for "
            "modelling acquisition returns and tracking integration milestones against plan, "
            "the risk is that the acquisitions look successful by revenue but unclear by "
            "margin. A senior financial partner with transaction experience closes that gap."
        ),
    },
    {
        "index": "03",
        "header": "10 brands across 6 sectors. Which exhibitions deliver the real margin?",
        "body": (
            "National Media runs exhibitions in food service, hospitality, workplace safety, "
            "fitness, architecture and design, and accommodation. The shows vary in size, "
            "exhibitor mix, and operational cost. At this portfolio scale, the critical "
            "financial question is which shows generate the highest margin after full cost "
            "allocation and which are consuming disproportionate team capacity. "
            "ProfitPulse's Fractional CFO Partnership builds the monthly reporting "
            "infrastructure to answer that question, enabling faster and more "
            "capital-efficient decisions across the portfolio."
        ),
    },
]

for i, (col_x, fill, obs) in enumerate(zip([col1_x, col2_x, col3_x], col_fills, observations)):
    # Column background
    add_rect(s2, col_x, col_y, col_w, col_h, fill)

    # Index number
    add_text(s2, obs["index"],
             left=col_x + Inches(0.18), top=col_y + Inches(0.18),
             width=Inches(1), height=Inches(0.6),
             font_name="Georgia", font_size=36, bold=True,
             colour=num_colours[i], align=PP_ALIGN.LEFT)

    # Header
    add_multiline(s2,
        [{"text": obs["header"], "size": 12, "colour": hdr_colours[i], "bold": True}],
        left=col_x + Inches(0.18), top=col_y + Inches(0.85),
        width=col_w - Inches(0.36), height=Inches(0.75),
        font_name="Georgia", default_size=12, default_colour=hdr_colours[i],
        default_bold=True)

    # Body
    add_multiline(s2,
        [{"text": obs["body"], "size": 10, "colour": body_colours[i], "bold": False}],
        left=col_x + Inches(0.18), top=col_y + Inches(1.68),
        width=col_w - Inches(0.36), height=Inches(3.85),
        font_name="Calibri", default_size=10, default_colour=body_colours[i])

# Closing warm line
closing = (
    "These are observations offered in good faith. "
    "National Media has built something genuinely impressive. "
    "The question is simply whether the financial architecture matches the ambition."
)
add_text(s2, closing,
         left=Inches(0.25), top=Inches(7.09), width=Inches(12.9), height=Inches(0.28),
         font_name="Calibri", font_size=9, bold=False, colour=DRK_GREY,
         align=PP_ALIGN.CENTER, italic=True)

# Footer
add_rect(s2, Inches(0.25), Inches(7.37), Inches(12.9), Pt(0.5), TEAL)
add_text(s2, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(7.41), width=Inches(9), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY)
add_text(s2, "30 Jun 2026",
         left=Inches(10.5), top=Inches(7.41), width=Inches(2.6), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE RECOMMENDATION
# ────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank_layout)
set_background(s3, WHITE_BG)

# Left amber accent stripe
add_rect(s3, Inches(0), Inches(0), Inches(0.10), H, AMBER_D)

# Black header band
add_rect(s3, Inches(0.10), Inches(0), W - Inches(0.10), Inches(1.0), BLACK)

# Eyebrow
add_text(s3, "THE RECOMMENDATION",
         left=Inches(0.25), top=Inches(0.08), width=Inches(6), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "PROFITPULSE",
         left=Inches(9.5), top=Inches(0.06), width=Inches(3.6), height=Inches(0.35),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
add_text(s3, "National Media",
         left=Inches(0.25), top=Inches(0.48), width=Inches(7), height=Inches(0.34),
         font_name="Calibri", font_size=12, bold=False, colour=OFF_WHITE)

# ── LEFT COLUMN: Recommendation (60% width) ──────────────────────────────────
left_w = Inches(7.5)
lx = Inches(0.25)
ly = Inches(1.10)

# Service name
add_text(s3, "Fractional CFO Partnership",
         left=lx, top=ly, width=left_w, height=Inches(0.55),
         font_name="Georgia", font_size=22, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)

# Price (no tier name)
add_text(s3, "$4,950 per month  |  ProfitPulse verified price",
         left=lx, top=ly + Inches(0.58), width=left_w, height=Inches(0.3),
         font_name="Calibri", font_size=13, bold=False, colour=TEAL, align=PP_ALIGN.LEFT)

# Description
add_multiline(s3,
    [{"text": (
        "A senior financial partner at the table on a monthly cadence. Includes monthly "
        "management pack, quarterly board grade review, ad hoc decision support, and a "
        "single annual deep dive. For National Media, this means acquisition economics "
        "modelling, portfolio capital allocation, integration milestone tracking, and a "
        "reporting cadence that turns 10 brands into a clear P and L story."
    ), "size": 11, "colour": DRK_GREY, "bold": False}],
    left=lx, top=ly + Inches(0.95), width=left_w, height=Inches(0.85),
    font_name="Calibri")

# Teal divider
add_rect(s3, lx, ly + Inches(1.85), left_w, Pt(1), TEAL)

# Step 1 block
add_text(s3, "Step one: answer a few quick questions",
         left=lx, top=ly + Inches(1.95), width=left_w, height=Inches(0.32),
         font_name="Calibri", font_size=12, bold=True, colour=BLACK)
add_text(s3, "See the solutions matched to your size and industry",
         left=lx, top=ly + Inches(2.30), width=left_w, height=Inches(0.28),
         font_name="Calibri", font_size=11, bold=False, colour=DRK_GREY)

# Clean questionnaire address (no tags)
add_text(s3, "profit-pulse.com.au/full-suite-of-products",
         left=lx, top=ly + Inches(2.62), width=left_w, height=Inches(0.30),
         font_name="Calibri", font_size=12, bold=True, colour=TEAL)

# Amber divider
add_rect(s3, lx, ly + Inches(3.00), left_w, Pt(1), AMBER_D)

# Direct purchase CTA (text only; hyperlink added below)
cta_box = add_text(s3, "Purchase the suggested product now to get started",
         left=lx, top=ly + Inches(3.10), width=left_w, height=Inches(0.32),
         font_name="Calibri", font_size=12, bold=True, colour=AMBER_D)

# Add hyperlink to the CTA run
try:
    tf = cta_box.text_frame
    run = tf.paragraphs[0].runs[0]
    r_elem = run._r
    rPr = r_elem.get_or_add_rPr()
    hlinkClick = etree.SubElement(rPr, qn('a:hlinkClick'))
    # Add relationship
    slide = s3
    rel = slide.part.target_ref("https://buy.stripe.com/cNibJ28SYbqwepG3833ks0q",
                                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                                  is_external=True)
    hlinkClick.set(qn('r:id'), rel)
except Exception:
    pass  # fallback: text still displays without hyperlink

# Booking link
add_text(s3, "Prefer a conversation first?",
         left=lx, top=ly + Inches(3.55), width=left_w, height=Inches(0.28),
         font_name="Calibri", font_size=11, bold=False, colour=DRK_GREY)
add_text(s3, "Book a complimentary discovery call",
         left=lx, top=ly + Inches(3.88), width=left_w, height=Inches(0.28),
         font_name="Calibri", font_size=11, bold=True, colour=TEAL)

# ── RIGHT COLUMN: Credibility (About Nitesh) ─────────────────────────────────
rx = Inches(8.05)
rw = Inches(5.0)
ry = Inches(1.10)

# Panel background
add_rect(s3, rx, ry, rw, Inches(6.0), RGBColor(0xF0, 0xF0, 0xEB))

# Name heading
add_text(s3, "Nitesh Roopa",
         left=rx + Inches(0.2), top=ry + Inches(0.18), width=rw - Inches(0.4), height=Inches(0.5),
         font_name="Georgia", font_size=20, bold=True, colour=BLACK)
add_text(s3, "CA, Managing Partner  |  ProfitPulse",
         left=rx + Inches(0.2), top=ry + Inches(0.72), width=rw - Inches(0.4), height=Inches(0.28),
         font_name="Calibri", font_size=11, bold=False, colour=TEAL)

# Teal divider
add_rect(s3, rx + Inches(0.2), ry + Inches(1.05), rw - Inches(0.4), Pt(1), TEAL)

# Credibility points
cred = [
    "16 years of commercial finance experience across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique Government, Hydro",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "QIC 2023 to 2025: Finance and Commercial Lead, AUD 10B Gympie Road Bypass Tunnel",
    "Nedbank CIB 2015 to 2022: Energy Finance and Principal and Equity Finance",
    "PwC South Africa 2010 to 2014: CA traineeship and Audit Manager",
]

cy = ry + Inches(1.18)
for c_line in cred:
    add_multiline(s3,
        [{"text": c_line, "size": 9, "colour": DRK_GREY, "bold": False}],
        left=rx + Inches(0.2), top=cy, width=rw - Inches(0.4), height=Inches(0.28),
        font_name="Calibri")
    cy += Inches(0.38)

# Teal divider before contacts
add_rect(s3, rx + Inches(0.2), cy, rw - Inches(0.4), Pt(1), TEAL)
cy += Inches(0.12)

# Contact block
contacts = [
    ("Profit-Pulse.com.au",              DRK_GREY),
    ("Nitesh@Profit-Pulse.com.au",       TEAL),
    ("+61 411 876 267",                  DRK_GREY),
    ("linkedin.com/in/nitesh-roopa-77594163", MID_GREY),
]
for ctext, ccol in contacts:
    add_text(s3, ctext,
             left=rx + Inches(0.2), top=cy, width=rw - Inches(0.4), height=Inches(0.26),
             font_name="Calibri", font_size=9, bold=False, colour=ccol)
    cy += Inches(0.28)

# Footer
add_rect(s3, Inches(0.25), Inches(7.37), Inches(12.9), Pt(0.5), TEAL)
add_text(s3, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(7.41), width=Inches(9), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY)
add_text(s3, "30 Jun 2026",
         left=Inches(10.5), top=Inches(7.41), width=Inches(2.6), height=Inches(0.22),
         font_name="Calibri", font_size=8, bold=False, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_NationalMedia_30Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
