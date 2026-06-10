"""
ProfitPulse Nightly Outreach Engine – Brief Builder
Target: National Media Pty Ltd
Date stamp: 11 Jun 2026
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Brand colours ────────────────────────────────────────────────────────────
BLACK   = RGBColor(0x00, 0x00, 0x00)
TEAL    = RGBColor(0x01, 0xA2, 0x96)
AMBER_B = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D = RGBColor(0xF6, 0xA1, 0x02)
GOLD    = RGBColor(0xE3, 0xA7, 0x12)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHT  = RGBColor(0xE6, 0xE5, 0xDE)

# ── Canvas ───────────────────────────────────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

STRIPE_W   = Inches(0.12)   # left amber stripe
HEADER_H   = Inches(1.0)    # black header band
FOOTER_H   = Inches(0.28)
FOOTER_TOP = H - FOOTER_H

def new_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def blank_slide(prs):
    layout = prs.slide_layouts[6]   # completely blank
    return prs.slides.add_slide(layout)

# ── Low-level shape helpers ───────────────────────────────────────────────────
def add_rect(slide, left, top, width, height, fill_rgb, line_rgb=None, line_pt=0):
    shape = slide.shapes.add_shape(
        1,   # MSO_SHAPE_TYPE.RECTANGLE = 1
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, left, top, width, height, text, font_size, font_rgb,
                bold=False, italic=False, align=PP_ALIGN.LEFT,
                font_name="Calibri", word_wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = word_wrap
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.color.rgb = font_rgb
    run.font.bold  = bold
    run.font.italic = italic
    run.font.name  = font_name
    return txb

def set_para(para, text, font_size, font_rgb, bold=False, italic=False,
             align=PP_ALIGN.LEFT, font_name="Calibri"):
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = font_rgb
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    return run

def add_tf_textbox(slide, left, top, width, height):
    """Returns shape with empty text_frame for multi-paragraph use."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = True
    return txb

# ── Shared chrome ─────────────────────────────────────────────────────────────
def add_chrome(slide, eyebrow, date_str="11 Jun 2026"):
    # White background
    add_rect(slide, 0, 0, W, H, WHITE)
    # Left amber stripe
    add_rect(slide, 0, 0, STRIPE_W, H, AMBER_D)
    # Black header band
    add_rect(slide, 0, 0, W, HEADER_H, BLACK)
    # Eyebrow in header (left side, after stripe)
    add_textbox(slide,
        STRIPE_W + Inches(0.18), Inches(0.32),
        Inches(7), Inches(0.35),
        eyebrow, 11, OFFWHT, bold=False, font_name="Calibri")
    # PROFITPULSE brand name (right side of header)
    add_textbox(slide,
        Inches(9.8), Inches(0.28),
        Inches(3.3), Inches(0.45),
        "PROFITPULSE", 14, AMBER_D, bold=True,
        align=PP_ALIGN.RIGHT, font_name="Calibri")
    # Footer line
    add_rect(slide, STRIPE_W, FOOTER_TOP, W - STRIPE_W, Pt(1), TEAL)
    footer_txt = (
        "Prepared by Nitesh Roopa  CA, Managing Partner, ProfitPulse  |  "
        "Profit-Pulse.com.au  |  " + date_str
    )
    add_textbox(slide,
        STRIPE_W + Inches(0.12), FOOTER_TOP + Inches(0.04),
        Inches(12.5), Inches(0.22),
        footer_txt, 7.5, RGBColor(0x88, 0x88, 0x88), font_name="Calibri")

# ── Stat card helper ──────────────────────────────────────────────────────────
def add_stat_card(slide, left, top, width, height,
                  number, label_line1, label_line2, source):
    """Black tile, teal top stripe, white number, white/offwhite label."""
    TEAL_STRIPE = Inches(0.045)
    # Black body
    add_rect(slide, left, top, width, height, BLACK)
    # Teal top stripe
    add_rect(slide, left, top, width, TEAL_STRIPE, TEAL)
    # Number
    add_textbox(slide,
        left + Inches(0.1), top + TEAL_STRIPE + Inches(0.04),
        width - Inches(0.12), Inches(0.48),
        number, 26, AMBER_B, bold=True,
        align=PP_ALIGN.LEFT, font_name="Georgia")
    # Label
    lbl_top = top + TEAL_STRIPE + Inches(0.52)
    add_textbox(slide,
        left + Inches(0.1), lbl_top,
        width - Inches(0.12), Inches(0.46),
        label_line1 + (" " + label_line2 if label_line2 else ""),
        9, OFFWHT, bold=False, font_name="Calibri")
    # Source
    src_top = top + height - Inches(0.22)
    add_textbox(slide,
        left + Inches(0.1), src_top,
        width - Inches(0.12), Inches(0.2),
        source, 7, RGBColor(0xAA, 0xAA, 0xAA), font_name="Calibri")

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1  –  Commercial Intelligence Brief
# ═══════════════════════════════════════════════════════════════════════════════
def build_slide1(prs):
    slide = blank_slide(prs)
    add_chrome(slide, "COMMERCIAL INTELLIGENCE BRIEF")

    CONTENT_LEFT = STRIPE_W + Inches(0.22)
    CONTENT_W    = W - CONTENT_LEFT - Inches(0.22)

    # Company name (large, below header)
    add_textbox(slide,
        CONTENT_LEFT, HEADER_H + Inches(0.1),
        CONTENT_W, Inches(0.56),
        "National Media", 36, BLACK, bold=True, font_name="Georgia")

    # One-line descriptor
    add_textbox(slide,
        CONTENT_LEFT, HEADER_H + Inches(0.62),
        CONTENT_W, Inches(0.28),
        "B2B exhibition and events producer  |  Bundall, Gold Coast QLD",
        11, TEAL, bold=False, font_name="Calibri")

    # ── Stat cards row ──────────────────────────────────────────────────────
    CARD_TOP  = HEADER_H + Inches(1.02)
    CARD_H    = Inches(1.55)
    CARD_GAP  = Inches(0.14)
    CARD_W    = (CONTENT_W - 5 * CARD_GAP) / 6

    cards = [
        ("$18.4M",  "Annual revenue",      "",                 "Smart50 Awards 2025"),
        ("42%",     "Revenue growth",       "3yr avg",          "Smart50 Awards 2025"),
        ("48",      "Team members",         "",                 "Smart50 Awards 2025"),
        ("#28",     "Smart50 rank",         "2025",             "SmartCompany Nov 2025"),
        ("30+",     "Event brands",         "developed",        "nationalmedia.com.au"),
        ("850+",    "Exhibitors across",    "event portfolio",  "Brad Langton, 2024"),
    ]
    for i, (num, l1, l2, src) in enumerate(cards):
        left = CONTENT_LEFT + i * (CARD_W + CARD_GAP)
        add_stat_card(slide, left, CARD_TOP, CARD_W, CARD_H, num, l1, l2, src)

    # ── Section label: Key Commercial Signals ───────────────────────────────
    SIG_TOP = CARD_TOP + CARD_H + Inches(0.18)
    add_textbox(slide,
        CONTENT_LEFT, SIG_TOP,
        Inches(3), Inches(0.22),
        "KEY COMMERCIAL SIGNALS", 9, TEAL, bold=True, font_name="Calibri")

    signals = [
        ("Smart50 #28 and AFR Fast 100 both awarded in 2025, confirming multi-year revenue compounding across the event "
         "portfolio.  (SmartCompany Nov 2025, exhibitionindustrynews.com.au)"),
        ("Acquired five Foodservice and Hospitality events from Specialised Events in May 2024, adding Foodservice "
         "Australia, Food and Hospitality QLD, Aged Care Catering Summit, National Restaurant Conference, and "
         "Australian Chef of the Year.  (foodandhospitality.com.au, May 2024)"),
        ("Design and Build Week at ICC Sydney, 11 to 13 June 2026, combining Futurebuild Australia, Design Show "
         "Australia, and Kitchen and Bath Show into a single national destination event.  (tradefairdates.com)"),
        ("National events footprint now spans Brisbane, Sydney, Melbourne, and Gold Coast with 47,000 sqm of "
         "exhibition space and over 35,000 attendees expected across the portfolio.  (Brad Langton, 2024)"),
        ("AFR Fast 100 article notes revenue has doubled in a single year, underpinned by growth in Foodservice "
         "and Workplace Safety portfolios.  (exhibitionindustrynews.com.au, 2025)"),
    ]

    SIG_LIST_TOP = SIG_TOP + Inches(0.25)
    SIG_LINE_H   = Inches(0.34)
    for j, sig in enumerate(signals):
        add_textbox(slide,
            CONTENT_LEFT + Inches(0.08), SIG_LIST_TOP + j * SIG_LINE_H,
            CONTENT_W - Inches(0.1), SIG_LINE_H,
            "•  " + sig,
            8.5, BLACK, font_name="Calibri")

    return slide

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2  –  The Opportunity
# ═══════════════════════════════════════════════════════════════════════════════
def build_slide2(prs):
    slide = blank_slide(prs)
    add_chrome(slide, "THE OPPORTUNITY")

    # Subtitle
    add_textbox(slide,
        STRIPE_W + Inches(0.22), HEADER_H + Inches(0.08),
        Inches(11), Inches(0.3),
        "National Media  |  Three commercial observations from ProfitPulse",
        11, BLACK, bold=False, font_name="Calibri")

    # ── Three observation columns ────────────────────────────────────────────
    COL_TOP  = HEADER_H + Inches(0.48)
    COL_H    = FOOTER_TOP - COL_TOP - Inches(0.44)
    COL_GAP  = Inches(0.14)
    AVAIL_W  = W - STRIPE_W - Inches(0.22) * 2
    COL_W    = (AVAIL_W - 2 * COL_GAP) / 3
    COL_LEFT = [
        STRIPE_W + Inches(0.22),
        STRIPE_W + Inches(0.22) + COL_W + COL_GAP,
        STRIPE_W + Inches(0.22) + 2 * (COL_W + COL_GAP),
    ]
    fills = [TEAL, BLACK, GOLD]
    text_colours = [BLACK, WHITE, BLACK]

    observations = [
        (
            "01",
            "Acquisition integration: know which events earn",
            ("National Media now manages more than 30 event brands, including five acquired Foodservice and "
             "Hospitality titles added in May 2024. Each event carries its own venue commitment, staffing "
             "structure, exhibitor base, and overhead. Without a line by line view of which events contribute "
             "margin and which consume it, growth capital tends to follow volume rather than return. A "
             "Strategic Growth Diagnostic maps revenue, capacity, and margin headroom across the combined "
             "portfolio and produces a 12 month growth plan with capital allocation steps, so the next "
             "phase of expansion is deliberate rather than reactive.")
        ),
        (
            "02",
            "Multi-city events create a cashflow timing challenge",
            ("The events business model concentrates cash inflows at contract signing and exhibitor deposit "
             "stages, often six to twelve months before event dates. Expenses then surge in the final "
             "weeks before each show. With events now running across Brisbane, Sydney, Melbourne, and "
             "Gold Coast at different times of the year, these timing mismatches stack. A 13 Week Cash "
             "Flow Build gives the finance team a rolling forward view across all upcoming event dates, "
             "replacing the common end of quarter surprise with a planned, scenario tested cash calendar "
             "that the team can run independently.")
        ),
        (
            "03",
            "Growth momentum needs a funded road map",
            ("From 48 to 50 or more people, from $18.4 million toward a doubling trajectory confirmed by "
             "both Smart50 and AFR Fast 100, National Media has built genuine commercial momentum. At "
             "this scale, the constraint shifts from growth appetite to financial architecture: where to "
             "concentrate the next dollar, how much capacity exists before the next hire is required, "
             "and which events warrant increased marketing investment. ProfitPulse works alongside "
             "owner led businesses at exactly this inflection point, building the plan and the numbers "
             "that turn rapid growth into durable margin.")
        ),
    ]

    for i, (idx, header, body) in enumerate(observations):
        fc = fills[i]
        tc = text_colours[i]
        LEFT = COL_LEFT[i]
        # Column fill
        add_rect(slide, LEFT, COL_TOP, COL_W, COL_H, fc)
        # Index number
        add_textbox(slide,
            LEFT + Inches(0.18), COL_TOP + Inches(0.16),
            COL_W - Inches(0.2), Inches(0.5),
            idx, 32, tc, bold=True, font_name="Georgia")
        # Header line
        add_textbox(slide,
            LEFT + Inches(0.18), COL_TOP + Inches(0.66),
            COL_W - Inches(0.25), Inches(0.5),
            header, 11, tc, bold=True, font_name="Calibri")
        # Body
        add_textbox(slide,
            LEFT + Inches(0.18), COL_TOP + Inches(1.2),
            COL_W - Inches(0.25), COL_H - Inches(1.3),
            body, 9.5, tc, bold=False, font_name="Calibri")

    # Warm closing line
    add_textbox(slide,
        STRIPE_W + Inches(0.22), FOOTER_TOP - Inches(0.38),
        W - STRIPE_W - Inches(0.44), Inches(0.3),
        ("These observations are offered in good faith. National Media has built something genuinely impressive. "
         "The question is simply whether the financial architecture matches the ambition."),
        9, RGBColor(0x44, 0x44, 0x44), italic=True, font_name="Calibri")

    return slide

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3  –  Recommendation and How to Start
# ═══════════════════════════════════════════════════════════════════════════════
def build_slide3(prs):
    slide = blank_slide(prs)
    add_chrome(slide, "THE RECOMMENDATION")

    BODY_TOP = HEADER_H + Inches(0.12)
    BODY_H   = FOOTER_TOP - BODY_TOP - Inches(0.06)
    GAP      = Inches(0.2)
    COL_L_W  = Inches(7.0)
    COL_L_LEFT = STRIPE_W + Inches(0.22)
    COL_R_LEFT = COL_L_LEFT + COL_L_W + GAP
    COL_R_W    = W - COL_R_LEFT - Inches(0.22)

    # ── LEFT COLUMN: recommendation ──────────────────────────────────────────
    # Service name (no dash)
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(0.08),
        COL_L_W, Inches(0.52),
        "Strategic Growth Diagnostic",
        22, BLACK, bold=True, font_name="Georgia")

    # Price (plain figure, no tier name)
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(0.6),
        Inches(3.5), Inches(0.3),
        "$5,000  one off",
        13, TEAL, bold=True, font_name="Calibri")

    # Description
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(0.96),
        COL_L_W, Inches(0.82),
        ("A six week engagement that maps revenue, capacity, and margin headroom across the combined event "
         "portfolio, then produces a 12 month growth plan with funding and capital allocation steps. "
         "Three scenario modelling. For National Media, this is the funded road map that turns a "
         "high growth event portfolio into a deliberately scaled business."),
        10, BLACK, font_name="Calibri")

    # Divider line
    add_rect(slide, COL_L_LEFT, BODY_TOP + Inches(1.86), COL_L_W, Pt(1), TEAL)

    # Step 1 block
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(1.96),
        Inches(2), Inches(0.26),
        "STEP ONE", 9, TEAL, bold=True, font_name="Calibri")

    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(2.2),
        COL_L_W, Inches(0.26),
        "Answer a few quick questions and see the solutions matched to your size and industry.",
        10, BLACK, font_name="Calibri")

    # Questionnaire address (clean, no UTM tags on brief per rule)
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(2.5),
        COL_L_W, Inches(0.28),
        "profit-pulse.com.au/full-suite-of-products",
        10, TEAL, bold=False, font_name="Calibri")

    # CTA button block (amber background, text = CTA, hyperlink = Stripe)
    CTA_TOP = BODY_TOP + Inches(2.9)
    CTA_H   = Inches(0.44)
    add_rect(slide, COL_L_LEFT, CTA_TOP, Inches(4.2), CTA_H, AMBER_D)
    # CTA text (the hyperlink to Stripe is applied below in code)
    cta_tb = slide.shapes.add_textbox(
        COL_L_LEFT + Inches(0.18), CTA_TOP + Inches(0.08),
        Inches(3.9), Inches(0.3))
    cta_tf = cta_tb.text_frame
    cta_para = cta_tf.paragraphs[0]
    cta_run = cta_para.add_run()
    cta_run.text = "Purchase the suggested product now to get started"
    cta_run.font.size = Pt(10)
    cta_run.font.bold = True
    cta_run.font.color.rgb = BLACK
    cta_run.font.name = "Calibri"
    # Add hyperlink to Stripe (hidden behind text, address never shown as text)
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from lxml import etree
    stripe_url = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"
    rId = slide.part.relate_to(stripe_url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    rPr = cta_run._r.get_or_add_rPr()
    hlinkClick = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick")
    hlinkClick.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rId)

    # Prefer a conversation
    add_textbox(slide,
        COL_L_LEFT, BODY_TOP + Inches(3.46),
        COL_L_W, Inches(0.22),
        "Prefer a conversation first?  Book a complimentary discovery call.",
        9.5, BLACK, italic=True, font_name="Calibri")

    # Booking link text (hyperlinked)
    bl_tb = slide.shapes.add_textbox(
        COL_L_LEFT, BODY_TOP + Inches(3.68),
        Inches(5), Inches(0.22))
    bl_tf = bl_tb.text_frame
    bl_para = bl_tf.paragraphs[0]
    bl_run = bl_para.add_run()
    bl_run.text = "bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au"
    bl_run.font.size = Pt(8.5)
    bl_run.font.color.rgb = TEAL
    bl_run.font.name = "Calibri"
    booking_url = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
    rId2 = slide.part.relate_to(booking_url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    rPr2 = bl_run._r.get_or_add_rPr()
    hlinkClick2 = etree.SubElement(rPr2, "{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick")
    hlinkClick2.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rId2)

    # ── RIGHT COLUMN: About Nitesh ───────────────────────────────────────────
    # Teal header tile
    add_rect(slide, COL_R_LEFT, BODY_TOP, COL_R_W, Inches(0.8), TEAL)
    add_textbox(slide,
        COL_R_LEFT + Inches(0.14), BODY_TOP + Inches(0.08),
        COL_R_W - Inches(0.18), Inches(0.34),
        "Nitesh Roopa", 16, BLACK, bold=True, font_name="Georgia")
    add_textbox(slide,
        COL_R_LEFT + Inches(0.14), BODY_TOP + Inches(0.44),
        COL_R_W - Inches(0.18), Inches(0.24),
        "CA, Managing Partner  |  ProfitPulse", 9, BLACK, font_name="Calibri")

    cred_items = [
        "16 years of experience across 4 countries",
        "52 deals executed and managed",
        "Largest single deal: USD 1.3 billion  Cahora Bassa Hydro, Mozambique",
        "Total GRBT project value in QLD: over AUD 10 billion",
    ]
    for k, cred in enumerate(cred_items):
        add_textbox(slide,
            COL_R_LEFT + Inches(0.14), BODY_TOP + Inches(0.96) + k * Inches(0.42),
            COL_R_W - Inches(0.18), Inches(0.38),
            "•  " + cred, 9.5, BLACK, font_name="Calibri")

    # Divider
    add_rect(slide, COL_R_LEFT, BODY_TOP + Inches(2.72), COL_R_W, Pt(1), AMBER_D)

    # Contact block
    contact_lines = [
        ("Profit-Pulse.com.au",                      TEAL),
        ("Nitesh@Profit-Pulse.com.au",               BLACK),
        ("+61 411 876 267",                           BLACK),
        ("linkedin.com/in/nitesh-roopa-77594163",     TEAL),
    ]
    for m, (line, col) in enumerate(contact_lines):
        add_textbox(slide,
            COL_R_LEFT + Inches(0.14),
            BODY_TOP + Inches(2.86) + m * Inches(0.36),
            COL_R_W - Inches(0.18), Inches(0.3),
            line, 9.5, col, font_name="Calibri")

    return slide

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
OUT_DIR   = "/home/user/Daily-Outreach/Out-reach efforts"
PPTX_PATH = f"{OUT_DIR}/National_Media_Brief_11Jun2026.pptx"

prs = new_prs()
build_slide1(prs)
build_slide2(prs)
build_slide3(prs)
prs.save(PPTX_PATH)
print(f"Saved: {PPTX_PATH}")
