"""
Brief_TaskforceAustralia_12Jun2026.pptx
Section 6 house style: white background, amber left stripe, black header band,
black stat tiles with teal top accent, AMBER numbers, OFF_WHITE labels.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import copy, os

OUT_DIR = "/home/user/Daily-Outreach/Out-reach efforts"
OUT_FILE = os.path.join(OUT_DIR, "Brief_TaskforceAustralia_12Jun2026.pptx")

# Brand colours
BLACK    = RGBColor(0x00, 0x00, 0x00)
TEAL     = RGBColor(0x01, 0xA2, 0x96)
AMBER_L  = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D  = RGBColor(0xF6, 0xA1, 0x02)
GOLD     = RGBColor(0xE3, 0xA7, 0x12)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE= RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY = RGBColor(0x99, 0x99, 0x99)

# Slide dimensions: 13.333" x 7.5"
SW = Inches(13.333)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH

def inches(n): return Inches(n)
def pts(n):    return Pt(n)

def add_shape(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_width=None):
    from pptx.util import Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.line.fill.background()
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if line_rgb:
        shape.line.color.rgb = line_rgb
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_tb(slide, text, left, top, width, height,
           font_size=12, bold=False, color=WHITE,
           align=PP_ALIGN.LEFT, wrap=True, font_name="Calibri"):
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txb

def add_hyperlink_tb(slide, text, url, left, top, width, height,
                     font_size=12, bold=False, color=TEAL, align=PP_ALIGN.LEFT):
    from pptx.oxml.ns import qn
    from lxml import etree
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    # Add hyperlink relationship to the slide part
    rId = slide.part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True
    )
    # Attach hlinkClick to the run's rPr
    r_elem = run._r
    rPr = r_elem.get_or_add_rPr()
    hlinkClick = etree.SubElement(
        rPr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick'
    )
    hlinkClick.set(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',
        rId
    )
    return txb

# ─── SLIDE 1: Prospect Profile ──────────────────────────────────────────────

sl1 = prs.slides.add_slide(prs.slide_layouts[6])  # blank

# White background
add_shape(sl1, 0, 0, SW, SH, fill_rgb=WHITE)

# Amber left stripe (~0.1")
add_shape(sl1, 0, 0, inches(0.1), SH, fill_rgb=AMBER_D)

# Black header band (full width, ~1.0" tall)
HDR_H = inches(1.0)
add_shape(sl1, 0, 0, SW, HDR_H, fill_rgb=BLACK)

# Eyebrow left label
add_tb(sl1, "PROSPECT BRIEF", inches(0.2), inches(0.08),
       inches(4), inches(0.4), font_size=9, bold=True, color=OFF_WHITE)

# Right eyebrow: TASKFORCE AUSTRALIA in TEAL
add_tb(sl1, "TASKFORCE AUSTRALIA", inches(8.5), inches(0.08),
       inches(4.7), inches(0.4), font_size=9, bold=True, color=TEAL,
       align=PP_ALIGN.RIGHT)

# Company name (large, white, header band)
add_tb(sl1, "Taskforce Australia", inches(0.2), inches(0.35),
       inches(9), inches(0.55), font_size=26, bold=True, color=WHITE)

# Date top right
add_tb(sl1, "12 Jun 2026", inches(10.5), inches(0.35),
       inches(2.7), inches(0.55), font_size=12, bold=False, color=OFF_WHITE,
       align=PP_ALIGN.RIGHT)

# ── Framing line under header ──
add_tb(sl1, "Property maintenance technology platform | Melbourne VIC | Founded 2014",
       inches(0.2), inches(1.05), inches(10), inches(0.4),
       font_size=11, bold=False, color=BLACK)

# ── 6 Stat tiles (row of 6 across the slide) ──
# Each tile: black bg, teal top stripe 0.06", AMBER number, OFF_WHITE label, MID_GREY source
TILE_W  = inches(2.0)
TILE_H  = inches(1.55)
TILE_TOP = inches(1.55)
TILE_GAP = inches(0.12)
STRIPE_H = inches(0.06)
LEFT_OFF = inches(0.2)

stats = [
    ("$12.8M",  "Annual Revenue",         "Smart50 2025"),
    ("19",       "Team Members",           "Smart50 2025"),
    ("#37",      "Smart50 2025 Rank",      "SmartCompany"),
    ("31%",      "3Yr Avg Growth",         "Smart50 2025"),
    ("140K+",    "RentSafe Jobs",          "rentsafe.taskforce.com.au"),
    ("2014",     "Year Founded",           "Smart50 2023"),
]

for i, (num, lbl, src) in enumerate(stats):
    tx = LEFT_OFF + i * (TILE_W + TILE_GAP)
    # Black tile body
    add_shape(sl1, tx, TILE_TOP, TILE_W, TILE_H, fill_rgb=BLACK)
    # Teal top stripe
    add_shape(sl1, tx, TILE_TOP, TILE_W, STRIPE_H, fill_rgb=TEAL)
    # Number (AMBER_L, large)
    add_tb(sl1, num, tx + inches(0.08), TILE_TOP + inches(0.12),
           TILE_W - inches(0.16), inches(0.65),
           font_size=26, bold=True, color=AMBER_L, align=PP_ALIGN.CENTER)
    # Label (OFF_WHITE)
    add_tb(sl1, lbl, tx + inches(0.06), TILE_TOP + inches(0.78),
           TILE_W - inches(0.12), inches(0.42),
           font_size=9, bold=False, color=OFF_WHITE, align=PP_ALIGN.CENTER)
    # Source (MID_GREY, small)
    add_tb(sl1, src, tx + inches(0.06), TILE_TOP + inches(1.2),
           TILE_W - inches(0.12), inches(0.3),
           font_size=7, bold=False, color=MID_GREY, align=PP_ALIGN.CENTER)

# ── Commercial Signals heading ──
SIG_TOP = inches(3.25)
add_tb(sl1, "COMMERCIAL SIGNALS", inches(0.2), SIG_TOP,
       inches(5), inches(0.3), font_size=9, bold=True, color=TEAL)

signals = [
    "Three consecutive Smart50 listings (2023, 2024, 2025); rank improved 46 to 37",
    "Stated growth target: 50 to 60 percent CAGR over three years, self funded",
    "RentRepair subscription maintenance launched at AREC 2023 from $54 per month",
    "RentSafe platform: 140,000+ jobs delivered across 300 real estate offices",
    "National network of over 5,000 tradespeople across Australia",
]

for j, sig in enumerate(signals):
    y = SIG_TOP + inches(0.32) + j * inches(0.54)
    # Teal bullet dot
    add_shape(sl1, inches(0.22), y + inches(0.13), inches(0.07), inches(0.07), fill_rgb=TEAL)
    add_tb(sl1, sig, inches(0.36), y,
           inches(12.8), inches(0.48),
           font_size=10, bold=False, color=BLACK)

# ── Bottom bar ──
add_shape(sl1, 0, SH - inches(0.3), SW, inches(0.3), fill_rgb=BLACK)
add_tb(sl1, "Prepared by ProfitPulse | Profit-Pulse.com.au | Confidential",
       inches(0.2), SH - inches(0.28), inches(8), inches(0.26),
       font_size=8, bold=False, color=OFF_WHITE)
add_tb(sl1, "1 / 3", SW - inches(1.5), SH - inches(0.28),
       inches(1.3), inches(0.26), font_size=8, bold=False,
       color=OFF_WHITE, align=PP_ALIGN.RIGHT)

# ─── SLIDE 2: Three Observations ────────────────────────────────────────────

sl2 = prs.slides.add_slide(prs.slide_layouts[6])

add_shape(sl2, 0, 0, SW, SH, fill_rgb=WHITE)
add_shape(sl2, 0, 0, inches(0.1), SH, fill_rgb=AMBER_D)

HDR_H2 = inches(1.0)
add_shape(sl2, 0, 0, SW, HDR_H2, fill_rgb=BLACK)

add_tb(sl2, "OBSERVATIONS", inches(0.2), inches(0.08),
       inches(4), inches(0.4), font_size=9, bold=True, color=OFF_WHITE)
add_tb(sl2, "TASKFORCE AUSTRALIA", inches(8.5), inches(0.08),
       inches(4.7), inches(0.4), font_size=9, bold=True, color=TEAL,
       align=PP_ALIGN.RIGHT)

add_tb(sl2, "Three signals that define the growth picture",
       inches(0.2), inches(0.38), inches(9), inches(0.5),
       font_size=22, bold=True, color=WHITE)

add_tb(sl2, "12 Jun 2026", inches(10.5), inches(0.38),
       inches(2.7), inches(0.5), font_size=12, bold=False,
       color=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Three columns
COL_W   = inches(4.0)
COL_H   = inches(5.5)
COL_TOP = inches(1.1)
C1_LEFT = inches(0.2)
C2_LEFT = inches(4.5)
C3_LEFT = inches(8.8)

# Column 01 TEAL bg
add_shape(sl2, C1_LEFT, COL_TOP, COL_W, COL_H, fill_rgb=TEAL)
add_tb(sl2, "01", C1_LEFT + inches(0.15), COL_TOP + inches(0.15),
       inches(0.7), inches(0.55), font_size=30, bold=True, color=BLACK)
add_tb(sl2, "Platform leverage meets product complexity",
       C1_LEFT + inches(0.15), COL_TOP + inches(0.72),
       COL_W - inches(0.3), inches(0.9),
       font_size=14, bold=True, color=BLACK)
obs1 = (
    "$12.8M revenue with 19 people signals strong platform leverage. "
    "The launch of RentRepair adds subscription economics alongside the "
    "existing compliance job model. These two revenue streams have "
    "different cost structures, different cash cycles and different margin "
    "profiles. Without a revenue stream map and margin by product line, "
    "the growth plan rests on an untested assumption that both will scale "
    "together without cannibalising cash or attention."
)
add_tb(sl2, obs1, C1_LEFT + inches(0.15), COL_TOP + inches(1.65),
       COL_W - inches(0.3), inches(3.6),
       font_size=10, bold=False, color=BLACK)

# Column 02 BLACK bg
add_shape(sl2, C2_LEFT, COL_TOP, COL_W, COL_H, fill_rgb=BLACK)
add_tb(sl2, "02", C2_LEFT + inches(0.15), COL_TOP + inches(0.15),
       inches(0.7), inches(0.55), font_size=30, bold=True, color=AMBER_L)
add_tb(sl2, "Three rankings confirm the trajectory. The next chapter needs a costed plan.",
       C2_LEFT + inches(0.15), COL_TOP + inches(0.72),
       COL_W - inches(0.3), inches(0.9),
       font_size=14, bold=True, color=WHITE)
obs2 = (
    "Rank 46 to 37 across three Smart50 cycles and a 31 percent average "
    "growth rate are proof points, not a plan. The stated 50 to 60 percent "
    "CAGR target requires a concrete answer to three questions: what headcount "
    "is needed to hit that number, how much cash does the ramp consume before "
    "new revenue arrives, and what is the RentRepair subscriber acquisition "
    "cost and payback period? Without those numbers, the ambition is real "
    "but the roadmap is not yet costed."
)
add_tb(sl2, obs2, C2_LEFT + inches(0.15), COL_TOP + inches(1.65),
       COL_W - inches(0.3), inches(3.6),
       font_size=10, bold=False, color=OFF_WHITE)

# Column 03 GOLD bg
add_shape(sl2, C3_LEFT, COL_TOP, COL_W, COL_H, fill_rgb=GOLD)
add_tb(sl2, "03", C3_LEFT + inches(0.15), COL_TOP + inches(0.15),
       inches(0.7), inches(0.55), font_size=30, bold=True, color=BLACK)
add_tb(sl2, "Subscription economics are different from service economics",
       C3_LEFT + inches(0.15), COL_TOP + inches(0.72),
       COL_W - inches(0.3), inches(0.9),
       font_size=14, bold=True, color=BLACK)
obs3 = (
    "RentRepair from $54 per month is a low entry price designed to drive "
    "volume adoption. The compliance job model (RentSafe) earns on each "
    "transaction. Mixing the two means upfront acquisition costs for "
    "subscribers sit against a slower cash recovery curve than the job "
    "flow generates today. Modelling unit economics for RentRepair "
    "separately, with cohort payback and churn sensitivity, is essential "
    "before scaling marketing spend on the subscription product."
)
add_tb(sl2, obs3, C3_LEFT + inches(0.15), COL_TOP + inches(1.65),
       COL_W - inches(0.3), inches(3.6),
       font_size=10, bold=False, color=BLACK)

# Bottom bar
add_shape(sl2, 0, SH - inches(0.3), SW, inches(0.3), fill_rgb=BLACK)
add_tb(sl2, "Prepared by ProfitPulse | Profit-Pulse.com.au | Confidential",
       inches(0.2), SH - inches(0.28), inches(8), inches(0.26),
       font_size=8, bold=False, color=OFF_WHITE)
add_tb(sl2, "2 / 3", SW - inches(1.5), SH - inches(0.28),
       inches(1.3), inches(0.26), font_size=8, bold=False,
       color=OFF_WHITE, align=PP_ALIGN.RIGHT)

# ─── SLIDE 3: CTA ───────────────────────────────────────────────────────────

sl3 = prs.slides.add_slide(prs.slide_layouts[6])

add_shape(sl3, 0, 0, SW, SH, fill_rgb=WHITE)
add_shape(sl3, 0, 0, inches(0.1), SH, fill_rgb=AMBER_D)

HDR_H3 = inches(1.0)
add_shape(sl3, 0, 0, SW, HDR_H3, fill_rgb=BLACK)

add_tb(sl3, "NEXT STEP", inches(0.2), inches(0.08),
       inches(4), inches(0.4), font_size=9, bold=True, color=OFF_WHITE)
add_tb(sl3, "PROFITPULSE", inches(8.5), inches(0.08),
       inches(4.7), inches(0.4), font_size=9, bold=True, color=TEAL,
       align=PP_ALIGN.RIGHT)

add_tb(sl3, "Your growth is financeable. Let us show you how.",
       inches(0.2), inches(0.38), inches(9), inches(0.5),
       font_size=22, bold=True, color=WHITE)

add_tb(sl3, "12 Jun 2026", inches(10.5), inches(0.38),
       inches(2.7), inches(0.5), font_size=12, bold=False,
       color=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Vertical divider
add_shape(sl3, inches(6.7), inches(1.1), inches(0.02), inches(6.1), fill_rgb=AMBER_D)

# ── LEFT column: Service offer ──
L = inches(0.2)
LW = inches(6.3)

add_tb(sl3, "Strategic Growth Diagnostic",
       L, inches(1.15), LW, inches(0.55),
       font_size=20, bold=True, color=BLACK)

# Price tile
add_shape(sl3, L, inches(1.75), inches(2.5), inches(0.55), fill_rgb=BLACK)
add_tb(sl3, "$5,000  one off", L + inches(0.1), inches(1.78),
       inches(2.3), inches(0.48), font_size=16, bold=True, color=AMBER_L)

desc = (
    "A structured diagnostic that maps your customer concentration, "
    "identifies your highest profitability growth levers and builds a "
    "costed three year growth plan calibrated to your actual cash and "
    "headcount constraints."
)
add_tb(sl3, desc, L, inches(2.42), LW, inches(1.1),
       font_size=11, bold=False, color=BLACK)

add_tb(sl3, "Deliverable: a board ready financial model and growth roadmap "
            "with explicit assumptions, sensitivity ranges and a clear first "
            "ninety day action plan.",
       L, inches(3.6), LW, inches(0.75),
       font_size=10, bold=False, color=BLACK)

# Divider line
add_shape(sl3, L, inches(4.42), LW, inches(0.02), fill_rgb=AMBER_D)

add_tb(sl3, "Step one: answer a few quick questions",
       L, inches(4.5), LW, inches(0.35),
       font_size=11, bold=True, color=BLACK)

add_tb(sl3, "profit-pulse.com.au/full-suite-of-products",
       L, inches(4.9), LW, inches(0.35),
       font_size=11, bold=False, color=TEAL)

# Stripe CTA button appearance (rectangle + text)
add_shape(sl3, L, inches(5.35), inches(3.8), inches(0.5), fill_rgb=TEAL)
# Hyperlink via textbox overlay
add_hyperlink_tb(sl3,
    "Purchase the suggested product now to get started",
    "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
    L + inches(0.1), inches(5.37), inches(3.6), inches(0.46),
    font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_tb(sl3, "Or book a call:",
       L, inches(5.98), inches(2.0), inches(0.3),
       font_size=10, bold=False, color=BLACK)
add_hyperlink_tb(sl3,
    "Schedule a call with Nitesh",
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
    L + inches(1.5), inches(5.98), LW - inches(1.5), inches(0.3),
    font_size=10, bold=True, color=TEAL)

# ── RIGHT column: About Nitesh ──
R = inches(6.85)
RW = inches(6.3)

add_tb(sl3, "Nitesh Roopa  CA",
       R, inches(1.15), RW, inches(0.45),
       font_size=18, bold=True, color=BLACK)

add_tb(sl3, "Managing Partner, ProfitPulse",
       R, inches(1.65), RW, inches(0.35),
       font_size=12, bold=False, color=TEAL)

creds = [
    "16 years in corporate finance and commercial advisory",
    "52 completed transactions totalling AUD 10B gross deal value",
    "Largest single transaction: USD 1.3B",
    "Advisory work across Australia, South Africa, UK and UAE",
    "QIC (Queensland Investment Corporation) 2023 to 2025",
    "Nedbank Corporate and Investment Banking 2015 to 2022",
    "PwC Deals 2010 to 2014",
    "Qualified CA under SAICA (South African Institute of Chartered Accountants)",
]

for k, cred in enumerate(creds):
    y = inches(2.1) + k * inches(0.44)
    add_shape(sl3, R, y + inches(0.14), inches(0.07), inches(0.07), fill_rgb=TEAL)
    add_tb(sl3, cred, R + inches(0.15), y,
           RW - inches(0.2), inches(0.4),
           font_size=10, bold=False, color=BLACK)

# Contact block
CB_TOP = inches(5.8)
add_shape(sl3, R, CB_TOP, RW, inches(0.02), fill_rgb=BLACK)
add_tb(sl3, "Nitesh@Profit-Pulse.com.au   |   +61 411 876 267   |   Profit-Pulse.com.au",
       R, CB_TOP + inches(0.06), RW, inches(0.35),
       font_size=9, bold=False, color=BLACK)

# Bottom bar
add_shape(sl3, 0, SH - inches(0.3), SW, inches(0.3), fill_rgb=BLACK)
add_tb(sl3, "Prepared by ProfitPulse | Profit-Pulse.com.au | Confidential",
       inches(0.2), SH - inches(0.28), inches(8), inches(0.26),
       font_size=8, bold=False, color=OFF_WHITE)
add_tb(sl3, "3 / 3", SW - inches(1.5), SH - inches(0.28),
       inches(1.3), inches(0.26), font_size=8, bold=False,
       color=OFF_WHITE, align=PP_ALIGN.RIGHT)

# ─── Save ────────────────────────────────────────────────────────────────────
prs.save(OUT_FILE)
print(f"Saved: {OUT_FILE}")
