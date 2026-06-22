"""
ProfitPulse Brief Builder | Taskforce Australia | 23 Jun 2026
House style: white body, amber left stripe, black header band.
Zero dashes in all output. No tier names visible.
Three slides: Commercial Intelligence Brief / The Opportunity / The Recommendation.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from lxml import etree

BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY  = RGBColor(0x88, 0x88, 0x88)
DRK_GREY  = RGBColor(0x44, 0x44, 0x44)
LT_GREY   = RGBColor(0xF8, 0xF8, 0xF8)

W = Inches(13.333)
H = Inches(7.5)


def set_bg(slide, colour):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = colour


def rct(slide, l, t, w, h, fill, border=None, bw=Pt(1)):
    sh = slide.shapes.add_shape(1, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if border:
        sh.line.color.rgb = border
        sh.line.width = bw
    else:
        sh.line.fill.background()
    return sh


def add_tb(slide, text, l, t, w, h, font="Calibri", size=12,
           bold=False, italic=False, col=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = col
    return box, run


def hyperlink(run, slide, url):
    rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
    rPr = run._r.get_or_add_rPr()
    hl = etree.SubElement(rPr, qn('a:hlinkClick'))
    hl.set(qn('r:id'), rId)


def slide_header(slide, eyebrow, right="PROFITPULSE"):
    rct(slide, Inches(0.1), 0, W - Inches(0.1), Inches(1.0), BLACK)
    add_tb(slide, eyebrow, Inches(0.25), Inches(0.12), Inches(8.5), Inches(0.4),
           size=11, bold=True, col=OFF_WHITE)
    add_tb(slide, right, Inches(8.6), Inches(0.12), Inches(4.6), Inches(0.4),
           size=11, bold=True, col=TEAL, align=PP_ALIGN.RIGHT)


def slide_footer(slide, date="23 Jun 2026"):
    add_tb(slide,
           "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
           Inches(0.25), Inches(7.15), Inches(9.5), Inches(0.28),
           size=9, col=DRK_GREY)
    add_tb(slide, date, Inches(10.3), Inches(7.15), Inches(2.9), Inches(0.28),
           size=9, col=DRK_GREY, align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BL = prs.slide_layouts[6]


# =============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# =============================================================================
s1 = prs.slides.add_slide(BL)
set_bg(s1, WHITE)

# Amber left stripe
rct(s1, 0, 0, Inches(0.1), H, AMBER_D)

slide_header(s1, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
add_tb(s1, "Taskforce Australia",
       Inches(0.25), Inches(1.06), Inches(10), Inches(0.72),
       font="Georgia", size=44, bold=True, col=BLACK)

# Descriptor
add_tb(s1, "Property maintenance and safety compliance platform  |  Burnley, Melbourne VIC",
       Inches(0.25), Inches(1.82), Inches(12.5), Inches(0.35),
       font="Calibri", size=13, col=TEAL)

# -- Stat cards (6 across) ----------------------------------------------------
STATS = [
    ("$12.8M",  "Revenue",       "FY25",           "Smart50 2025"),
    ("31%",     "Revenue",       "Growth Rate",     "Smart50 2025"),
    ("19",      "In-House",      "Team",            "Smart50 2025"),
    ("5,500",   "Trade",         "Network",         "Smart50 2025"),
    ("618%",    "Housing Jobs",  "YoY Growth",      "Smart50 2025"),
    ("#37",     "Smart50",       "2025 Rank",       "SmartCompany 2025"),
]

CARD_TOP = Inches(2.25)
CARD_H   = Inches(1.45)
CARD_W   = Inches(2.08)
CARD_GAP = Inches(0.06)
CX0      = Inches(0.25)

for i, (num, lab1, lab2, src) in enumerate(STATS):
    cx = CX0 + i * (CARD_W + CARD_GAP)
    rct(s1, cx, CARD_TOP, CARD_W, CARD_H, BLACK)
    rct(s1, cx, CARD_TOP, CARD_W, Inches(0.07), TEAL)
    add_tb(s1, num, cx + Inches(0.1), CARD_TOP + Inches(0.1),
           CARD_W - Inches(0.2), Inches(0.55),
           font="Georgia", size=28, bold=True, col=AMBER_B)
    add_tb(s1, lab1, cx + Inches(0.1), CARD_TOP + Inches(0.69),
           CARD_W - Inches(0.2), Inches(0.27),
           font="Calibri", size=11, col=OFF_WHITE)
    add_tb(s1, lab2, cx + Inches(0.1), CARD_TOP + Inches(0.96),
           CARD_W - Inches(0.2), Inches(0.27),
           font="Calibri", size=11, col=OFF_WHITE)
    add_tb(s1, src, cx + Inches(0.1), CARD_TOP + Inches(1.22),
           CARD_W - Inches(0.2), Inches(0.2),
           font="Calibri", size=7, col=MID_GREY)

# -- Revenue bar chart --------------------------------------------------------
CX_CHT  = Inches(9.4)
CW_CHT  = Inches(3.7)
CT_CHT  = Inches(3.82)
CH_CHT  = Inches(3.05)

add_tb(s1, "REVENUE GROWTH", CX_CHT, CT_CHT, CW_CHT, Inches(0.28),
       size=11, bold=True, col=TEAL)

rct(s1, CX_CHT, CT_CHT + Inches(0.33), CW_CHT, CH_CHT - Inches(0.33),
    RGBColor(0xF2, 0xF2, 0xF2))

BAR_MAX_H = Inches(1.85)
BAR_BASELINE = CT_CHT + CH_CHT - Inches(0.42)
BAR_W = Inches(1.15)
bx = CX_CHT + Inches(0.35)

for yr_lbl, rev, bar_col in [("FY24", 9.7, AMBER_D), ("FY25", 12.8, TEAL)]:
    bh = BAR_MAX_H * (rev / 14.0)
    by = BAR_BASELINE - bh
    rct(s1, bx, by, BAR_W, bh, bar_col)
    add_tb(s1, f"${rev}M", bx, by - Inches(0.28), BAR_W, Inches(0.26),
           font="Calibri", size=10, bold=True, col=BLACK, align=PP_ALIGN.CENTER)
    add_tb(s1, yr_lbl, bx, BAR_BASELINE + Inches(0.03), BAR_W, Inches(0.26),
           font="Calibri", size=10, col=DRK_GREY, align=PP_ALIGN.CENTER)
    bx += BAR_W + Inches(0.35)

add_tb(s1, "Source: Smart50 2024 and 2025, SmartCompany",
       CX_CHT, CT_CHT + CH_CHT + Inches(0.04), CW_CHT, Inches(0.22),
       font="Calibri", size=7, col=MID_GREY)

# -- Key commercial signals ---------------------------------------------------
SX = Inches(0.25)
SW = Inches(9.0)
SY = Inches(3.82)

add_tb(s1, "KEY COMMERCIAL SIGNALS", SX, SY, SW, Inches(0.28),
       size=11, bold=True, col=TEAL)

SIGS = [
    "Housing jobs grew 618% year-on-year, making housing the dominant service area  (Smart50 2025, SmartCompany)",
    "Housing division confirmed as strongest and most profitable area of the business  (Smart50 2025)",
    "Equity plan for staff and customers in active development  (Smart50 2025, SmartCompany)",
    "AI-enhanced scheduling platform expansion underway as part of four-point growth agenda  (Smart50 2025)",
    "Three consecutive Smart50 appearances: 2023, 2024, and 2025, confirming sustained performance  (SmartCompany)",
]

sy = SY + Inches(0.34)
for sig in SIGS:
    add_tb(s1, "•  " + sig, SX, sy, SW, Inches(0.37),
           font="Calibri", size=11, col=BLACK)
    sy += Inches(0.48)

slide_footer(s1)


# =============================================================================
# SLIDE 2: THE OPPORTUNITY — THREE COMMERCIAL OBSERVATIONS
# =============================================================================
s2 = prs.slides.add_slide(BL)
set_bg(s2, WHITE)
rct(s2, 0, 0, Inches(0.1), H, AMBER_D)

slide_header(s2, "THE OPPORTUNITY", "Taskforce Australia")

add_tb(s2,
       "Taskforce Australia  |  Three commercial observations from ProfitPulse",
       Inches(0.25), Inches(1.06), Inches(12.5), Inches(0.35),
       font="Calibri", size=13, col=DRK_GREY)

OBS = [
    (TEAL, WHITE,
     "01",
     "618 percent growth in one service line reshapes the financial architecture",
     ("When housing jobs grow 618 percent in a single year the revenue mix, cost "
      "structure, and customer concentration profile of the business all change "
      "simultaneously. The housing division is the acknowledged margin leader, but "
      "at this growth velocity the question shifts from which division is strongest "
      "to how strong, by how much, and on what terms. Without a service line "
      "profitability map that allocates sub-contractor costs, technology investment, "
      "and coordination time to each job type, the income statement cannot tell that "
      "story. Growing faster into the wrong margin structure has destroyed value in "
      "many platform businesses. Knowing the real housing unit economics now "
      "protects against that risk.")),
    (BLACK, OFF_WHITE,
     "02",
     "5,500 tradespeople managed by 19 people creates leverage and concentration risk simultaneously",
     ("Nineteen people generating $12.8 million creates revenue per head of "
      "approximately $674,000. That is the signature of an asset-light platform, "
      "not a traditional trades business. The platform economics are real. But they "
      "are only sustainable if the customer base is sufficiently diversified and all "
      "service lines clear their contribution hurdles after full cost allocation. As "
      "housing clients grow to dominate job volume, the concentration of that revenue "
      "in a small number of housing providers becomes the primary financial risk. A "
      "customer concentration and profitability map would quantify that exposure "
      "clearly and identify which accounts are both large and genuinely high margin.")),
    (GOLD, BLACK,
     "03",
     "The equity plan and growth agenda need a financial foundation before activation",
     ("Taskforce has stated it will create an equity plan for staff and customers. "
      "Equity plans work when anchored to a clear view of business value, margin "
      "quality, and growth trajectory. That foundation requires knowing where the "
      "earnings actually come from, which service lines and customers drive the "
      "repeatable margin, and what the valuation story looks like at exit or "
      "restructuring. Building that picture now, while the housing division is "
      "growing at 618 percent and before the equity plan is finalised, means the "
      "plan is anchored to the right numbers and the right incentive levers "
      "from the start.")),
]

COL_W   = Inches(4.1)
COL_GAP = Inches(0.1)
CX0_OBS = Inches(0.25)
COL_TOP = Inches(1.5)
COL_H   = Inches(5.2)

for i, (bg, tc, num, heading, body) in enumerate(OBS):
    cx = CX0_OBS + i * (COL_W + COL_GAP)
    rct(s2, cx, COL_TOP, COL_W, COL_H, bg)
    add_tb(s2, num,
           cx + Inches(0.18), COL_TOP + Inches(0.15),
           COL_W - Inches(0.3), Inches(0.5),
           font="Georgia", size=30, bold=True, col=tc)
    add_tb(s2, heading,
           cx + Inches(0.18), COL_TOP + Inches(0.75),
           COL_W - Inches(0.3), Inches(0.85),
           font="Calibri", size=12, bold=True, col=tc, wrap=True)
    add_tb(s2, body,
           cx + Inches(0.18), COL_TOP + Inches(1.65),
           COL_W - Inches(0.3), Inches(3.4),
           font="Calibri", size=10, col=tc, wrap=True)

add_tb(s2,
       "These observations are offered in good faith. Taskforce has built something "
       "genuinely impressive in property services, and the 618 percent housing growth "
       "signal is one of the strongest single-year moves on any Smart50 list this year. "
       "The question is simply whether the financial architecture now matches the ambition.",
       Inches(0.25), Inches(6.8), W - Inches(0.5), Inches(0.45),
       font="Calibri", size=10, italic=True, col=DRK_GREY)

slide_footer(s2)


# =============================================================================
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# =============================================================================
s3 = prs.slides.add_slide(BL)
set_bg(s3, WHITE)
rct(s3, 0, 0, Inches(0.1), H, AMBER_D)

slide_header(s3, "THE RECOMMENDATION", "Taskforce Australia")

Q_URL      = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_URL = "https://buy.stripe.com/bJe00kc5agKQepG0ZV3ks1x"
BOOK_URL   = ("https://bookings.cloud.microsoft/book/"
              "ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# -- Left column --------------------------------------------------------------
LX = Inches(0.25)
LW = Inches(7.65)
LY = Inches(1.1)

add_tb(s3, "Operational Intelligence Review",
       LX, LY, LW, Inches(0.62),
       font="Georgia", size=26, bold=True, col=BLACK)

add_tb(s3, "$6,500 one off",
       LX, LY + Inches(0.66), LW, Inches(0.38),
       font="Calibri", size=16, bold=True, col=TEAL)

add_tb(s3,
       "A six week deep dive across customer concentration and profitability, "
       "product and service line margin, workforce capacity, and operational "
       "bottlenecks. Delivers a prioritised 12 month action list calibrated "
       "to the specific growth signals visible in this business.",
       LX, LY + Inches(1.1), LW, Inches(0.8),
       font="Calibri", size=11, col=DRK_GREY, wrap=True)

# Step 1 box (teal background)
rct(s3, LX, LY + Inches(2.0), LW, Inches(1.45), TEAL)

add_tb(s3, "Step one: answer a few quick questions",
       LX + Inches(0.15), LY + Inches(2.1), LW - Inches(0.3), Inches(0.38),
       font="Calibri", size=13, bold=True, col=BLACK)

add_tb(s3, "See the solutions matched to your size and industry",
       LX + Inches(0.15), LY + Inches(2.52), LW - Inches(0.3), Inches(0.32),
       font="Calibri", size=11, col=BLACK)

# Clean questionnaire URL, hyperlinked to same clean address (no UTM on brief)
_, q_run = add_tb(s3, "profit-pulse.com.au/full-suite-of-products",
                  LX + Inches(0.15), LY + Inches(2.88), LW - Inches(0.3), Inches(0.48),
                  font="Calibri", size=12, bold=True, col=BLACK)
hyperlink(q_run, s3, Q_URL)

# Already know separator
rct(s3, LX, LY + Inches(3.55), LW, Pt(1), DRK_GREY)
add_tb(s3, "Already know this is the priority? You can begin with Operational Intelligence Review, $6,500 one off.",
       LX, LY + Inches(3.65), LW, Inches(0.32),
       font="Calibri", size=10, col=DRK_GREY)

# CTA with hidden Stripe link (only CTA words carry the link, address never shown)
_, stripe_run = add_tb(s3, "Purchase the suggested product now to get started",
                        LX, LY + Inches(4.0), LW, Inches(0.35),
                        font="Calibri", size=12, bold=True, col=AMBER_D)
hyperlink(stripe_run, s3, STRIPE_URL)

# Prefer conversation
add_tb(s3, "Prefer a conversation first?",
       LX, LY + Inches(4.5), LW, Inches(0.3),
       font="Calibri", size=11, col=DRK_GREY)

_, book_run = add_tb(s3, "Book a complimentary discovery call",
                      LX, LY + Inches(4.82), LW, Inches(0.3),
                      font="Calibri", size=11, bold=True, col=TEAL)
hyperlink(book_run, s3, BOOK_URL)

# -- Right column -------------------------------------------------------------
RX = Inches(8.1)
RW = Inches(4.95)
RY = Inches(1.1)
RH = Inches(5.95)

rct(s3, RX, RY, RW, RH, LT_GREY)
rct(s3, RX, RY, Inches(0.07), RH, TEAL)

add_tb(s3, "Nitesh Roopa",
       RX + Inches(0.22), RY + Inches(0.15), RW - Inches(0.35), Inches(0.55),
       font="Georgia", size=22, bold=True, col=BLACK)

add_tb(s3, "CA, Managing Partner",
       RX + Inches(0.22), RY + Inches(0.73), RW - Inches(0.35), Inches(0.3),
       font="Calibri", size=13, col=DRK_GREY)

add_tb(s3, "ProfitPulse",
       RX + Inches(0.22), RY + Inches(1.06), RW - Inches(0.35), Inches(0.38),
       font="Calibri", size=16, bold=True, col=TEAL)

rct(s3, RX + Inches(0.22), RY + Inches(1.52), RW - Inches(0.44), Pt(1.5), TEAL)

CREDS = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique",
    "Total GRBT project value: over AUD 10 billion",
]
cy = RY + Inches(1.68)
for c in CREDS:
    add_tb(s3, c, RX + Inches(0.22), cy, RW - Inches(0.35), Inches(0.33),
           font="Calibri", size=10, col=DRK_GREY)
    cy += Inches(0.4)

rct(s3, RX + Inches(0.22), cy + Inches(0.05), RW - Inches(0.44), Pt(1), DRK_GREY)
cy += Inches(0.22)

for contact_text, ccol in [
    ("Profit-Pulse.com.au",                        DRK_GREY),
    ("Nitesh@Profit-Pulse.com.au",                 TEAL),
    ("+61 411 876 267",                            DRK_GREY),
    ("linkedin.com/in/nitesh-roopa-77594163",      MID_GREY),
]:
    add_tb(s3, contact_text, RX + Inches(0.22), cy, RW - Inches(0.35), Inches(0.3),
           font="Calibri", size=10, col=ccol)
    cy += Inches(0.33)

slide_footer(s3)

# =============================================================================
# SAVE
# =============================================================================
OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_23Jun2026.pptx"
prs.save(OUT)
print(f"PPTX saved: {OUT}")
