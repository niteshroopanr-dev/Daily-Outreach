"""
ProfitPulse Brief Builder - Version 3
Target: Taskforce Australia | Date: 20 Jun 2026
Three-slide prospect-facing deck. V3 house style. White background. Brand colours only. Zero dashes.
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
LT_GREY   = RGBColor(0xCC, 0xCC, 0xCC)
PANEL_BG  = RGBColor(0xF5, 0xF5, 0xF5)

W = Inches(13.333)
H = Inches(7.5)

DATE = "20 Jun 2026"
QUESTIONNAIRE_CLEAN_URL = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_G3_COMMAND       = "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D"
BOOKING_LINK            = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


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
             font_name="Arial", font_size=11, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, wrap=True, italic=False):
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


def add_hyperlink_text(slide, text, url, left, top, width, height,
                       font_name="Arial", font_size=11, bold=True, colour=TEAL):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
    rPr = run._r.get_or_add_rPr()
    hlinkClick = etree.SubElement(rPr, qn("a:hlinkClick"))
    hlinkClick.set(qn("r:id"), rId)
    return txBox


def add_amber_stripe(slide):
    add_rect(slide, Inches(0), Inches(0), Inches(0.12), H, AMBER_D)


def add_header(slide, eyebrow, right_label):
    add_rect(slide, Inches(0.12), Inches(0), W - Inches(0.12), Inches(1.0), BLACK)
    add_text(slide, eyebrow,
             Inches(0.28), Inches(0.2), Inches(8.5), Inches(0.55),
             font_name="Arial", font_size=12, bold=True, colour=OFF_WHITE)
    add_text(slide, right_label,
             Inches(10.5), Inches(0.2), Inches(2.7), Inches(0.55),
             font_name="Arial", font_size=11, bold=True, colour=TEAL,
             align=PP_ALIGN.RIGHT)


def add_footer(slide):
    add_rect(slide, Inches(0.12), Inches(7.2), W - Inches(0.12), Pt(1), LT_GREY)
    add_text(slide,
             "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
             Inches(0.28), Inches(7.25), Inches(9.5), Inches(0.22),
             font_name="Arial", font_size=8, bold=False, colour=DRK_GREY)
    add_text(slide, DATE,
             Inches(10.5), Inches(7.25), Inches(2.7), Inches(0.22),
             font_name="Arial", font_size=8, bold=False, colour=DRK_GREY,
             align=PP_ALIGN.RIGHT)


def stat_card(slide, left, top, width, height, number, label1, label2, source):
    add_rect(slide, left, top, width, height, BLACK)
    add_rect(slide, left, top, width, Pt(4), TEAL)
    add_text(slide, number,
             left + Inches(0.12), top + Inches(0.1),
             width - Inches(0.24), Inches(0.55),
             font_name="Georgia", font_size=26, bold=True, colour=WHITE)
    add_text(slide, label1,
             left + Inches(0.12), top + Inches(0.65),
             width - Inches(0.24), Inches(0.22),
             font_name="Arial", font_size=9.5, bold=False, colour=OFF_WHITE)
    if label2:
        add_text(slide, label2,
                 left + Inches(0.12), top + Inches(0.87),
                 width - Inches(0.24), Inches(0.2),
                 font_name="Arial", font_size=9, bold=False, colour=MID_GREY)
    add_text(slide, source,
             left + Inches(0.12), top + height - Inches(0.2),
             width - Inches(0.24), Inches(0.2),
             font_name="Arial", font_size=7, bold=False, colour=MID_GREY,
             italic=True)


# =============================================================================
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# =============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# =============================================================================
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)
add_amber_stripe(s1)
add_header(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

add_text(s1, "Taskforce Australia",
         Inches(0.28), Inches(1.05), Inches(12.7), Inches(0.65),
         font_name="Georgia", font_size=40, bold=True, colour=BLACK)

add_text(s1, "Property maintenance and compliance technology platform  |  Burnley, Melbourne VIC",
         Inches(0.28), Inches(1.72), Inches(12.7), Inches(0.3),
         font_name="Arial", font_size=11, bold=False, colour=TEAL)

# Stat cards
card_top = Inches(2.08)
card_h   = Inches(1.38)
cards_data = [
    ("$12.8M",  "Revenue FY2025",        "Smart50 award",          "Source: SmartCompany Smart50 2025"),
    ("31%",     "YoY growth FY2025",     "vs $9.7M in FY2024",     "Source: SmartCompany Smart50 2025"),
    ("19",      "Team members",          "Full time 2025",         "Source: SmartCompany Smart50 2025"),
    ("140K+",   "Jobs deployed",         "via RentSafe platform",  "Source: taskforce.com.au"),
    ("300+",    "Real estate offices",   "on platform 2024",       "Source: taskforce.com.au"),
    ("2014",    "Year founded",          "Burnley Melbourne VIC",  "Source: SmartCompany Smart50 2025"),
]
total_card_area = W - Inches(0.28) - Inches(0.12)
gap = Inches(0.09)
card_w = (total_card_area - 5 * gap) / 6

for i, (num, lab1, lab2, src) in enumerate(cards_data):
    cx = Inches(0.28) + i * (card_w + gap)
    stat_card(s1, cx, card_top, card_w, card_h, num, lab1, lab2, src)

# Key Commercial Signals label
add_text(s1, "KEY COMMERCIAL SIGNALS",
         Inches(0.28), Inches(3.6), Inches(8.1), Inches(0.28),
         font_name="Arial", font_size=10.5, bold=True, colour=TEAL)

signals = [
    ("Telstra Best of Business 2024, Victorian State Winner for Outstanding Growth",
     "Telstra Best of Business Awards 2024 (telstra.com.au/best-of-business-awards and eliteagent.com)"),
    ("Most Innovative Proptech, 2024 Proptech Awards Sydney",
     "2024 Proptech Association Australia Awards (industry press)"),
    ("Smart50 rank improved year on year: 2023 listed, 2024 rank 46, 2025 rank 37",
     "SmartCompany Smart50 2024 award citation and Smart50 2025 award citation"),
    ("Doubled agency clients in 12 months to over 300 real estate offices",
     "Telstra Best of Business 2024 award citation (eliteagent.com)"),
    ("RentSafe deployed across 20 consumer brands and 180 major real estate brands since 2021",
     "taskforce.com.au product pages and SmartCompany Smart50 2025"),
]

sig_y = Inches(3.96)
for sig_text, sig_src in signals:
    add_text(s1, sig_text,
             Inches(0.4), sig_y, Inches(7.8), Inches(0.22),
             font_name="Arial", font_size=10, bold=False, colour=BLACK)
    add_text(s1, sig_src,
             Inches(0.4), sig_y + Inches(0.21), Inches(7.8), Inches(0.17),
             font_name="Arial", font_size=7.5, bold=False, colour=MID_GREY, italic=True)
    sig_y += Inches(0.46)

# Revenue chart section label
add_text(s1, "REVENUE GROWTH (AUD MILLION)",
         Inches(8.6), Inches(3.6), Inches(4.6), Inches(0.28),
         font_name="Arial", font_size=10.5, bold=True, colour=TEAL)

# Chart bars
bar_base_y = Inches(6.6)
bar_area_h = Inches(2.62)
bar_w      = Inches(1.08)
bar_gap    = Inches(0.61)
bar_x0     = Inches(8.7)
max_rev    = 12.8

revenues = [
    (7.43,  "FY2023", "Smart50 2023"),
    (9.7,   "FY2024", "Smart50 2024"),
    (12.8,  "FY2025", "Smart50 2025"),
]

for i, (rev, yr, src) in enumerate(revenues):
    bx = bar_x0 + i * (bar_w + bar_gap)
    bh = bar_area_h * (rev / max_rev)
    by = bar_base_y - bh
    add_rect(s1, bx, by, bar_w, bh, TEAL)
    add_text(s1, f"${rev}M",
             bx, by - Inches(0.22), bar_w, Inches(0.22),
             font_name="Arial", font_size=9.5, bold=True, colour=BLACK,
             align=PP_ALIGN.CENTER)
    add_text(s1, yr,
             bx, bar_base_y + Inches(0.03), bar_w, Inches(0.2),
             font_name="Arial", font_size=9, bold=False, colour=DRK_GREY,
             align=PP_ALIGN.CENTER)
    add_text(s1, src,
             bx, bar_base_y + Inches(0.23), bar_w, Inches(0.18),
             font_name="Arial", font_size=7, bold=False, colour=MID_GREY,
             align=PP_ALIGN.CENTER, italic=True)

add_rect(s1, bar_x0, bar_base_y, Inches(4.2), Pt(1.5), DRK_GREY)

add_footer(s1)


# =============================================================================
# SLIDE 2: THE OPPORTUNITY
# =============================================================================
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)
add_amber_stripe(s2)
add_header(s2, "THE OPPORTUNITY", "PROFITPULSE")

add_text(s2, "Taskforce Australia   |   Three commercial observations from ProfitPulse",
         Inches(0.28), Inches(1.06), Inches(12.7), Inches(0.28),
         font_name="Arial", font_size=11, bold=False, colour=DRK_GREY)

col_top = Inches(1.42)
col_h   = Inches(5.5)
col_gap = Inches(0.12)
col_w   = (W - Inches(0.12) - Inches(0.28) - 2 * col_gap) / 3

observations = [
    (TEAL, BLACK,
     "01",
     "Three streams, one margin question",
     ("Taskforce runs three distinct service lines: RentSafe compliance checks, "
      "RentRepair maintenance, and consumer brand maintenance partnerships. "
      "At $12.8 million with 19 staff, the business is lean and the top line story is "
      "compelling. But sustained growth across three streams creates a pressing question: "
      "which stream actually delivers margin after platform costs, tradesperson "
      "coordination, and compliance overhead are fully allocated? "
      "A product profitability engagement maps that in three weeks, "
      "giving the founders a clear ranking of every revenue line by gross margin, "
      "contribution margin, and operational drag.")),
    (BLACK, WHITE,
     "02",
     "Platform scale without a financial map creates risk",
     ("Three consecutive Smart50 appearances, each at a higher rank than the last, "
      "and a Telstra Outstanding Growth award in Victoria for 2024 confirm Taskforce "
      "as one of the fastest moving property tech businesses in Australia. "
      "But rapid platform expansion typically outpaces the financial architecture. "
      "With 300 plus real estate offices, 20 consumer brands, and more than 5,000 "
      "tradespeople on the network, the margin profile of each segment is almost "
      "certainly uneven. Scaling the least profitable activity hardest is a common "
      "outcome when that picture is not clear. "
      "ProfitPulse can map this in one focused three week engagement.")),
    (GOLD, BLACK,
     "03",
     "Where does the next dollar go at $12.8 million?",
     ("Revenue grew from $7.43 million in FY2023 to $12.8 million in FY2025, "
      "a 72 percent cumulative increase in two years verified across three independent "
      "Smart50 award citations. That momentum is real. The strategic question now is "
      "capital allocation: which product line deserves the next hire, the next "
      "marketing dollar, the next technology investment? "
      "A Product and Service Line Profitability engagement produces a ranked view "
      "of every revenue line by gross margin, contribution margin, and drag, "
      "giving the founders a defensible basis for every growth decision "
      "in the next 12 months.")),
]

for i, (fill_col, text_col, idx, heading, body) in enumerate(observations):
    cx = Inches(0.28) + i * (col_w + col_gap)
    add_rect(s2, cx, col_top, col_w, col_h, fill_col)
    add_text(s2, idx,
             cx + Inches(0.2), col_top + Inches(0.15),
             col_w - Inches(0.4), Inches(0.5),
             font_name="Georgia", font_size=30, bold=True, colour=text_col)
    add_text(s2, heading,
             cx + Inches(0.2), col_top + Inches(0.7),
             col_w - Inches(0.4), Inches(0.65),
             font_name="Arial", font_size=12, bold=True, colour=text_col, wrap=True)
    add_text(s2, body,
             cx + Inches(0.2), col_top + Inches(1.42),
             col_w - Inches(0.4), col_h - Inches(1.62),
             font_name="Arial", font_size=10, bold=False, colour=text_col, wrap=True)

add_text(s2,
         ("These are observations offered in good faith. "
          "Taskforce has built something genuinely impressive. "
          "The question is simply whether the financial architecture is now keeping pace with the ambition."),
         Inches(0.28), Inches(7.0), Inches(12.7), Inches(0.38),
         font_name="Arial", font_size=9.5, bold=False, colour=DRK_GREY, italic=True)

add_footer(s2)


# =============================================================================
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# =============================================================================
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)
add_amber_stripe(s3)
add_header(s3, "THE RECOMMENDATION", "PROFITPULSE")

# ── LEFT COLUMN ───────────────────────────────────────────────────────────────
lx = Inches(0.28)
lw = Inches(7.55)

add_text(s3, "Product and Service Line Profitability",
         lx, Inches(1.08), lw, Inches(0.55),
         font_name="Georgia", font_size=20, bold=True, colour=BLACK)

add_text(s3, "$3,950 one off",
         lx, Inches(1.67), lw, Inches(0.32),
         font_name="Arial", font_size=14, bold=True, colour=TEAL)

add_text(s3,
         ("A three week project ranking every product or service line by gross margin, "
          "contribution margin, and operational drag. "
          "Gives Taskforce Australia clarity on where to scale, where to fix, and where to stop."),
         lx, Inches(2.04), lw, Inches(0.55),
         font_name="Arial", font_size=10.5, bold=False, colour=BLACK, wrap=True)

add_rect(s3, lx, Inches(2.65), lw, Pt(1.5), TEAL)

# Step 1 block
add_text(s3, "STEP ONE",
         lx, Inches(2.74), lw, Inches(0.25),
         font_name="Arial", font_size=10.5, bold=True, colour=TEAL)

add_text(s3, "Answer a few quick questions",
         lx, Inches(3.02), lw, Inches(0.28),
         font_name="Arial", font_size=12.5, bold=True, colour=BLACK)

add_text(s3, "See the solutions matched to your size and industry.",
         lx, Inches(3.34), lw, Inches(0.25),
         font_name="Arial", font_size=10.5, bold=False, colour=BLACK)

# Questionnaire address hyperlinked, shown clean with no UTM
add_hyperlink_text(s3,
                   "profit-pulse.com.au/full-suite-of-products",
                   QUESTIONNAIRE_CLEAN_URL,
                   lx, Inches(3.62), lw, Inches(0.28),
                   font_name="Arial", font_size=10.5, bold=False, colour=TEAL)

add_rect(s3, lx, Inches(4.0), lw, Pt(1.2), LT_GREY)

# Stripe CTA: link hidden behind words, address never shown as text
add_hyperlink_text(s3,
                   "Purchase the suggested product now to get started",
                   STRIPE_G3_COMMAND,
                   lx, Inches(4.1), lw, Inches(0.32),
                   font_name="Arial", font_size=12, bold=True, colour=AMBER_D)

add_text(s3, "Product and Service Line Profitability   |   $3,950 one off",
         lx, Inches(4.46), lw, Inches(0.25),
         font_name="Arial", font_size=9.5, bold=False, colour=DRK_GREY)

add_rect(s3, lx, Inches(4.8), lw, Pt(1.2), LT_GREY)

add_text(s3, "Prefer a conversation first?",
         lx, Inches(4.9), lw, Inches(0.25),
         font_name="Arial", font_size=10.5, bold=False, colour=BLACK)

add_hyperlink_text(s3,
                   "Book a complimentary discovery call",
                   BOOKING_LINK,
                   lx, Inches(5.19), lw, Inches(0.28),
                   font_name="Arial", font_size=11, bold=True, colour=TEAL)

add_text(s3, "bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/",
         lx, Inches(5.5), lw, Inches(0.22),
         font_name="Arial", font_size=8.5, bold=False, colour=DRK_GREY)

add_rect(s3, lx, Inches(5.85), lw, Pt(1.0), LT_GREY)

add_text(s3,
         "Also identified: Strategic Growth Diagnostic   |   Customer Concentration and Profitability Map",
         lx, Inches(5.93), lw, Inches(0.25),
         font_name="Arial", font_size=8.5, bold=False, colour=MID_GREY)

# ── RIGHT COLUMN: About Nitesh ────────────────────────────────────────────────
rx = Inches(8.1)
rw = Inches(5.1)

add_rect(s3, rx - Inches(0.15), Inches(1.08), rw + Inches(0.2), Inches(6.05), PANEL_BG)

add_text(s3, "Nitesh Roopa",
         rx, Inches(1.18), rw, Inches(0.5),
         font_name="Georgia", font_size=20, bold=True, colour=BLACK)

add_text(s3, "CA, Managing Partner",
         rx, Inches(1.72), rw, Inches(0.26),
         font_name="Arial", font_size=11, bold=False, colour=DRK_GREY)

add_text(s3, "ProfitPulse",
         rx, Inches(2.0), rw, Inches(0.3),
         font_name="Georgia", font_size=14, bold=True, colour=TEAL)

add_rect(s3, rx, Inches(2.36), rw - Inches(0.3), Pt(1.5), TEAL)

cred_items = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed across career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa hydro, Mozambique",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "QIC 2023 to 2025: Finance and Commercial Lead, AUD 10B Gympie Road Bypass Tunnel",
    "Nedbank CIB 2015 to 2022: Energy Finance and Principal and Equity Finance",
    "PwC South Africa 2010 to 2014: CA traineeship, Top 40 listed clients",
]
cy = Inches(2.52)
for cp in cred_items:
    add_text(s3, cp, rx, cy, rw, Inches(0.24),
             font_name="Arial", font_size=9.5, bold=False, colour=DRK_GREY, wrap=True)
    cy += Inches(0.3)

add_rect(s3, rx, cy + Inches(0.05), rw - Inches(0.3), Pt(1), LT_GREY)
cy += Inches(0.2)

contact_lines = [
    ("Profit-Pulse.com.au",                 DRK_GREY),
    ("Nitesh@Profit-Pulse.com.au",          TEAL),
    ("+61 411 876 267",                     DRK_GREY),
    ("linkedin.com/in/nitesh-roopa-77594163", DRK_GREY),
]
for c_text, c_col in contact_lines:
    add_text(s3, c_text, rx, cy, rw, Inches(0.24),
             font_name="Arial", font_size=9.5, bold=False, colour=c_col)
    cy += Inches(0.26)

add_footer(s3)


# =============================================================================
OUT_PATH = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_20Jun2026.pptx"
prs.save(OUT_PATH)
print(f"PPTX saved: {OUT_PATH}")
