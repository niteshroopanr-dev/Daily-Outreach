"""
ProfitPulse Brief PDF Builder
Target: Taskforce Australia | Date: 27 Jun 2026
White background house style. Brand colours only. Zero dashes.
Page size 960pt x 540pt (widescreen 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

pt = 1
W = 960
H = 540

# ── Brand colours ─────────────────────────────────────────────────────────────
C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WH   = HexColor("#E6E5DE")
C_MID_GY   = C_BLACK    # map secondary text to black
C_DRK_GY   = C_BLACK    # map body text to black
C_LGT_GY   = C_OFF_WH   # map panel background to off-white
C_BDR_GY   = C_TEAL     # map border to teal
C_LBL_GY   = C_TEAL     # map label/footer text to teal
C_CCL_GY   = C_TEAL     # map separator to teal


def rl_y(screen_y):
    """Convert top-origin y to ReportLab bottom-origin y."""
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def draw_text(c, text, x, y_top, font_name="Helvetica", font_size=12,
              colour=C_BLACK, align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font_name, font_size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font_name, font_size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font_name, font_size)
        x = x + (max_width - tw) / 2
    c.drawString(x, rl_y(y_top + font_size * 1.2), text)


def draw_wrapped_text(c, text, x, y_top, max_width, font_name="Helvetica",
                      font_size=10, colour=C_BLACK, line_spacing=1.4):
    words = text.split()
    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip()
        if c.stringWidth(test, font_name, font_size) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    c.setFillColor(colour)
    c.setFont(font_name, font_size)
    y = rl_y(y_top + font_size * 1.2)
    lh = font_size * line_spacing
    for ln in lines:
        c.drawString(x, y, ln)
        y -= lh
    return y_top + len(lines) * lh / 72 * 72


def house_header(c, eyebrow):
    """Draw house style header: amber left stripe + black header band."""
    # White page background (already set)
    # Amber left stripe
    fill_rect(c, 0, 0, 9.6, H, C_AMBER_D)
    # Black header band
    fill_rect(c, 9.6, 0, W - 9.6, 72, C_BLACK)
    # Eyebrow label
    draw_text(c, eyebrow, 18, 18, font_name="Helvetica-Bold", font_size=11,
              colour=C_OFF_WH)
    # PROFITPULSE right
    c.setFillColor(C_TEAL)
    c.setFont("Helvetica-Bold", 11)
    tw = c.stringWidth("PROFITPULSE", "Helvetica-Bold", 11)
    c.drawString(W - tw - 18, rl_y(18 + 13.2), "PROFITPULSE")


def house_footer(c, date_str="27 Jun 2026"):
    """Draw footer separator and text."""
    fill_rect(c, 9.6, H - 30, W - 9.6, 0.75, C_CCL_GY)
    draw_text(c, "Prepared by Nitesh Roopa CA  |  Managing Partner  |  ProfitPulse  |  Profit-Pulse.com.au",
              18, H - 29, font_name="Helvetica", font_size=7, colour=C_LBL_GY)
    c.setFillColor(C_LBL_GY)
    c.setFont("Helvetica", 7)
    tw = c.stringWidth(date_str, "Helvetica", 7)
    c.drawString(W - tw - 18, rl_y(H - 29 + 8.4), date_str)


def stat_card_pdf(c, x, y_top, w, h, number, label, source):
    """Draw a stat card: black tile, teal top stripe."""
    fill_rect(c, x, y_top, w, h, C_BLACK)
    fill_rect(c, x, y_top, w, 3, C_TEAL)
    # Number
    draw_text(c, number, x + 6, y_top + 8, font_name="Helvetica-Bold", font_size=22,
              colour=C_AMBER_B)
    # Label
    draw_wrapped_text(c, label, x + 6, y_top + 40, w - 12,
                      font_name="Helvetica", font_size=8, colour=C_OFF_WH, line_spacing=1.3)
    # Source
    draw_text(c, source, x + 6, y_top + h - 16, font_name="Helvetica-Oblique", font_size=6,
              colour=C_TEAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_27Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))

# Page 1 background
fill_rect(c, 0, 0, W, H, C_WHITE)
house_header(c, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
c.setFillColor(C_BLACK)
c.setFont("Helvetica-Bold", 38)
c.drawString(18, rl_y(78 + 44), "TASKFORCE AUSTRALIA")

# Descriptor
draw_text(c, "Field service management platform and property safety compliance, Burnley, Melbourne VIC",
          18, 124, font_name="Helvetica", font_size=10, colour=C_MID_GY)

# Stat cards
cards = [
    ("$12.8M", "Revenue FY2025", "Smart50 2025, rank 37"),
    ("31%", "Three year average growth", "Smart50 2025"),
    ("19", "Team members", "Smart50 2025"),
    ("5,500", "Tradespeople in network", "Smart50 2024/2025"),
    ("140K+", "Jobs via RentSafe since 2021", "Smart50 2025"),
    ("2014", "Year founded", "Company records"),
]
CARD_W = 149
CARD_H = 108
CARD_Y = 140
CARD_GAP = 5
card_x = 18
for num, lbl, src in cards:
    stat_card_pdf(c, card_x, CARD_Y, CARD_W, CARD_H, num, lbl, src)
    card_x += CARD_W + CARD_GAP

# Revenue bar chart
CX = 18
CY_CHART = 260
BAR_W = 65
BAR_MAX_H = 65
GAP = 10
# 2024 bar
h2024 = int(BAR_MAX_H * 9.7 / 12.8)
fill_rect(c, CX, CY_CHART, BAR_W, h2024, C_TEAL)
draw_text(c, "$9.7M", CX, CY_CHART - 12, font_name="Helvetica-Bold", font_size=8, colour=C_BLACK)
draw_text(c, "FY2024", CX + 18, CY_CHART + h2024 + 3, font_name="Helvetica", font_size=7, colour=C_MID_GY)
# 2025 bar
fill_rect(c, CX + BAR_W + GAP, CY_CHART, BAR_W, BAR_MAX_H, C_AMBER_D)
draw_text(c, "$12.8M", CX + BAR_W + GAP, CY_CHART - 12, font_name="Helvetica-Bold", font_size=8, colour=C_BLACK)
draw_text(c, "FY2025", CX + BAR_W + GAP + 14, CY_CHART + BAR_MAX_H + 3,
          font_name="Helvetica", font_size=7, colour=C_MID_GY)

draw_text(c, "REVENUE GROWTH (AUD, verified)",
          CX, CY_CHART - 24, font_name="Helvetica-Bold", font_size=7.5, colour=C_TEAL)
draw_text(c, "Source: SmartCompany Smart50 2024 and 2025 award citations",
          CX, CY_CHART + BAR_MAX_H + 16, font_name="Helvetica-Oblique", font_size=6, colour=C_TEAL)

# Key Commercial Signals
SX = 300
SY = 258
draw_text(c, "KEY COMMERCIAL SIGNALS",
          SX, SY - 14, font_name="Helvetica-Bold", font_size=7.5, colour=C_TEAL)

signals = [
    "•  RentRepair launched as Australia's only complete subscription property maintenance service "
    "(Essential and Comprehensive cover). (taskforce.com.au/rentrepair)",

    "•  140,000+ jobs delivered through RentSafe across 20 major consumer brands since 2021. "
    "Brands include Brilliant Lighting, Nero, Universal Fans. (SmartCompany Smart50 2025)",

    "•  Agency clients doubled in 12 months. Over 300 real estate offices and 180 major real "
    "estate brands now using the platform. (SmartCompany Smart50 2025)",

    "•  Three consecutive Smart50 appearances: 2023, 2024 rank 46, 2025 rank 37. Revenue climbed "
    "from $9.7M to $12.8M year on year. (SmartCompany)",

    "•  Won Most Innovative Proptech at the 2024 Proptech Awards, Sydney. (PropTech Australia 2024)",
]

sy = SY
for sig in signals:
    draw_wrapped_text(c, sig, SX, sy, W - SX - 18, font_name="Helvetica",
                      font_size=8.5, colour=C_BLACK, line_spacing=1.3)
    sy += 46

house_footer(c)
c.showPage()


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
house_header(c, "THE OPPORTUNITY")

draw_text(c, "Taskforce Australia: Three commercial observations from ProfitPulse",
          18, 78, font_name="Helvetica", font_size=11, colour=C_MID_GY)

obs = [
    {
        "bg": C_TEAL,
        "tc": C_WHITE,
        "nc": C_WHITE,
        "idx": "01",
        "hdr": "Revenue per person is exceptional. Margin per client is unknown.",
        "body": (
            "At $12.8 million across 19 employees, Taskforce generates $674,000 in revenue per person, "
            "placing it among the most efficient field service businesses in Australia. But the platform "
            "coordinates more than 5,500 external tradespeople, meaning a significant share of revenue "
            "flows through to subcontractor fees. Until each manufacturer client and each real estate office "
            "is ranked by gross margin contribution and coordination cost, growth is managed by revenue "
            "alone. A Customer Concentration and Profitability Map reveals which accounts to grow, which "
            "to reprice, and which consume more than they return."
        ),
    },
    {
        "bg": C_BLACK,
        "tc": C_OFF_WH,
        "nc": C_AMBER_B,
        "idx": "02",
        "hdr": "Two channels with different economics need separate financial discipline.",
        "body": (
            "RentSafe manages compliance for manufacturers and real estate offices on a per job model. "
            "RentRepair offers subscription maintenance at a fixed monthly fee. These are structurally "
            "different businesses within one entity: one is volume and frequency driven, the other is "
            "recurring revenue with customer lifetime value mechanics. A business that prices both "
            "against the same cost base will misread its margins in both directions. Understanding which "
            "channel funds growth and which consumes it is the foundational question before the next "
            "expansion phase."
        ),
    },
    {
        "bg": C_GOLD,
        "tc": C_BLACK,
        "nc": C_BLACK,
        "idx": "03",
        "hdr": "Doubling agency clients signals growth. Cash timing needs matching attention.",
        "body": (
            "Doubling agency clients in a year is a genuine commercial achievement. It also means more "
            "work in progress, more debtor cycles, and more subcontractor payments ahead of client "
            "settlement. As Taskforce moves into subscription billing through RentRepair, the cash "
            "conversion model changes again. A disciplined working capital review and a 13-week cash "
            "flow build are the natural next steps for a business that has outgrown its original revenue "
            "model and is building a second engine alongside the first."
        ),
    },
]

COL_START_Y = 98
COL_H_PDF = 380
COL_W_PDF = 305
COL_GAP_PDF = 5
col_x = 12.8

for o in obs:
    fill_rect(c, col_x, COL_START_Y, COL_W_PDF, COL_H_PDF, o["bg"])
    draw_text(c, o["idx"], col_x + 14, COL_START_Y + 14,
              font_name="Helvetica-Bold", font_size=28, colour=o["nc"])
    draw_wrapped_text(c, o["hdr"], col_x + 14, COL_START_Y + 60,
                      COL_W_PDF - 28, font_name="Helvetica-Bold", font_size=10.5,
                      colour=o["tc"], line_spacing=1.35)
    draw_wrapped_text(c, o["body"], col_x + 14, COL_START_Y + 115,
                      COL_W_PDF - 28, font_name="Helvetica", font_size=9,
                      colour=o["tc"], line_spacing=1.35)
    col_x += COL_W_PDF + COL_GAP_PDF

# Warm closing line
draw_wrapped_text(
    c,
    "These observations are offered in good faith. Taskforce Australia has built a remarkable platform. "
    "The question is simply whether the financial architecture is keeping pace with the commercial ambition.",
    18, COL_START_Y + COL_H_PDF + 10, W - 36,
    font_name="Helvetica-Oblique", font_size=9.5, colour=C_MID_GY, line_spacing=1.3
)

house_footer(c)
c.showPage()


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION
# ══════════════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
house_header(c, "THE RECOMMENDATION")

# Left column
LX = 18
LW = 540
LY = 84

draw_text(c, "Customer Concentration and Profitability Map",
          LX, LY, font_name="Helvetica-Bold", font_size=16, colour=C_BLACK)

draw_text(c, "$3,950  one off  |  ProfitPulse verified price",
          LX, LY + 26, font_name="Helvetica", font_size=11, colour=C_TEAL)

draw_wrapped_text(
    c,
    "A three week project ranking every customer by revenue, gross margin contribution, and effort "
    "to serve. For Taskforce Australia, this maps both manufacturer clients and real estate agency "
    "accounts to reveal which drive the real margin after tradie fees, platform costs, and "
    "coordination effort are allocated. Output: a clear action list.",
    LX, LY + 50, LW - 10, font_name="Helvetica", font_size=9.5, colour=C_DRK_GY, line_spacing=1.4
)

fill_rect(c, LX, LY + 128, LW - 10, 1.5, C_TEAL)

STEP_Y_PDF = LY + 135
draw_text(c, "STEP ONE", LX, STEP_Y_PDF, font_name="Helvetica-Bold", font_size=8, colour=C_TEAL)
draw_text(c, "Answer a few quick questions",
          LX, STEP_Y_PDF + 14, font_name="Helvetica-Bold", font_size=13, colour=C_BLACK)
draw_text(c, "See the solutions matched to your size and industry",
          LX, STEP_Y_PDF + 31, font_name="Helvetica", font_size=10, colour=C_DRK_GY)

# Questionnaire address as visible text (clean, no UTM)
c.setFillColor(C_TEAL)
c.setFont("Helvetica", 10)
c.drawString(LX, rl_y(STEP_Y_PDF + 50 + 12), "profit-pulse.com.au/full-suite-of-products")

fill_rect(c, LX, STEP_Y_PDF + 58, LW - 10, 1, C_AMBER_D)

# Direct CTA text (Stripe link not shown as text)
draw_text(c, "Purchase the suggested product now to get started",
          LX, STEP_Y_PDF + 64, font_name="Helvetica-Bold", font_size=11, colour=C_AMBER_D)

draw_text(c, "Prefer a conversation first?",
          LX, STEP_Y_PDF + 88, font_name="Helvetica", font_size=9, colour=C_MID_GY)
draw_text(c, "Book a complimentary discovery call",
          LX, STEP_Y_PDF + 102, font_name="Helvetica", font_size=10, colour=C_TEAL)

# Right column credibility panel
RX = 575
RW = 372
fill_rect(c, RX, 75, RW, H - 75 - 28, C_LGT_GY)
fill_rect(c, RX, 75, RW, 3, C_TEAL)

PX_PDF = RX + 16
PW_PDF = RW - 32

draw_text(c, "Nitesh Roopa", PX_PDF, 80,
          font_name="Helvetica-Bold", font_size=18, colour=C_BLACK)
draw_text(c, "CA, Managing Partner  |  ProfitPulse", PX_PDF, 106,
          font_name="Helvetica", font_size=10, colour=C_TEAL)

fill_rect(c, PX_PDF, 122, PW_PDF, 1, C_TEAL)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion (Cahora Bassa, Mozambique)",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "CA qualification: SAICA, South Africa",
    "Practice: PwC South Africa, Nedbank CIB, QIC, ProfitPulse",
]
cy_pdf = 128
for cp in cred:
    draw_wrapped_text(c, "  " + cp, PX_PDF, cy_pdf, PW_PDF,
                      font_name="Helvetica", font_size=8.5, colour=C_DRK_GY, line_spacing=1.3)
    cy_pdf += 26

fill_rect(c, PX_PDF, cy_pdf + 4, PW_PDF, 1, C_AMBER_D)
cy_pdf += 10
for line in ["Profit-Pulse.com.au", "Nitesh@Profit-Pulse.com.au",
             "+61 411 876 267", "linkedin.com/in/nitesh-roopa-77594163"]:
    draw_text(c, line, PX_PDF, cy_pdf, font_name="Helvetica", font_size=8.5,
              colour=C_DRK_GY)
    cy_pdf += 16

house_footer(c)
c.showPage()
c.save()
print(f"PDF saved: {out_path}")
