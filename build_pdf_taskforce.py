"""
ProfitPulse Brief PDF Builder | Taskforce Australia | 23 Jun 2026
House style: white body, amber left stripe, black header band.
Page: 960pt x 540pt (13.333in x 7.5in at 72dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color, white, black
from reportlab.pdfbase import pdfmetrics

W = 960
H = 540

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_W    = HexColor("#E6E5DE")
C_MID_GRY  = HexColor("#888888")
C_DRK_GRY  = HexColor("#444444")
C_LT_GRY   = HexColor("#F8F8F8")
C_TEAL_HL  = HexColor("#01A296")


def ry(y_top, h=0):
    return H - y_top - h


def fill(c, x, y_top, w, h, col):
    c.setFillColor(col)
    c.rect(x, ry(y_top, h), w, h, fill=1, stroke=0)


def stroke_fill(c, x, y_top, w, h, fc, sc, lw=1):
    c.setFillColor(fc)
    c.setStrokeColor(sc)
    c.setLineWidth(lw)
    c.rect(x, ry(y_top, h), w, h, fill=1, stroke=1)


def hln(c, x, y_top, w, col, lw=1.5):
    c.setStrokeColor(col)
    c.setLineWidth(lw)
    c.line(x, ry(y_top), x + w, ry(y_top))


def text(c, s, x, y_top, size, col, font="Helvetica", align="left", max_w=None):
    c.setFillColor(col)
    c.setFont(font, size)
    bline = ry(y_top + size * 0.85)
    if align == "right" and max_w:
        tw = c.stringWidth(s, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(s, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, bline, s)


def wrap_text(c, s, x, y_top, max_w, size, col, font="Helvetica", leading=None):
    if leading is None:
        leading = size * 1.5
    c.setFillColor(col)
    c.setFont(font, size)
    words = s.split()
    lines, cur = [], ""
    for w_word in words:
        test = (cur + " " + w_word).strip()
        if c.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w_word
    if cur:
        lines.append(cur)
    y = y_top
    for line in lines:
        c.drawString(x, ry(y + size * 0.85), line)
        y += leading
    return y


def page_header(c, eyebrow, right="PROFITPULSE"):
    fill(c, 6, 0, W - 6, 55, C_BLACK)
    text(c, eyebrow, 20, 10, 11, C_OFF_W, "Helvetica-Bold")
    text(c, right, 640, 10, 11, C_TEAL, "Helvetica-Bold", align="right", max_w=300)


def page_footer(c, date="23 Jun 2026"):
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         20, 522, 9, C_DRK_GRY, "Helvetica")
    text(c, date, 710, 522, 9, C_DRK_GRY, "Helvetica",
         align="right", max_w=230)


OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_23Jun2026.pdf"
c = canvas.Canvas(OUT, pagesize=(W, H))
c.setTitle("Taskforce Australia | ProfitPulse Brief | 23 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Operational Intelligence Review")


# =========================================================================
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# =========================================================================
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_D)
page_header(c, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
text(c, "Taskforce Australia", 20, 75, 42, C_BLACK, "Helvetica-Bold")

# Descriptor
text(c, "Property maintenance and safety compliance platform  |  Burnley, Melbourne VIC",
     20, 128, 13, C_TEAL, "Helvetica")

# -- Stat cards ---------------------------------------------------------------
STATS = [
    ("$12.8M", "Revenue",       "FY25",        "Smart50 2025"),
    ("31%",    "Revenue",       "Growth Rate", "Smart50 2025"),
    ("19",     "In-House",      "Team",        "Smart50 2025"),
    ("5,500",  "Trade",         "Network",     "Smart50 2025"),
    ("618%",   "Housing Jobs",  "YoY Growth",  "Smart50 2025"),
    ("#37",    "Smart50",       "2025 Rank",   "SmartCompany 2025"),
]

CARD_W   = 150
CARD_H   = 100
CARD_GAP = 5
CX0      = 20

for i, (num, l1, l2, src) in enumerate(STATS):
    cx = CX0 + i * (CARD_W + CARD_GAP)
    fill(c, cx, 148, CARD_W, CARD_H, C_BLACK)
    fill(c, cx, 148, CARD_W, 5, C_TEAL)
    text(c, num,  cx + 8, 160, 24, C_AMBER_B, "Helvetica-Bold")
    text(c, l1,   cx + 8, 195, 10, C_OFF_W,   "Helvetica")
    text(c, l2,   cx + 8, 210, 10, C_OFF_W,   "Helvetica")
    text(c, src,  cx + 8, 232, 7,  C_MID_GRY, "Helvetica-Oblique")

# -- Revenue chart ------------------------------------------------------------
CHX = 680
CHY = 270
CHW = 265
CHH = 215

text(c, "REVENUE GROWTH", CHX, CHY - 14, 10, C_TEAL, "Helvetica-Bold")
fill(c, CHX, CHY, CHW, CHH, HexColor("#F2F2F2"))

BAR_MAX_H = 145
BAR_BASE  = CHY + CHH - 30
BAR_W     = 82

bx = CHX + 25
for yr_lbl, rev, bar_col in [("FY24", 9.7, C_AMBER_D), ("FY25", 12.8, C_TEAL)]:
    bh = int(BAR_MAX_H * (rev / 14.0))
    by = BAR_BASE - bh
    fill(c, bx, by, BAR_W, bh, bar_col)
    text(c, f"${rev}M", bx, by - 16, 9, C_BLACK, "Helvetica-Bold",
         align="center", max_w=BAR_W)
    text(c, yr_lbl, bx, BAR_BASE + 6, 9, C_DRK_GRY, "Helvetica",
         align="center", max_w=BAR_W)
    bx += BAR_W + 20

text(c, "Source: Smart50 2024 and 2025, SmartCompany",
     CHX, CHY + CHH + 10, 7, C_MID_GRY, "Helvetica-Oblique")

# -- Key commercial signals ---------------------------------------------------
text(c, "KEY COMMERCIAL SIGNALS", 20, 270, 10, C_TEAL, "Helvetica-Bold")

SIGS = [
    "Housing jobs grew 618% year-on-year, making housing the dominant service area  (Smart50 2025)",
    "Housing division confirmed as strongest and most profitable area  (Smart50 2025, SmartCompany)",
    "Equity plan for staff and customers in active development  (Smart50 2025)",
    "AI-enhanced scheduling platform expansion underway, four-point growth agenda  (Smart50 2025)",
    "Three consecutive Smart50 appearances: 2023, 2024, and 2025  (SmartCompany)",
]

sy = 290
for sig in SIGS:
    wrap_text(c, "•  " + sig, 20, sy, 640, 10, C_BLACK, "Helvetica", leading=14)
    sy += 36

page_footer(c)
c.showPage()


# =========================================================================
# PAGE 2: THE OPPORTUNITY — THREE COMMERCIAL OBSERVATIONS
# =========================================================================
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_D)
page_header(c, "THE OPPORTUNITY", "Taskforce Australia")

text(c, "Taskforce Australia  |  Three commercial observations from ProfitPulse",
     20, 63, 11, C_DRK_GRY, "Helvetica-Oblique")

OBS2 = [
    (C_TEAL, C_WHITE,
     "01",
     "618 percent growth in one service line reshapes the financial architecture",
     ("When housing jobs grow 618 percent in a single year the revenue mix, cost "
      "structure, and customer concentration profile all change simultaneously. "
      "The housing division is the acknowledged margin leader, but at this velocity "
      "the question shifts from which division is strongest to how strong, by how "
      "much, and on what terms. Without a service line profitability map allocating "
      "sub-contractor costs, technology investment, and coordination time to each "
      "job type, the income statement cannot tell that story. Knowing the real "
      "housing unit economics now protects against growing into the wrong margin "
      "structure.")),
    (C_BLACK, C_OFF_W,
     "02",
     "5,500 tradespeople managed by 19 people creates leverage and concentration risk simultaneously",
     ("Nineteen people generating $12.8 million creates revenue per head of "
      "approximately $674,000. That is the signature of an asset-light platform. "
      "The platform economics are real. But they are only sustainable if the "
      "customer base is sufficiently diversified and all service lines clear their "
      "contribution hurdles after full cost allocation. As housing clients grow "
      "to dominate job volume, the concentration of revenue in a small number of "
      "housing providers becomes the primary financial risk. A customer "
      "concentration map would quantify that exposure and identify which accounts "
      "are both large and genuinely high margin.")),
    (C_GOLD, C_BLACK,
     "03",
     "The equity plan and growth agenda need a financial foundation before activation",
     ("Taskforce has stated it will create an equity plan for staff and customers. "
      "Equity plans work when anchored to a clear view of business value, margin "
      "quality, and growth trajectory. That foundation requires knowing where "
      "earnings actually come from, which service lines and customers drive the "
      "repeatable margin, and what the valuation story looks like at exit or "
      "restructuring. Building that picture now, while the housing division is "
      "growing at 618 percent and before the equity plan is finalised, means the "
      "plan is anchored to the right numbers from the start.")),
]

COL_W2  = 297
COL_GAP2 = 8
COL_X0_2 = 20
COL_TOP2 = 80
COL_H2   = 390

for i, (bg, tc, num2, hdg, body2) in enumerate(OBS2):
    cx2 = COL_X0_2 + i * (COL_W2 + COL_GAP2)
    fill(c, cx2, COL_TOP2, COL_W2, COL_H2, bg)
    text(c, num2, cx2 + 12, COL_TOP2 + 10, 24, tc, "Helvetica-Bold")
    wrap_text(c, hdg, cx2 + 12, COL_TOP2 + 48, COL_W2 - 24, 11, tc,
              "Helvetica-Bold", leading=15)
    wrap_text(c, body2, cx2 + 12, COL_TOP2 + 118, COL_W2 - 24, 9.5, tc,
              "Helvetica", leading=13)

wrap_text(c,
          "These observations are offered in good faith. Taskforce has built something "
          "genuinely impressive in property services, and the 618 percent housing growth "
          "signal is one of the strongest single-year moves on any Smart50 list this year. "
          "The question is simply whether the financial architecture now matches the ambition.",
          20, 482, W - 40, 9, C_DRK_GRY, "Helvetica-Oblique", leading=12)

page_footer(c)
c.showPage()


# =========================================================================
# PAGE 3: THE RECOMMENDATION AND HOW TO START
# =========================================================================
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_D)
page_header(c, "THE RECOMMENDATION", "Taskforce Australia")

Q_URL_CLEAN = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_URL2 = "https://buy.stripe.com/bJe00kc5agKQepG0ZV3ks1x"
BOOK_URL2   = ("https://bookings.cloud.microsoft/book/"
               "ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# -- LEFT COLUMN (CTAs) -------------------------------------------------------
LX2 = 20
LW2 = 555
LY2 = 62

text(c, "Operational Intelligence Review", LX2, LY2, 22, C_BLACK, "Helvetica-Bold")
text(c, "$6,500 one off", LX2, LY2 + 32, 14, C_TEAL, "Helvetica-Bold")

wrap_text(c,
          "A six week deep dive across customer concentration and profitability, "
          "product and service line margin, workforce capacity, and operational "
          "bottlenecks. Delivers a prioritised 12 month action list calibrated "
          "to the specific growth signals visible in this business.",
          LX2, LY2 + 56, LW2, 10, C_DRK_GRY, "Helvetica", leading=14)

# Step 1 box
S1Y = LY2 + 110
S1H = 100
fill(c, LX2, S1Y, LW2, S1H, C_TEAL)
text(c, "Step one: answer a few quick questions",
     LX2 + 10, S1Y + 10, 12, C_BLACK, "Helvetica-Bold")
text(c, "See the solutions matched to your size and industry",
     LX2 + 10, S1Y + 32, 10, C_BLACK, "Helvetica")

# Clean questionnaire URL (visible, hyperlinked)
text(c, "profit-pulse.com.au/full-suite-of-products",
     LX2 + 10, S1Y + 56, 11, C_BLACK, "Helvetica-Bold")
c.linkURL(Q_URL_CLEAN,
          (LX2 + 10, ry(S1Y + 56 + 12), LX2 + 10 + 310, ry(S1Y + 56 - 2)),
          relative=0)

# Direct purchase
fill(c, LX2, S1Y + S1H + 8, LW2, 55, HexColor("#FFFBF0"))
stroke_fill(c, LX2, S1Y + S1H + 8, LW2, 55, HexColor("#FFFBF0"), C_AMBER_D, 1)

text(c, "Already know this is the priority? Operational Intelligence Review, $6,500 one off.",
     LX2 + 10, S1Y + S1H + 16, 9, C_DRK_GRY, "Helvetica")

# "Purchase" words carry Stripe link (Stripe URL never shown as text)
text(c, "Purchase the suggested product now to get started",
     LX2 + 10, S1Y + S1H + 36, 11, C_AMBER_D, "Helvetica-Bold")
c.linkURL(STRIPE_URL2,
          (LX2 + 10, ry(S1Y + S1H + 36 + 13), LX2 + 10 + 318, ry(S1Y + S1H + 36 - 2)),
          relative=0)

# Discovery call
text(c, "Prefer a conversation first?",
     LX2, S1Y + S1H + 78, 10, C_DRK_GRY, "Helvetica")
text(c, "Book a complimentary discovery call",
     LX2, S1Y + S1H + 96, 11, C_TEAL, "Helvetica-Bold")
c.linkURL(BOOK_URL2,
          (LX2, ry(S1Y + S1H + 96 + 12), LX2 + 220, ry(S1Y + S1H + 96 - 2)),
          relative=0)

# -- RIGHT COLUMN (credentials) -----------------------------------------------
RX2 = 590
RW2 = 358
RY2 = 62
RH2 = 430

fill(c, RX2, RY2, RW2, RH2, C_LT_GRY)
fill(c, RX2, RY2, 5, RH2, C_TEAL)

text(c, "Nitesh Roopa", RX2 + 14, RY2 + 10, 20, C_BLACK, "Helvetica-Bold")
text(c, "CA, Managing Partner", RX2 + 14, RY2 + 40, 12, C_DRK_GRY, "Helvetica")
text(c, "ProfitPulse", RX2 + 14, RY2 + 60, 15, C_TEAL, "Helvetica-Bold")
hln(c, RX2 + 14, RY2 + 88, RW2 - 28, C_TEAL, 1.5)

CRS = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique",
    "Total GRBT project value: over AUD 10 billion",
]
cy2 = RY2 + 100
for cred in CRS:
    text(c, cred, RX2 + 14, cy2, 9.5, C_DRK_GRY, "Helvetica")
    cy2 += 18

hln(c, RX2 + 14, cy2 + 5, RW2 - 28, C_DRK_GRY, 1)
cy2 += 18

for ct, cc in [
    ("Profit-Pulse.com.au",                   C_DRK_GRY),
    ("Nitesh@Profit-Pulse.com.au",            C_TEAL),
    ("+61 411 876 267",                       C_DRK_GRY),
    ("linkedin.com/in/nitesh-roopa-77594163", C_MID_GRY),
]:
    text(c, ct, RX2 + 14, cy2, 9.5, cc, "Helvetica")
    cy2 += 18

page_footer(c)

c.save()
print(f"PDF saved: {OUT}")
