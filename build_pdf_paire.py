"""
ProfitPulse Brief PDF Builder
Target: Paire | Date: 10 Jun 2026
White background house style, 3 slides, brand colours only, zero dashes.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.units import inch

# Page dimensions
W = 960   # pt
H = 540   # pt

# Brand colours
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_MID_GREY  = HexColor("#888888")
C_DRK_GREY  = HexColor("#444444")


# ── Coordinate helpers ────────────────────────────────────────────────────────
def rl_y(screen_y):
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def fill_stroke_rect(c, x, y_top, w, h, fill_colour, stroke_colour, line_width=1):
    c.setFillColor(fill_colour)
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=1)


def stroke_rect(c, x, y_top, w, h, stroke_colour, line_width=1):
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold",
        align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline_y = rl_y(y_top + size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline_y, text)


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica",
                leading=None, align="left"):
    if leading is None:
        leading = size * 1.45
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip()
        if c.stringWidth(test, font, size) <= max_w:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    y = y_top
    for line in lines:
        if align == "right":
            tw = c.stringWidth(line, font, size)
            c.drawString(x + max_w - tw, rl_y(y + size), line)
        else:
            c.drawString(x, rl_y(y + size), line)
        y += leading
    return y


def hline(c, x, y_top, w, colour, thickness=1.0):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def vline(c, x, y_top, h, colour, thickness=1.0):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, rl_y(y_top), x, rl_y(y_top + h))


def stat_card(c, x, y, w, h, number, label1, label2, source):
    """Black tile with thin teal top stripe."""
    fill_rect(c, x, y, w, h, C_BLACK)
    fill_rect(c, x, y, w, 4, C_TEAL)
    txt(c, number,   x + 8, y + 7,  22, C_WHITE,     font="Helvetica-Bold")
    if label1:
        txt(c, label1, x + 8, y + 34, 9,  C_OFF_WHITE, font="Helvetica")
    if label2:
        txt(c, label2, x + 8, y + 46, 9,  C_OFF_WHITE, font="Helvetica")
    if source:
        txt(c, source, x + 8, y + h - 12, 7, C_MID_GREY, font="Helvetica-Oblique")


# ── Output ────────────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_10Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Paire | ProfitPulse Brief | 10 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Product and Service Line Profitability")

DATE_LABEL  = "10 Jun 2026"
FOOTER_LEFT = "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au"

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 6, H, C_AMBER_B)          # amber left stripe
fill_rect(c, 6, 0, W - 6, 68, C_BLACK)       # header band

txt(c, "COMMERCIAL INTELLIGENCE BRIEF", 16, 14, 10, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "PAIRE", 16, 14, 10, C_OFF_WHITE, font="Helvetica-Bold",
    align="right", max_width=928)

# Company name and descriptor
txt(c, "Paire",   16, 78, 38, C_BLACK, font="Helvetica-Bold")
txt(c, "Sustainable direct to consumer apparel brand, South Melbourne VIC",
    16, 124, 11, C_TEAL, font="Helvetica")

# Section label
txt(c, "KEY METRICS", 16, 148, 10, C_TEAL, font="Helvetica-Bold")

# 5 stat cards (width ~174pt each, gap ~9pt, y=164)
CARD_W = 174
CARD_H = 96
CARD_Y = 164
cx_list = [16, 199, 382, 565, 748]
card_data = [
    ("$10M",  "Revenue FY2025",      "",                           "SmartCompany Smart50 2025"),
    ("64%",   "Revenue growth",      "three year average",         "SmartCompany Smart50 2025"),
    ("26",    "Employees",           "16 FTE plus 10 casual",      "SmartCompany Jul 2025"),
    ("15",    "International wholesale", "stores SG and MY",       "businessnewsaustralia.com 2025"),
    ("70%",   "Repeat customer rate","brand loyalty metric",       "SmartCompany 2025"),
]
for (num, l1, l2, src), cx in zip(card_data, cx_list):
    stat_card(c, cx, CARD_Y, CARD_W, CARD_H, num, l1, l2, src)

# Revenue trajectory chart (bottom left)
txt(c, "REVENUE TRAJECTORY", 16, 272, 10, C_TEAL, font="Helvetica-Bold")

CHART_X_START = 50
CHART_BOTTOM  = 460
CHART_MAX_H   = 142  # pt for $10M

bars_data = [
    ("FY21", 1.0,  "$1M"),
    ("FY22", 1.7,  "$1.7M"),
    ("FY24", 6.0,  "$6M"),
    ("FY25", 10.0, "$10M"),
]
bar_slot = 90
bar_w    = 56

for i, (yr, val, lbl) in enumerate(bars_data):
    bx = CHART_X_START + i * bar_slot
    bh = int(val / 10.0 * CHART_MAX_H)
    by = CHART_BOTTOM - bh
    fill_rect(c, bx, by, bar_w, bh, C_TEAL)
    txt(c, lbl, bx, by - 17, 9, C_BLACK, font="Helvetica-Bold",
        align="center", max_width=bar_w)
    txt(c, yr,  bx, CHART_BOTTOM + 6, 9, C_DRK_GREY, font="Helvetica",
        align="center", max_width=bar_w)

hline(c, 30, CHART_BOTTOM, 400, C_DRK_GREY, 1.0)
txt(c, "Sources: SmartCompany Smart50 2025, SmartCompany articles and Shark Tank recap 2024",
    16, 480, 7, C_MID_GREY, font="Helvetica-Oblique")

# Key Commercial Signals (right side)
txt(c, "KEY COMMERCIAL SIGNALS", 436, 272, 10, C_TEAL, font="Helvetica-Bold")

signals = [
    ("Smart50 2025 rank 15 confirms revenue above AUD 5M eligibility threshold",
     "SmartCompany Smart50 2025"),
    ("Permanent bricks and mortar retail: QV Melbourne CBD flagship and South Melbourne store",
     "Ragtrader 2025 / Inside Retail Australia 2025"),
    ("International wholesale: 12 Boarding Gate stores Singapore, 3 Planet Traveller stores Malaysia",
     "businessnewsaustralia.com 2025"),
    ("Raised AUD 500,000 for 4 percent equity on Shark Tank Australia, implied valuation AUD 12.5M",
     "SmartCompany Shark Tank recap Nov 2024"),
    ("CEO Nathan Yun publicly stated focus is now profitability over hyper growth",
     "SmartCompany 2025"),
    ("20,000 units of underwear sold monthly, confirming category scale in one product line",
     "SmartCompany 2025"),
]

sy = 290
for sig, src in signals:
    fill_rect(c, 436, sy + 4, 6, 6, C_TEAL)
    txt_wrapped(c, sig, 448, sy, 496, 10, C_BLACK, font="Helvetica", leading=14)
    txt(c, src, 448, sy + 18, 7, C_MID_GREY, font="Helvetica-Oblique")
    sy += 40

# Footer
hline(c, 6, 508, W - 6, C_DRK_GREY, 0.8)
txt(c, FOOTER_LEFT, 16, 516, 8, C_DRK_GREY, font="Helvetica")
txt(c, DATE_LABEL, 16, 516, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=928)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 6, H, C_AMBER_B)
fill_rect(c, 6, 0, W - 6, 68, C_BLACK)

txt(c, "THE OPPORTUNITY", 16, 8, 10, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "Paire: three commercial observations from ProfitPulse",
    16, 38, 10, C_MID_GREY, font="Helvetica")
txt(c, "PAIRE", 16, 8, 10, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=928)

# Three columns
COL_W = 318
COL_Y = 68
COL_H = 390

obs_data = [
    {
        "bg": C_TEAL, "num_col": C_BLACK, "txt_col": C_BLACK,
        "idx": "01",
        "header": "Four channels, one P and L: where is the margin?",
        "body": (
            "Paire now earns revenue across four distinct commercial channels: "
            "DTC online, QV Melbourne CBD store, South Melbourne store, and wholesale "
            "export across 15 retail doors in Singapore and Malaysia. Each channel "
            "carries a materially different cost and margin structure. DTC online "
            "captures the highest margin per unit with no intermediary. Physical "
            "retail adds fixed lease, staff, and fit-out costs. Wholesale reduces "
            "per-unit margin but opens volume at scale. Without a channel by channel "
            "profitability view, growth capital flows to the highest revenue source "
            "rather than the highest return source."
        ),
    },
    {
        "bg": C_BLACK, "num_col": C_AMBER_B, "txt_col": C_OFF_WHITE,
        "idx": "02",
        "header": "Six product categories and one question: which ones win?",
        "body": (
            "Paire launched in 2020 with a single product and now sells socks, "
            "underwear, bras, tops, leggings, activewear, and loungewear. Selling "
            "20,000 units of underwear each month signals genuine scale in at least "
            "one category. In DTC apparel brands at this stage, a small share of "
            "product lines typically drives the majority of gross margin while "
            "others generate volume at low or zero contribution. Knowing which "
            "categories to scale, which to reprice, and which are diluting return "
            "on working capital is the foundation for every inventory, marketing, "
            "and expansion decision the business faces now."
        ),
    },
    {
        "bg": C_GOLD, "num_col": C_BLACK, "txt_col": C_BLACK,
        "idx": "03",
        "header": "Investor backed business needs a margin map, not just a revenue line",
        "body": (
            "Raising AUD 500,000 for 4 percent equity on Shark Tank Australia "
            "created an implied valuation of AUD 12.5 million and attached investor "
            "expectations. The sharks called it the most successful company to appear "
            "on Shark Tank Australia. CEO Nathan Yun has publicly stated the next "
            "chapter is profitability over hyper growth. That pivot requires a "
            "financial architecture that answers: which products carry the business, "
            "which channels deliver real return, and where does capital generate the "
            "highest ROI. A product and channel profitability map is the prerequisite "
            "for that conversation."
        ),
    },
]

for i, obs in enumerate(obs_data):
    ox = 6 + i * (COL_W + 2)
    fill_rect(c, ox, COL_Y, COL_W, COL_H, obs["bg"])
    txt(c, obs["idx"],  ox + 14, COL_Y + 14, 32, obs["num_col"], font="Helvetica-Bold")
    txt_wrapped(c, obs["header"], ox + 14, COL_Y + 62, COL_W - 28,
                12, obs["txt_col"], font="Helvetica-Bold", leading=17)
    txt_wrapped(c, obs["body"], ox + 14, COL_Y + 108, COL_W - 28,
                10, obs["txt_col"], font="Helvetica", leading=15)

# Warm line
hline(c, 6, 466, W - 6, C_DRK_GREY, 0.8)
txt_wrapped(c,
    "These are observations offered in good faith. Paire has built something genuinely impressive. "
    "The question is simply whether the financial architecture now matches the ambition.",
    16, 474, W - 32, 10, C_DRK_GREY, font="Helvetica-Oblique", leading=14)

# Footer
hline(c, 6, 508, W - 6, C_DRK_GREY, 0.8)
txt(c, FOOTER_LEFT, 16, 516, 8, C_DRK_GREY, font="Helvetica")
txt(c, DATE_LABEL, 16, 516, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=928)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 6, H, C_AMBER_B)
fill_rect(c, 6, 0, W - 6, 68, C_BLACK)

txt(c, "THE RECOMMENDATION", 16, 14, 10, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "PAIRE", 16, 14, 10, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=928)

# Vertical divider
vline(c, 590, 68, 432, C_AMBER_D, 1.0)

# LEFT column
LX = 16
LW = 562

# Service name and price
txt(c, "Product and Service Line Profitability",
    LX, 78, 18, C_BLACK, font="Helvetica-Bold")
txt(c, "$3,950 one off", LX, 106, 13, C_TEAL, font="Helvetica-Bold")

txt_wrapped(c,
    "A three week project ranking every product line and revenue channel by gross "
    "margin, contribution margin, and operational drag, with a kill, fix, or scale "
    "decision on each. For Paire: maps six product categories across four channels "
    "and shows where the Shark Tank capital generates the highest return.",
    LX, 128, LW, 11, C_BLACK, font="Helvetica", leading=16)

hline(c, LX, 196, LW, C_TEAL, 1)

txt(c, "STEP ONE", LX, 204, 9, C_TEAL, font="Helvetica-Bold")
txt(c, "Answer a few quick questions", LX, 222, 13, C_BLACK, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry",
    LX, 244, 11, C_DRK_GREY, font="Helvetica")
txt(c, "profit-pulse.com.au/full-suite-of-products",
    LX, 263, 12, C_TEAL, font="Helvetica")

hline(c, LX, 288, LW, C_DRK_GREY, 0.8)

# Link annotation for questionnaire
c.saveState()
c.setFillColor(Color(0, 0, 0, alpha=0))
c.rect(LX, rl_y(288), LW, 20, fill=0, stroke=0)
c.linkURL("https://profit-pulse.com.au/full-suite-of-products",
          (LX, rl_y(290), LX + 380, rl_y(260)), relative=0)
c.restoreState()

# Direct CTA with hyperlink annotation
txt(c, "Purchase the suggested product now to get started",
    LX, 298, 12, C_AMBER_D, font="Helvetica-Bold")
c.linkURL("https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
          (LX, rl_y(316), LX + 460, rl_y(290)), relative=0)

hline(c, LX, 328, LW, C_DRK_GREY, 0.8)
txt(c, "Prefer a conversation first?", LX, 336, 10, C_DRK_GREY, font="Helvetica")
txt(c, "Book a complimentary discovery call",
    LX, 356, 11, C_TEAL, font="Helvetica-Bold")
c.linkURL(
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
    (LX, rl_y(374), LX + 310, rl_y(348)), relative=0)

txt(c, "Supporting services: C2 Working Capital Unlock  |  D2 Pricing Reset",
    LX, 392, 9, C_MID_GREY, font="Helvetica")

# RIGHT column: Credibility panel
RX  = 606
RW  = W - RX - 14

txt(c, "Nitesh Roopa", RX, 78, 17, C_BLACK, font="Helvetica-Bold")
txt(c, "CA, Managing Partner", RX, 106, 11, C_DRK_GREY, font="Helvetica")
txt(c, "ProfitPulse", RX, 124, 14, C_TEAL, font="Helvetica-Bold")
hline(c, RX, 148, RW, C_TEAL, 1.2)

cred_items = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Total GRBT value: over AUD 10 billion",
    "Largest deal: USD 1.3B (Cahora Bassa)",
]
cy_r = 158
for item in cred_items:
    fill_rect(c, RX, cy_r + 3, 6, 6, C_TEAL)
    txt(c, item, RX + 12, cy_r, 10, C_BLACK, font="Helvetica")
    cy_r += 24

hline(c, RX, cy_r + 4, RW, C_DRK_GREY, 0.8)
cy_r += 14

contact_data = [
    ("Profit-Pulse.com.au",              C_BLACK),
    ("Nitesh@Profit-Pulse.com.au",       C_TEAL),
    ("+61 411 876 267",                  C_BLACK),
    ("linkedin.com/in/nitesh-roopa-77594163", C_DRK_GREY),
]
for ctext, ccol in contact_data:
    txt(c, ctext, RX, cy_r, 10, ccol, font="Helvetica")
    cy_r += 22

# Footer
hline(c, 6, 508, W - 6, C_DRK_GREY, 0.8)
txt(c, FOOTER_LEFT, 16, 516, 8, C_DRK_GREY, font="Helvetica")
txt(c, DATE_LABEL, 16, 516, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=928)

# ── Save ──────────────────────────────────────────────────────────────────────
c.save()
print(f"PDF saved: {out_path}")
