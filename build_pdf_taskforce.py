"""
Brief_TaskforceAustralia_12Jun2026.pdf
Three slides matching the PPTX layout, using ReportLab canvas.
Page size: 960pt x 540pt (widescreen 13.333" x 7.5" at 72dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.utils import simpleSplit
import os

OUT_DIR  = "/home/user/Daily-Outreach/Out-reach efforts"
OUT_FILE = os.path.join(OUT_DIR, "Brief_TaskforceAustralia_12Jun2026.pdf")

PW = 960.0
PH = 540.0

# Brand colours
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_L   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_MID_GREY  = HexColor("#999999")

def ry(y): return PH - y  # convert top-origin y to ReportLab bottom-origin

def fill_rect(cv, x, y, w, h, color):
    cv.setFillColor(color)
    cv.rect(x, ry(y + h), w, h, fill=1, stroke=0)

def txt(cv, s, x, y, size=10, bold=False, color=C_BLACK, align="left", max_w=None):
    cv.setFillColor(color)
    fn = "Helvetica-Bold" if bold else "Helvetica"
    cv.setFont(fn, size)
    if align == "right" and max_w:
        cv.drawRightString(x + max_w, ry(y + size * 0.78), s)
    elif align == "center" and max_w:
        cx = x + max_w / 2
        cv.drawCentredString(cx, ry(y + size * 0.78), s)
    else:
        cv.drawString(x, ry(y + size * 0.78), s)

def txt_wrap(cv, s, x, y, max_w, line_h, size=10, bold=False, color=C_BLACK, max_lines=20):
    cv.setFillColor(color)
    fn = "Helvetica-Bold" if bold else "Helvetica"
    cv.setFont(fn, size)
    lines = simpleSplit(s, fn, size, max_w)
    for i, line in enumerate(lines[:max_lines]):
        cv.drawString(x, ry(y + size * 0.78 + i * line_h), line)

def dot(cv, x, y, r=4, color=C_TEAL):
    cv.setFillColor(color)
    cv.circle(x, ry(y), r, fill=1, stroke=0)

c = canvas.Canvas(OUT_FILE, pagesize=(PW, PH))

# ═══════════════════════════════════════════════════════════════
# PAGE 1 — Prospect Profile
# ═══════════════════════════════════════════════════════════════

# White bg
fill_rect(c, 0, 0, PW, PH, C_WHITE)
# Amber left stripe
fill_rect(c, 0, 0, 7, PH, C_AMBER_D)
# Black header band
HDR_H = 72
fill_rect(c, 0, 0, PW, HDR_H, C_BLACK)

# Eyebrow labels
txt(c, "PROSPECT BRIEF", 16, 6, size=8, bold=True, color=C_OFF_WHITE)
txt(c, "TASKFORCE AUSTRALIA", 614, 6, size=8, bold=True, color=C_TEAL,
    align="right", max_w=340)

# Company name
txt(c, "Taskforce Australia", 16, 25, size=24, bold=True, color=C_WHITE)
# Date
txt(c, "12 Jun 2026", 614, 25, size=11, bold=False, color=C_OFF_WHITE,
    align="right", max_w=340)

# Sub header
txt(c, "Property maintenance technology platform  |  Melbourne VIC  |  Founded 2014",
    16, 78, size=10, bold=False, color=C_BLACK)

# Stat tiles (6 tiles)
TW  = 144
TH  = 112
TTY = 111
TGAP = 9
STRIPE = 4

stats = [
    ("$12.8M",  "Annual Revenue",        "Smart50 2025"),
    ("19",       "Team Members",          "Smart50 2025"),
    ("#37",      "Smart50 2025 Rank",     "SmartCompany"),
    ("31%",      "3Yr Avg Growth",        "Smart50 2025"),
    ("140K+",    "RentSafe Jobs",         "rentsafe.taskforce.com.au"),
    ("2014",     "Year Founded",          "Smart50 2023"),
]

for i, (num, lbl, src) in enumerate(stats):
    tx = 16 + i * (TW + TGAP)
    fill_rect(c, tx, TTY, TW, TH, C_BLACK)
    fill_rect(c, tx, TTY, TW, STRIPE, C_TEAL)
    txt(c, num,  tx, TTY + 8,  size=24, bold=True,  color=C_AMBER_L, align="center", max_w=TW)
    txt(c, lbl,  tx, TTY + 60, size=8,  bold=False, color=C_OFF_WHITE, align="center", max_w=TW)
    txt(c, src,  tx, TTY + 88, size=7,  bold=False, color=C_MID_GREY,  align="center", max_w=TW)

# Commercial signals
SIG_Y = 236
txt(c, "COMMERCIAL SIGNALS", 16, SIG_Y, size=8, bold=True, color=C_TEAL)

signals = [
    "Three consecutive Smart50 listings (2023, 2024, 2025); rank improved 46 to 37",
    "Stated growth target: 50 to 60 percent CAGR over three years, self funded",
    "RentRepair subscription maintenance launched at AREC 2023 from $54 per month",
    "RentSafe platform: 140,000+ jobs delivered across 300 real estate offices",
    "National network of over 5,000 tradespeople across Australia",
]

for j, sig in enumerate(signals):
    sy = SIG_Y + 22 + j * 38
    dot(c, 22, sy + 5, r=3.5)
    txt(c, sig, 32, sy, size=10, bold=False, color=C_BLACK)

# Bottom bar
fill_rect(c, 0, PH - 22, PW, 22, C_BLACK)
txt(c, "Prepared by ProfitPulse  |  Profit-Pulse.com.au  |  Confidential",
    16, PH - 21, size=7, bold=False, color=C_OFF_WHITE)
txt(c, "1 / 3", 854, PH - 21, size=7, bold=False, color=C_OFF_WHITE,
    align="right", max_w=90)

c.showPage()

# ═══════════════════════════════════════════════════════════════
# PAGE 2 — Three Observations
# ═══════════════════════════════════════════════════════════════

fill_rect(c, 0, 0, PW, PH, C_WHITE)
fill_rect(c, 0, 0, 7, PH, C_AMBER_D)
fill_rect(c, 0, 0, PW, HDR_H, C_BLACK)

txt(c, "OBSERVATIONS", 16, 6, size=8, bold=True, color=C_OFF_WHITE)
txt(c, "TASKFORCE AUSTRALIA", 614, 6, size=8, bold=True, color=C_TEAL,
    align="right", max_w=340)
txt(c, "Three signals that define the growth picture",
    16, 25, size=20, bold=True, color=C_WHITE)
txt(c, "12 Jun 2026", 614, 25, size=11, bold=False, color=C_OFF_WHITE,
    align="right", max_w=340)

COL_W  = 286
COL_H  = 396
COL_Y  = 80
C1X    = 14
C2X    = 336
C3X    = 658
LH10   = 14  # line height for 10pt

# Col 01 TEAL
fill_rect(c, C1X, COL_Y, COL_W, COL_H, C_TEAL)
txt(c, "01", C1X + 10, COL_Y + 10, size=28, bold=True, color=C_BLACK)
txt_wrap(c, "Platform leverage meets product complexity",
         C1X + 10, COL_Y + 48, COL_W - 20, 18, size=13, bold=True, color=C_BLACK)
obs1 = ("$12.8M revenue with 19 people signals strong platform leverage. "
        "The launch of RentRepair adds subscription economics alongside the "
        "existing compliance job model. These two revenue streams have "
        "different cost structures, different cash cycles and different margin "
        "profiles. Without a revenue stream map and margin by product line, "
        "the growth plan rests on an untested assumption that both will scale "
        "together without cannibalising cash or attention.")
txt_wrap(c, obs1, C1X + 10, COL_Y + 120, COL_W - 20, LH10, size=9, bold=False, color=C_BLACK)

# Col 02 BLACK
fill_rect(c, C2X, COL_Y, COL_W, COL_H, C_BLACK)
txt(c, "02", C2X + 10, COL_Y + 10, size=28, bold=True, color=C_AMBER_L)
txt_wrap(c, "Three rankings confirm the trajectory. The next chapter needs a costed plan.",
         C2X + 10, COL_Y + 48, COL_W - 20, 18, size=13, bold=True, color=C_WHITE)
obs2 = ("Rank 46 to 37 across three Smart50 cycles and a 31 percent average "
        "growth rate are proof points, not a plan. The stated 50 to 60 percent "
        "CAGR target requires a concrete answer to three questions: what headcount "
        "is needed to hit that number, how much cash does the ramp consume before "
        "new revenue arrives, and what is the RentRepair subscriber acquisition "
        "cost and payback period? Without those numbers, the ambition is real "
        "but the roadmap is not yet costed.")
txt_wrap(c, obs2, C2X + 10, COL_Y + 120, COL_W - 20, LH10, size=9, bold=False, color=C_OFF_WHITE)

# Col 03 GOLD
fill_rect(c, C3X, COL_Y, COL_W, COL_H, C_GOLD)
txt(c, "03", C3X + 10, COL_Y + 10, size=28, bold=True, color=C_BLACK)
txt_wrap(c, "Subscription economics are different from service economics",
         C3X + 10, COL_Y + 48, COL_W - 20, 18, size=13, bold=True, color=C_BLACK)
obs3 = ("RentRepair from $54 per month is a low entry price designed to drive "
        "volume adoption. The compliance job model (RentSafe) earns on each "
        "transaction. Mixing the two means upfront acquisition costs for "
        "subscribers sit against a slower cash recovery curve than the job "
        "flow generates today. Modelling unit economics for RentRepair "
        "separately, with cohort payback and churn sensitivity, is essential "
        "before scaling marketing spend on the subscription product.")
txt_wrap(c, obs3, C3X + 10, COL_Y + 120, COL_W - 20, LH10, size=9, bold=False, color=C_BLACK)

fill_rect(c, 0, PH - 22, PW, 22, C_BLACK)
txt(c, "Prepared by ProfitPulse  |  Profit-Pulse.com.au  |  Confidential",
    16, PH - 21, size=7, bold=False, color=C_OFF_WHITE)
txt(c, "2 / 3", 854, PH - 21, size=7, bold=False, color=C_OFF_WHITE,
    align="right", max_w=90)

c.showPage()

# ═══════════════════════════════════════════════════════════════
# PAGE 3 — CTA
# ═══════════════════════════════════════════════════════════════

fill_rect(c, 0, 0, PW, PH, C_WHITE)
fill_rect(c, 0, 0, 7, PH, C_AMBER_D)
fill_rect(c, 0, 0, PW, HDR_H, C_BLACK)

txt(c, "NEXT STEP", 16, 6, size=8, bold=True, color=C_OFF_WHITE)
txt(c, "PROFITPULSE", 614, 6, size=8, bold=True, color=C_TEAL,
    align="right", max_w=340)
txt(c, "Your growth is financeable. Let us show you how.",
    16, 25, size=20, bold=True, color=C_WHITE)
txt(c, "12 Jun 2026", 614, 25, size=11, bold=False, color=C_OFF_WHITE,
    align="right", max_w=340)

# Vertical divider
DIV_X = 482
fill_rect(c, DIV_X, 78, 2, 440, C_AMBER_D)

LX  = 16
LW2 = 460

# Service name
txt(c, "Strategic Growth Diagnostic", LX, 82, size=18, bold=True, color=C_BLACK)

# Price tile
fill_rect(c, LX, 105, 180, 36, C_BLACK)
txt(c, "$5,000  one off", LX + 8, 109, size=15, bold=True, color=C_AMBER_L)

desc = ("A structured diagnostic that maps your customer concentration, "
        "identifies your highest profitability growth levers and builds a "
        "costed three year growth plan calibrated to your actual cash and "
        "headcount constraints.")
txt_wrap(c, desc, LX, 152, LW2, 14, size=10, bold=False, color=C_BLACK)

deliv = ("Deliverable: a board ready financial model and growth roadmap with "
         "explicit assumptions, sensitivity ranges and a clear first ninety "
         "day action plan.")
txt_wrap(c, deliv, LX, 215, LW2, 14, size=9, bold=False, color=C_BLACK)

# Divider line
fill_rect(c, LX, 264, LW2, 1, C_AMBER_D)

txt(c, "Step one: answer a few quick questions",
    LX, 272, size=10, bold=True, color=C_BLACK)
txt(c, "profit-pulse.com.au/full-suite-of-products",
    LX, 290, size=10, bold=False, color=C_TEAL)

# Stripe CTA button
fill_rect(c, LX, 310, 274, 34, C_TEAL)
txt(c, "Purchase the suggested product now to get started",
    LX + 137, 316, size=10, bold=True, color=C_WHITE, align="center", max_w=0)

txt(c, "Or book a call:", LX, 360, size=9, bold=False, color=C_BLACK)
txt(c, "Schedule a call with Nitesh",
    LX + 108, 360, size=9, bold=True, color=C_TEAL)

# ── RIGHT column ──
RX  = DIV_X + 16
RW2 = 455

txt(c, "Nitesh Roopa  CA", RX, 82, size=17, bold=True, color=C_BLACK)
txt(c, "Managing Partner, ProfitPulse", RX, 104, size=11, bold=False, color=C_TEAL)

creds = [
    "16 years in corporate finance and commercial advisory",
    "52 completed transactions totalling AUD 10B gross deal value",
    "Largest single transaction: USD 1.3B",
    "Advisory work across Australia, South Africa, UK and UAE",
    "QIC (Queensland Investment Corporation) 2023 to 2025",
    "Nedbank Corporate and Investment Banking 2015 to 2022",
    "PwC Deals 2010 to 2014",
    "Qualified CA under SAICA",
]

for k, cred in enumerate(creds):
    cy = 125 + k * 31
    dot(c, RX + 3, cy + 5, r=3, color=C_TEAL)
    txt(c, cred, RX + 12, cy, size=9, bold=False, color=C_BLACK)

# Contact block
CB_Y = 393
fill_rect(c, RX, CB_Y, RW2, 1, C_BLACK)
txt(c, "Nitesh@Profit-Pulse.com.au   |   +61 411 876 267   |   Profit-Pulse.com.au",
    RX, CB_Y + 6, size=8, bold=False, color=C_BLACK)

# Bottom bar
fill_rect(c, 0, PH - 22, PW, 22, C_BLACK)
txt(c, "Prepared by ProfitPulse  |  Profit-Pulse.com.au  |  Confidential",
    16, PH - 21, size=7, bold=False, color=C_OFF_WHITE)
txt(c, "3 / 3", 854, PH - 21, size=7, bold=False, color=C_OFF_WHITE,
    align="right", max_w=90)

c.showPage()
c.save()
print(f"Saved: {OUT_FILE}")
