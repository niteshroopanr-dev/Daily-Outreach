"""
ProfitPulse Brief PDF Builder, v3.3 house style
Target: Taskforce Australia | Date: 07 Aug 2026
Direct PDF generation, 3 slides, brand colours only, zero dashes.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
Mirrors build_brief_taskforce.py exactly, element for element.
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import simpleSplit

IN = 72.0
W = 13.333 * IN
H = 7.5 * IN

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WHITE= HexColor("#E6E5DE")
C_GREY1    = HexColor("#222222")
C_GREY2    = HexColor("#333333")
C_GREY3    = HexColor("#444444")
C_GREY4    = HexColor("#666666")
C_GREY5    = HexColor("#9A9A9A")
C_GREY6    = HexColor("#CCCCCC")
C_GREY7    = HexColor("#F2F2F0")
C_GREY8    = HexColor("#0A0A0A")

SERIF = "Times-Bold"
SERIF_R = "Times-Roman"
SANS  = "Helvetica"
SANS_B = "Helvetica-Bold"
SANS_I = "Helvetica-Oblique"

NEXT_DAY = "07 Aug 2026"
COMPANY = "Taskforce Australia"


def y(top):
    """Convert top origin (points from top) to ReportLab bottom origin."""
    return H - top


def rect(c, left, top, w, h, colour, alpha=None):
    c.setFillColor(colour)
    c.setFillAlpha(alpha if alpha is not None else 1)
    c.rect(left, y(top + h), w, h, fill=1, stroke=0)
    c.setFillAlpha(1)


def text(c, s, left, top, size, font=SANS, colour=C_BLACK, align="left", box_w=None, alpha=None):
    c.setFont(font, size)
    c.setFillColor(colour)
    c.setFillAlpha(alpha if alpha is not None else 1)
    baseline = y(top) - size * 0.85
    if align == "right" and box_w:
        c.drawRightString(left + box_w, baseline, s)
    else:
        c.drawString(left, baseline, s)
    c.setFillAlpha(1)


def wrapped(c, s, left, top, box_w, size, font=SANS, colour=C_BLACK, leading=None, max_lines=None, alpha=None):
    leading = leading or size * 1.32
    lines = simpleSplit(s, font, size, box_w)
    if max_lines:
        lines = lines[:max_lines]
    c.setFillAlpha(alpha if alpha is not None else 1)
    c.setFont(font, size)
    c.setFillColor(colour)
    cy = top
    for ln in lines:
        baseline = y(cy) - size * 0.85
        c.drawString(left, baseline, ln)
        cy += leading
    c.setFillAlpha(1)
    return cy


def chrome(c, eyebrow, company_label):
    c.setFillColor(C_WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    # Left accent stripe
    rect(c, 0, 0, 0.1 * IN, H, C_AMBER_D)
    # Header band
    rect(c, 0, 0, W, 1.0 * IN, C_BLACK)
    text(c, eyebrow, 0.35 * IN, 0.42 * IN, 12, font=SANS_B, colour=C_OFF_WHITE)
    text(c, company_label, 9.3 * IN, 0.42 * IN, 12, font=SANS_B, colour=C_TEAL,
         align="right", box_w=3.8 * IN)
    # Footer
    rect(c, 0.35 * IN, 7.05 * IN, 12.6 * IN, 0.75, C_BLACK, alpha=0.15)
    text(c, "Prepared by Nitesh Roopa, CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
         0.35 * IN, 7.2 * IN, 8, font=SANS, colour=C_BLACK, alpha=0.55)
    text(c, NEXT_DAY, 10.5 * IN, 7.2 * IN, 8, font=SANS, colour=C_BLACK, alpha=0.55,
         align="right", box_w=2.45 * IN)


def stat_card(c, left, top, w, h, number, label_lines, source):
    rect(c, left, top, w, h, C_BLACK)
    rect(c, left, top, w, 4, C_TEAL)
    text(c, number, left + 0.15 * IN, top + 0.16 * IN + 27 * 0.85, 27, font=SERIF, colour=C_AMBER_B)
    ly = top + 0.68 * IN
    for line in label_lines:
        text(c, line, left + 0.15 * IN, ly + 10.5 * 0.85, 10.5, font=SANS, colour=C_OFF_WHITE)
        ly += 10.5 * 1.35
    text(c, source, left + 0.15 * IN, top + h - 0.28 * IN + 7.5 * 0.85, 7.5, font=SANS_I, colour=C_OFF_WHITE, alpha=0.55)


c = canvas.Canvas("/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_07Aug2026.pdf",
                   pagesize=(W, H))

# ══════════════════════════════════════════════════════════════════
# SLIDE 1
# ══════════════════════════════════════════════════════════════════
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")
text(c, COMPANY, 0.35 * IN, 1.5 * IN, 40, font=SERIF, colour=C_BLACK)
text(c, "Property compliance and maintenance network, Burnley, Melbourne VIC",
     0.35 * IN, 2.05 * IN, 13, font=SANS, colour=C_BLACK, alpha=0.65)

cards = [
    ("$12.8M", ["Revenue, FY2025", "reported"], "SmartCompany Smart50 2025"),
    ("31%", ["Revenue growth,", "year on year"], "SmartCompany Smart50 2025"),
    ("19", ["Internal team", "members"], "SmartCompany Smart50 2025"),
    ("5,000", ["Tradespeople in", "national network"], "Taskforce, Real plus release"),
    ("500", ["Real estate agencies", "served nationally"], "PropertyMe integrator profile"),
    ("2014", ["Founded, Burnley,", "Melbourne VIC"], "SmartCompany Smart50 2025"),
]
n = len(cards)
card_w = 2.03 * IN
gap = 0.15 * IN
total_w = card_w * n + gap * (n - 1)
start_x = (W - total_w) / 2
card_top = 2.4 * IN
card_h = 1.6 * IN
for i, (num, lbl, src) in enumerate(cards):
    x = start_x + i * (card_w + gap)
    stat_card(c, x, card_top, card_w, card_h, num, lbl, src)

text(c, "KEY COMMERCIAL SIGNALS", 0.35 * IN, 4.5 * IN, 12, font=SANS_B, colour=C_TEAL)
signals = [
    "Revenue grew 31 percent year on year to $12.8 million, Smart50 2025 rank 37.",
    "New 2025 partnership with Real plus extends reach across property agencies nationally.",
    "Platform integrates with PropertyMe, MRI Property Tree and Console, serving 500 agencies.",
    "National network of 5,000 tradespeople delivers compliance, maintenance and warranty work.",
    "Named Victorian State Winner for Outstanding Growth, 2024 Telstra Best of Business Awards.",
    "Winner, Proptech Association Awards 2024, for the RentSafe and RentRepair platform.",
]
sy = 4.85 * IN
for sline in signals:
    text(c, "▪  " + sline, 0.35 * IN, sy, 11.5, font=SANS, colour=C_BLACK)
    sy += 11.5 * 1.55

c.showPage()

# ══════════════════════════════════════════════════════════════════
# SLIDE 2
# ══════════════════════════════════════════════════════════════════
chrome(c, "THE OPPORTUNITY", COMPANY.upper())
text(c, "Three commercial observations from ProfitPulse", 0.35 * IN, 1.4 * IN, 20, font=SERIF, colour=C_BLACK)

col_top = 1.7 * IN
col_h = 4.75 * IN
col_w = 4.03 * IN
gap2 = 0.12 * IN
col_fills = [C_TEAL, C_BLACK, C_GOLD]
col_text = [C_WHITE, C_OFF_WHITE, C_BLACK]
col_idx = [C_BLACK, C_TEAL, C_BLACK]

observations = [
    ("01", "Three lines, three margins",
     "Taskforce runs three lines: RentSafe compliance checks, RentRepair maintenance and manufacturer warranty servicing. Each carries a different cost base and margin profile. Past $12.8 million in revenue, the real question is whether every line is growing profit as fast as it is growing revenue."),
    ("02", "A lean team, a wide network",
     "Nineteen internal staff coordinate a national network of 5,000 tradespeople across compliance, maintenance and warranty work. That leverage is a real asset, and it also means margin now lives in how well each job type and service line is priced and tracked at scale."),
    ("03", "Three new doors just opened",
     "The 2025 partnership with Real plus, alongside integrations with PropertyMe, MRI Property Tree and Console, opens the platform to more of the 500 agencies it already serves. New distribution is only as valuable as the margin behind the services it sells."),
]

for i, (idx, head, body) in enumerate(observations):
    x = 0.35 * IN + i * (col_w + gap2)
    rect(c, x, col_top, col_w, col_h, col_fills[i])
    text(c, idx, x + 0.25 * IN, col_top + 0.75 * IN, 34, font=SERIF, colour=col_idx[i])
    wrapped(c, head, x + 0.25 * IN, col_top + 1.15 * IN, col_w - 0.5 * IN, 15, font=SERIF, colour=col_text[i], leading=18)
    wrapped(c, body, x + 0.25 * IN, col_top + 1.75 * IN, col_w - 0.5 * IN, 10.5, font=SANS, colour=col_text[i], leading=14.5)

wrapped(c, "These observations are offered in good faith. Taskforce has built something genuinely "
           "impressive since 2014. The question is simply whether the margin behind each service "
           "line is as clear as the growth number.",
        0.35 * IN, 6.65 * IN, 12.6 * IN, 10.5, font=SANS_I, colour=C_BLACK, alpha=0.65, leading=14)

c.showPage()

# ══════════════════════════════════════════════════════════════════
# SLIDE 3
# ══════════════════════════════════════════════════════════════════
chrome(c, "THE RECOMMENDATION", COMPANY.upper())

lx = 0.35 * IN
lw = 7.6 * IN

text(c, "Product and Service Line Profitability", lx, 1.55 * IN, 21, font=SERIF, colour=C_BLACK)
text(c, "$3,950 one off", lx, 1.98 * IN, 16, font=SANS_B, colour=C_TEAL)
wrapped(c,
        "A three week project ranking every product or service line by gross margin, "
        "contribution margin and operational drag. For Taskforce, that means a clear, "
        "sourced view of what RentSafe, RentRepair and manufacturer warranty work each "
        "contribute once network payouts, platform costs and service time are counted.",
        lx, 2.35 * IN, lw, 11, font=SANS, colour=C_BLACK, leading=15)

rect(c, lx, 3.25 * IN, lw, 1.25 * IN, C_OFF_WHITE)
rect(c, lx, 3.25 * IN, 0.06 * IN, 1.25 * IN, C_TEAL)
text(c, "Step one, answer a few quick questions", lx + 0.25 * IN, 3.55 * IN, 13, font=SANS_B, colour=C_BLACK)
text(c, "See the solutions matched to your size and industry.", lx + 0.25 * IN, 3.85 * IN, 11, font=SANS, colour=C_BLACK, alpha=0.75)
c.setFillColor(C_TEAL)
c.setFont(SANS_B, 12)
link_y = y(4.2 * IN) - 12 * 0.85
c.drawString(lx + 0.25 * IN, link_y, "profit-pulse.com.au/services/find-your-fit")
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          (lx + 0.25 * IN, link_y - 2, lx + 0.25 * IN + 260, link_y + 12), relative=0)

rect(c, lx, 4.7 * IN, lw, 0.55 * IN, C_AMBER_D)
c.setFillColor(C_BLACK)
c.setFont(SANS_B, 13)
cta_y = y(5.03 * IN) - 13 * 0.85
c.drawString(lx + 0.25 * IN, cta_y, "Purchase the suggested product now to get started")
c.linkURL("https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
          (lx, y(4.7 * IN + 0.55 * IN), lx + lw, y(4.7 * IN)), relative=0)

text(c, "Prefer a conversation first?", lx, 5.65 * IN, 11, font=SANS, colour=C_BLACK, alpha=0.75)
c.setFillColor(C_TEAL)
c.setFont(SANS_B, 11)
book_y = y(5.95 * IN) - 11 * 0.85
c.drawString(lx, book_y, "Book a complimentary discovery call")
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (lx, book_y - 2, lx + 210, book_y + 11), relative=0)

rx = 8.25 * IN
rw = 4.73 * IN
rect(c, rx, 1.15 * IN, rw, 5.6 * IN, C_BLACK)
rect(c, rx, 1.15 * IN, rw, 4, C_TEAL)

text(c, "Nitesh Roopa", rx + 0.28 * IN, 1.68 * IN, 18, font=SERIF, colour=C_AMBER_B)
text(c, "CA, Managing Partner, ProfitPulse", rx + 0.28 * IN, 2.05 * IN, 12, font=SANS, colour=C_WHITE)
rect(c, rx + 0.28 * IN, 2.28 * IN, rw - 0.56 * IN, 1, C_TEAL)

cred_lines = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
cy = 2.6 * IN
for line in cred_lines:
    text(c, line, rx + 0.28 * IN, cy, 11, font=SANS, colour=C_OFF_WHITE)
    cy += 11 * 1.7

rect(c, rx + 0.28 * IN, 4.25 * IN, rw - 0.56 * IN, 1, HexColor("#333333"))

contact_lines = [
    ("Profit-Pulse.com.au", C_OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL),
    ("+61 411 876 267", C_OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", C_OFF_WHITE),
]
cy = 4.6 * IN
for line, colour in contact_lines:
    sz = 11 if line != contact_lines[-1][0] else 10
    text(c, line, rx + 0.28 * IN, cy, sz, font=SANS, colour=colour)
    cy += 11 * 1.55

c.showPage()
c.save()
print("PDF saved")
