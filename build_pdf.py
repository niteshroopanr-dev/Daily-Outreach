"""
ProfitPulse Brief PDF Builder
Target: MuraConnect | Date: 28 Jul 2026
Mirrors build_brief.py exactly, same content, same coordinates, same colours.
Direct PDF generation (soffice PPTX to PDF conversion is unavailable in this
sandbox, so the PDF is built independently from the same layout data).
Page size 960 x 540 pt = 13.333in x 7.5in at 72dpi, matching the PPTX canvas.
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

BLACK     = HexColor("#000000")
TEAL      = HexColor("#01A296")
AMBER_B   = HexColor("#F8C806")
AMBER_D   = HexColor("#F6A102")
GOLD      = HexColor("#E3A712")
WHITE     = HexColor("#FFFFFF")
OFF_WHITE = HexColor("#E6E5DE")

PAGE_W = 960
PAGE_H = 540
IN = 72  # 1 inch in points

COMPANY = "MuraConnect"
DATE_STAMP = "28 Jul 2026"
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_URL = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

ALIGN = {"LEFT": TA_LEFT, "CENTER": TA_CENTER, "RIGHT": TA_RIGHT}


def _fontname(base, bold, italic):
    if base == "Helvetica":
        if bold and italic:
            return "Helvetica-BoldOblique"
        if bold:
            return "Helvetica-Bold"
        if italic:
            return "Helvetica-Oblique"
        return "Helvetica"
    if bold and italic:
        return "Times-BoldItalic"
    if bold:
        return "Times-Bold"
    if italic:
        return "Times-Italic"
    return "Times-Roman"


def rect(c, left_in, top_in, w_in, h_in, fill, stroke=None, stroke_w=1):
    x = left_in * IN
    y = PAGE_H - top_in * IN - h_in * IN
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(stroke_w)
        c.rect(x, y, w_in * IN, h_in * IN, fill=1, stroke=1)
    else:
        c.rect(x, y, w_in * IN, h_in * IN, fill=1, stroke=0)


def rrect(c, left_in, top_in, w_in, h_in, fill, stroke=None, stroke_w=1, radius=0.08):
    x = left_in * IN
    y = PAGE_H - top_in * IN - h_in * IN
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(stroke_w)
        c.roundRect(x, y, w_in * IN, h_in * IN, radius * IN, fill=1, stroke=1)
    else:
        c.roundRect(x, y, w_in * IN, h_in * IN, radius * IN, fill=1, stroke=0)


def text(c, s, left_in, top_in, w_in, h_in, size=11, colour=BLACK, bold=False,
         italic=False, align="LEFT", font="Helvetica", link=None, leading=None):
    style = ParagraphStyle("s", fontName=_fontname(font, bold, italic), fontSize=size,
                            leading=leading or size * 1.22, textColor=colour, alignment=ALIGN[align])
    p = Paragraph(s.replace("&", "&amp;"), style)
    w, h = p.wrap(w_in * IN, h_in * IN)
    x = left_in * IN
    y = PAGE_H - top_in * IN - h
    p.drawOn(c, x, y)
    if link:
        c.linkURL(link, (x, y, x + w_in * IN, y + h), relative=0, thickness=0)
    return h


def bullet_lines(c, lines, left_in, top_in, w_in, size=11, colour=BLACK, gap=0.37, font="Helvetica", bold=False):
    y = top_in
    for ln in lines:
        text(c, "•  " + ln, left_in, y, w_in, 0.32, size=size, colour=colour, font=font, bold=bold)
        y += gap


def footer(c):
    rect(c, 0.3, 7.0, 12.7, 1 / 72, BLACK)
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         0.3, 7.08, 9.5, 0.3, size=9, colour=BLACK)
    text(c, DATE_STAMP, 10.0, 7.08, 3.0, 0.3, size=9, colour=BLACK, align="RIGHT")


def header(c, eyebrow):
    rect(c, 0, 0, 13.333, 1.0, BLACK)
    text(c, eyebrow, 0.3, 0.32, 8.5, 0.4, size=13, colour=OFF_WHITE, bold=True)
    text(c, "PROFITPULSE", 9.5, 0.32, 3.5, 0.4, size=13, colour=TEAL, bold=True, align="RIGHT")


def stripe(c):
    rect(c, 0, 0, 0.1, 7.5, AMBER_B)


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MuraConnect_28Jul2026.pdf"
c = canvas.Canvas(out_path, pagesize=(PAGE_W, PAGE_H))

# ============================================================
# PAGE 1
# ============================================================
rect(c, 0, 0, 13.333, 7.5, WHITE)
header(c, "COMMERCIAL INTELLIGENCE BRIEF")
text(c, COMPANY, 0.3, 1.12, 10.5, 0.85, size=38, colour=TEAL, bold=True, font="Times")
text(c, "Indigenous owned recruitment and IT consultancy, Brisbane QLD",
     0.3, 1.95, 11.0, 0.4, size=13, colour=BLACK)

stat_cards = [
    ("$10.1M", "Annual revenue, FY2025", "Smart50 2025, Nov 2025"),
    ("90%", "Revenue growth year on year", "Smart50 2025 citation"),
    ("#6", "Smart50 2025 national rank", "SmartCompany, Nov 2025"),
    ("2017", "Year founded", "Company records"),
    ("51%", "Indigenous ownership stake", "Company website"),
]
card_y = 2.4
for i, (num, label, src) in enumerate(stat_cards):
    x = 0.3 + i * (2.3 + 0.15)
    rect(c, x, card_y, 2.3, 1.6, BLACK)
    rect(c, x, card_y, 2.3, 4 / 72, TEAL)
    text(c, num, x + 0.15, card_y + 0.16, 2.0, 0.55, size=26, colour=AMBER_B, bold=True, font="Times")
    text(c, label, x + 0.15, card_y + 0.78, 2.0, 0.5, size=11, colour=OFF_WHITE)
    text(c, src, x + 0.15, card_y + 1.30, 2.0, 0.25, size=8, colour=OFF_WHITE, italic=True)

text(c, "KEY COMMERCIAL SIGNALS", 0.3, 4.15, 6, 0.3, size=12, colour=TEAL, bold=True)
signals = [
    "6th of 50, 2025 Smart50 Awards: $10.1M revenue, 90% growth (SmartCompany)",
    "Founded 2017 by Phil Ahmat, Anthony Singh, Ash Kumar (Company records)",
    "Supply Nation accredited: 51% owned by CEO Phil Ahmat (MuraConnect site)",
    "Registered NSW Government supplier, buy.nsw panel (buy.nsw.gov.au)",
    "Team draws on more than 25 years serving major public, private clients (Company website)",
    "HQ at Fortitude Valley, inner Brisbane QLD (Company contact page)",
]
bullet_lines(c, signals, 0.3, 4.5, 12.4, size=11, colour=BLACK)

stripe(c)
footer(c)
c.showPage()

# ============================================================
# PAGE 2
# ============================================================
rect(c, 0, 0, 13.333, 7.5, WHITE)
header(c, "THE OPPORTUNITY")
text(c, "MuraConnect: three commercial observations from ProfitPulse",
     0.3, 1.1, 12.5, 0.4, size=16, colour=BLACK, bold=True, font="Times")

columns = [
    (TEAL, WHITE, OFF_WHITE, "01", "Growth is outrunning cash conversion",
     "MuraConnect grew revenue 90 percent to $10.1 million on a recruitment "
     "and IT project model where placed staff are usually paid weekly while "
     "enterprise and government clients settle on 30 to 60 day terms. Every "
     "new placement widens that gap before the cash lands. A Working Capital "
     "Unlock maps exactly where it sits."),
    (BLACK, AMBER_B, OFF_WHITE, "02", "Public sector clients concentrate the risk",
     "As a Supply Nation accredited, Indigenous owned provider on the buy.nsw "
     "panel, MuraConnect likely draws a large share of its $10.1 million "
     "revenue from a small number of public sector and enterprise "
     "relationships. Knowing which panels and clients drive margin, not just "
     "turnover, matters as contracts renew."),
    (GOLD, BLACK, BLACK, "03", "People cost is the business, so utilisation is the lever",
     "In a recruitment and IT consultancy, gross margin is set by billable "
     "utilisation and placement fee capture, not headcount growth. At 90 "
     "percent growth it is easy to add delivery capacity faster than "
     "utilisation is tracked. A monthly view of revenue per consultant keeps "
     "growth profitable, not just larger."),
]
col_xs = [0.3, 4.57, 8.84]
col_y = 1.65
col_w = 4.17
col_h = 3.9
for (fill, hcol, pcol, num, hdr, body), x in zip(columns, col_xs):
    rect(c, x, col_y, col_w, col_h, fill)
    text(c, num, x + 0.28, col_y + 0.22, col_w - 0.5, 0.7, size=32, colour=hcol, bold=True, font="Times")
    text(c, hdr, x + 0.28, col_y + 0.95, col_w - 0.55, 0.85, size=14, colour=hcol, bold=True)
    text(c, body, x + 0.28, col_y + 1.85, col_w - 0.55, 2.9, size=11, colour=pcol)

text(c, "These are observations offered in good faith. MuraConnect has built "
        "something genuinely impressive in a short time. The question is simply "
        "whether the financial architecture is keeping pace with the growth.",
     0.3, 5.85, 12.7, 0.6, size=11, colour=BLACK, italic=True)

stripe(c)
footer(c)
c.showPage()

# ============================================================
# PAGE 3
# ============================================================
rect(c, 0, 0, 13.333, 7.5, WHITE)
header(c, "THE RECOMMENDATION")

text(c, "Working Capital Unlock", 0.3, 1.15, 7.4, 0.55, size=24, colour=TEAL, bold=True, font="Times")
text(c, "$6,000 one off", 0.3, 1.72, 7.4, 0.4, size=17, colour=BLACK, bold=True)
text(c, "A four week project mapping cash trapped in debtors, work in progress "
        "and supplier terms, with a prioritised action list to release it.",
     0.3, 2.18, 7.4, 0.6, size=12, colour=BLACK)

rect(c, 0.3, 2.9, 7.4, 1.55, WHITE, stroke=TEAL, stroke_w=1.25)
text(c, "Step one, answer a few quick questions", 0.5, 3.0, 7.0, 0.35, size=13, colour=TEAL, bold=True)
text(c, "See the solutions matched to your size and industry.", 0.5, 3.38, 7.0, 0.32, size=11, colour=BLACK)
text(c, QUESTIONNAIRE_CLEAN, 0.5, 3.72, 7.0, 0.35, size=13, colour=TEAL, bold=True, link=QUESTIONNAIRE_URL)
rrect(c, 0.5, 4.12, 3.1, 0.24, TEAL)
c.linkURL(QUESTIONNAIRE_URL, (0.5 * IN, PAGE_H - 4.36 * IN, 3.6 * IN, PAGE_H - 4.12 * IN), relative=0, thickness=0)
text(c, "Find your fit in two minutes", 0.5, 4.13, 3.1, 0.22, size=10, colour=WHITE, bold=True, align="CENTER")

rrect(c, 0.3, 4.75, 4.9, 0.5, WHITE, stroke=AMBER_D, stroke_w=1.25)
c.linkURL(STRIPE_URL, (0.3 * IN, PAGE_H - 5.25 * IN, 5.2 * IN, PAGE_H - 4.75 * IN), relative=0, thickness=0)
text(c, "Purchase the suggested product now to get started", 0.3, 4.88, 4.9, 0.3, size=11, colour=AMBER_D,
     bold=True, align="CENTER", link=STRIPE_URL)
text(c, "$6,000 one off, ProfitPulse verified price", 0.3, 5.32, 6, 0.3, size=10, colour=BLACK)

text(c, "Prefer a conversation first?", 0.3, 5.85, 6, 0.32, size=12, colour=BLACK)
text(c, "Book a complimentary discovery call", 0.3, 6.2, 6, 0.35, size=12, colour=TEAL, bold=True, link=BOOKING_URL)

panel_x = 8.05
panel_w = 4.98
rect(c, panel_x, 1.15, panel_w, 5.75, BLACK)
text(c, "NITESH ROOPA", panel_x + 0.25, 1.35, panel_w - 0.5, 0.45, size=18, colour=AMBER_B, bold=True)
text(c, "CA, Managing Partner, ProfitPulse", panel_x + 0.25, 1.78, panel_w - 0.5, 0.35, size=13, colour=WHITE)
rect(c, panel_x + 0.25, 2.18, panel_w - 0.5, 1.5 / 72, TEAL)

bullet_lines(c, [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal USD 1.3 billion, Cahora Bassa",
    "AUD 10 billion Queensland infrastructure value",
], panel_x + 0.25, 2.35, panel_w - 0.5, size=11, colour=OFF_WHITE, gap=0.4)

text(c, "Fractional CFO for growth stage businesses", panel_x + 0.25, 3.85, panel_w - 0.5, 0.35,
     size=10.5, colour=OFF_WHITE, italic=True)

rect(c, panel_x + 0.25, 4.35, panel_w - 0.5, 1.5 / 72, TEAL)
text(c, "Profit-Pulse.com.au", panel_x + 0.25, 4.5, panel_w - 0.5, 0.3, size=11, colour=OFF_WHITE)
text(c, "Nitesh@Profit-Pulse.com.au", panel_x + 0.25, 4.85, panel_w - 0.5, 0.3, size=11, colour=TEAL)
text(c, "+61 411 876 267", panel_x + 0.25, 5.2, panel_w - 0.5, 0.3, size=11, colour=OFF_WHITE)
text(c, "linkedin.com/in/nitesh-roopa-77594163", panel_x + 0.25, 5.55, panel_w - 0.5, 0.3, size=11, colour=OFF_WHITE)

stripe(c)
footer(c)
c.showPage()

c.save()
print(f"PDF saved: {out_path}")
