"""
ProfitPulse Brief PDF Builder
Target: Sniip | Date: 19 Aug 2026
Direct PDF generation, 3 slides, brand colours only, zero dashes.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
Mirrors build_brief.py layout exactly.
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor
import textwrap

W = 960
H = 540

BLACK     = HexColor("#000000")
TEAL      = HexColor("#01A296")
AMBER_B   = HexColor("#F8C806")
AMBER_D   = HexColor("#F6A102")
GOLD      = HexColor("#E3A712")
WHITE     = HexColor("#FFFFFF")
OFF_WHITE = HexColor("#E6E5DE")

DATE_STR = "19 Aug 2026"
IN = 72  # 1 inch in points


def y(top_in):
    """Convert a top-origin inch coordinate to ReportLab bottom-origin points."""
    return H - top_in * IN


def rect(c, x_in, top_in, w_in, h_in, colour, stroke=None, line_w=1):
    x = x_in * IN
    w = w_in * IN
    h = h_in * IN
    top = y(top_in)
    c.setFillColor(colour)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(line_w)
        c.rect(x, top - h, w, h, fill=1, stroke=1)
    else:
        c.rect(x, top - h, w, h, fill=1, stroke=0)


def text(c, s, x_in, top_in, size, colour, font="Helvetica-Bold", align="left", max_w_in=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    x = x_in * IN
    baseline = y(top_in) - size * 0.85
    if align in ("right", "center") and max_w_in:
        tw = c.stringWidth(s, font, size)
        if align == "right":
            x = x_in * IN + max_w_in * IN - tw
        else:
            x = x_in * IN + (max_w_in * IN - tw) / 2
    c.drawString(x, baseline, s)


def wrapped(c, s, x_in, top_in, max_w_in, size, colour, font="Helvetica",
            leading=None, max_chars=None):
    leading = leading or size * 1.32
    if max_chars is None:
        max_chars = int(max_w_in * IN / (size * 0.52))
    lines = textwrap.wrap(s, max_chars)
    cy = top_in
    for ln in lines:
        text(c, ln, x_in, cy, size, colour, font=font)
        cy += leading / IN
    return cy


def bullets(c, items, x_in, top_in, max_w_in, size, colour, font="Helvetica",
            leading_in=0.24, max_chars=None):
    cy = top_in
    for it in items:
        cy = wrapped(c, "•  " + it, x_in, cy, max_w_in, size, colour, font=font,
                     max_chars=max_chars)
        cy += leading_in - (size * 1.32) / IN
    return cy


def chrome(c, eyebrow, header_right="PROFITPULSE"):
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    rect(c, 0, 0, 0.1, 7.5, AMBER_D)
    rect(c, 0, 0, 13.333, 1.0, BLACK)
    text(c, eyebrow, 0.35, 0.55, 12, OFF_WHITE, font="Helvetica-Bold")
    text(c, header_right, 9.0, 0.55, 13, TEAL, font="Helvetica-Bold",
         align="right", max_w_in=4.0)
    c.setStrokeColor(TEAL)
    c.setLineWidth(0.75)
    ly = y(7.03)
    c.line(0.35 * IN, ly, (0.35 + 12.6) * IN, ly)
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
            "Profit-Pulse.com.au", 0.35, 7.32, 8, BLACK, font="Helvetica")
    text(c, DATE_STR, 10.5, 7.32, 8, BLACK, font="Helvetica", align="right", max_w_in=2.45)


def stat_card(c, x_in, top_in, w_in, h_in, number, label_lines, source):
    rect(c, x_in, top_in, w_in, h_in, BLACK)
    rect(c, x_in, top_in, w_in, 0.055, TEAL)
    text(c, number, x_in + 0.15, top_in + 0.15, 26, AMBER_B, font="Helvetica-Bold")
    cy = top_in + 0.68
    for ln in label_lines:
        text(c, ln, x_in + 0.15, cy, 10.5, OFF_WHITE, font="Helvetica")
        cy += 0.19
    text(c, source, x_in + 0.15, top_in + h_in - 0.18, 7, OFF_WHITE, font="Helvetica-Oblique")


c = canvas.Canvas("/home/user/Daily-Outreach/Out-reach efforts/Brief_Sniip_19Aug2026.pdf",
                   pagesize=(W, H))

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ═════════════════════════════════════════════════════════════════════════
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")
text(c, "Sniip", 0.35, 1.25, 42, BLACK, font="Helvetica-Bold")
text(c, "Consumer and business bill payment platform, Brisbane, Queensland",
     0.35, 2.05, 13, BLACK, font="Helvetica")

cards = [
    ("2014", ["Founded in", "Brisbane"], "Sniip company site"),
    ("200K+", ["Business and personal", "customers"], "Sniip company site"),
    ("$500M+", ["Processed in", "bill payments"], "Sniip company site"),
    ("2025", ["AFR Fast 100", "debut year"], "AFR Fast 100 2025 list"),
    ("5", ["Major payment partners", "added"], "Company & partner news"),
]
cw, gap, x0, ctop, ch = 2.3, 0.15, 0.35, 2.45, 1.6
x = x0
for num, lbl, src in cards:
    stat_card(c, x, ctop, cw, ch, num, lbl, src)
    x += cw + gap

text(c, "KEY COMMERCIAL SIGNALS", 0.35, 4.28, 12, TEAL, font="Helvetica-Bold")
signals = [
    "Debuted on the AFR Fast 100 2025, confirming FY25 revenue above five million dollars.",
    "Won Best Innovation in Payments at the Fintech Australia Finnies Awards, June 2025.",
    "Named a finalist again in the same category at the 2026 Finnies Awards.",
    "Added BPAY, American Express, Qantas Frequent Flyer, Virgin and WEX Motorpass as partners.",
    "LinkedIn lists a team of 11 to 50 people running the platform from Brisbane.",
    "Founded in 2014 by Damien Vasta, who remains Founder and Chief Executive Officer.",
]
bullets(c, signals, 0.35, 4.62, 12.6, 11.5, BLACK, max_chars=105)
c.showPage()

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ═════════════════════════════════════════════════════════════════════════
chrome(c, "THE OPPORTUNITY", "SNIIP")
text(c, "Three commercial observations from ProfitPulse", 0.35, 1.45, 15, BLACK,
     font="Helvetica-Bold")

col_w, col_gap, col_top, col_h = 4.0, 0.13, 1.75, 4.55
cols = [
    (0.35, TEAL, WHITE, BLACK,
     "01", "Five partners, one finance team",
     "In roughly a year Sniip added BPAY, American Express, Qantas Frequent "
     "Flyer, Virgin Australia Business Flyer and WEX Motorpass. Each partner "
     "carries its own settlement cycle and margin profile. Layering five "
     "economic models onto one finance function in twelve months tests "
     "reporting fast."),
    (0.35 + col_w + col_gap, BLACK, WHITE, AMBER_B,
     "02", "A Fast 100 debut raises the bar",
     "Sniip's AFR Fast 100 2025 debut confirms revenue growth strong enough "
     "to clear the five million dollar threshold on a three year view. Lists "
     "like this draw investor and partner attention. Converting that "
     "attention well means board grade numbers are ready on request, not "
     "assembled after the fact."),
    (0.35 + 2 * (col_w + col_gap), GOLD, BLACK, BLACK,
     "03", "A rewards liability needs a steady hand",
     "Every bill paid can earn points across Qantas, Virgin and Amex "
     "programs. That is a growing liability sitting behind a fast growing "
     "transaction base. A monthly Fractional CFO Partnership keeps cash "
     "position, partner economics and the management pack as sharp as the "
     "product roadmap."),
]
for cx, fill, body_col, num_col, num, head, para in cols:
    rect(c, cx, col_top, col_w, col_h, fill)
    text(c, num, cx + 0.25, col_top + 0.6, 30, num_col, font="Helvetica-Bold")
    wrapped(c, head, cx + 0.25, col_top + 1.1, col_w - 0.5, 14.5, body_col,
            font="Helvetica-Bold", leading=18, max_chars=24)
    wrapped(c, para, cx + 0.25, col_top + 1.85, col_w - 0.5, 10.5, body_col,
            font="Helvetica", leading=14.5, max_chars=42)

wrapped(c, "These observations are offered in good faith. Sniip has built something "
           "genuinely impressive. The question is simply whether the financial "
           "architecture keeps pace with the partnership roadmap.",
        0.35, 6.55, 12.6, 10.5, BLACK, font="Helvetica-Oblique", max_chars=140)
c.showPage()

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ═════════════════════════════════════════════════════════════════════════
chrome(c, "THE RECOMMENDATION", "SNIIP")
lx, lw = 0.35, 7.3
text(c, "Fractional CFO Partnership", lx, 1.6, 24, BLACK, font="Helvetica-Bold")
text(c, "$4,950 per month", lx, 2.1, 16, TEAL, font="Helvetica-Bold")
text(c, "ProfitPulse verified price. The questionnaire confirms the exact fit for Sniip.",
     lx, 2.42, 9.5, BLACK, font="Helvetica-Oblique")
wrapped(c, "A senior financial partner at the table each month: management pack, "
           "quarterly board grade review, and support as new partnerships add "
           "complexity.", lx, 2.78, lw, 11.5, BLACK, font="Helvetica", max_chars=78)

rect(c, lx, 3.4, lw, 1.35, OFF_WHITE, stroke=TEAL, line_w=1)
text(c, "Step one, answer a few quick questions", lx + 0.2, 3.62, 12.5, BLACK,
     font="Helvetica-Bold")
text(c, "See the solutions matched to your size and industry.", lx + 0.2, 3.95, 10.5,
     BLACK, font="Helvetica")
text(c, "profit-pulse.com.au/services/find-your-fit", lx + 0.2, 4.28, 12, TEAL,
     font="Helvetica-Bold")

rect(c, lx, 5.0, lw, 0.62, TEAL)
text(c, "Purchase the suggested product now to get started", lx, 5.38, 13, WHITE,
     font="Helvetica-Bold", align="center", max_w_in=lw)

text(c, "Prefer a conversation first?", lx, 5.95, 10.5, BLACK, font="Helvetica")
text(c, "Book a complimentary discovery call", lx, 6.25, 12, AMBER_D, font="Helvetica-Bold")

rx, rw = 7.95, 5.03
rect(c, rx, 1.2, rw, 5.1, BLACK)
rect(c, rx, 1.2, rw, 0.055, TEAL)
text(c, "Nitesh Roopa", rx + 0.25, 1.65, 17, AMBER_B, font="Helvetica-Bold")
text(c, "CA, Managing Partner, ProfitPulse", rx + 0.25, 1.98, 11, OFF_WHITE, font="Helvetica")

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "AUD 10 billion GRBT project value in Queensland",
]
cred_end = bullets(c, cred, rx + 0.25, 2.45, rw - 0.5, 10.5, OFF_WHITE,
                    max_chars=34, leading_in=0.32)

div_top = cred_end + 0.14
c.setStrokeColor(TEAL)
c.setLineWidth(1)
ly = y(div_top)
c.line((rx + 0.25) * IN, ly, (rx + rw - 0.25) * IN, ly)

contact = [
    ("Profit-Pulse.com.au", OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au", TEAL),
    ("+61 411 876 267", OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", OFF_WHITE),
]
cy = div_top + 0.28
for line, col in contact:
    text(c, line, rx + 0.25, cy, 11, col, font="Helvetica")
    cy += 0.28

c.showPage()
c.save()
print("PDF saved: /home/user/Daily-Outreach/Out-reach efforts/Brief_Sniip_19Aug2026.pdf")
