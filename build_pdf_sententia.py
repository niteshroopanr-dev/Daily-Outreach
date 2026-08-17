"""
ProfitPulse Brief PDF Builder
Target: Sententia Consulting | Date: 18 Aug 2026
Direct PDF generation, matches the PPTX house style exactly, brand colours only.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

W = 960.0
H = 540.0
IN = 72.0  # points per inch

C_BLACK   = HexColor("#000000")
C_TEAL    = HexColor("#01A296")
C_AMBER_B = HexColor("#F8C806")
C_AMBER_D = HexColor("#F6A102")
C_GOLD    = HexColor("#E3A712")
C_WHITE   = HexColor("#FFFFFF")
C_OFFW    = HexColor("#E6E5DE")

SERIF = "Times-Bold"
SANS  = "Helvetica"
SANS_B = "Helvetica-Bold"
SANS_I = "Helvetica-Oblique"

DATE_STAMP = "18 Aug 2026"
COMPANY = "Sententia Consulting"


def y0(top, h=0):
    """Convert a top-origin y (and optional height) to ReportLab bottom origin."""
    return H - top - h


def rect(c, x, top, w, h, fill_colour, stroke_colour=None, stroke_w=1):
    c.setFillColor(fill_colour)
    if stroke_colour:
        c.setStrokeColor(stroke_colour)
        c.setLineWidth(stroke_w)
        c.rect(x, y0(top, h), w, h, fill=1, stroke=1)
    else:
        c.rect(x, y0(top, h), w, h, fill=1, stroke=0)


def hline(c, x, top, w, colour, thickness=1.2):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, y0(top), x + w, y0(top))


def text(c, s, x, top, size, colour, font=SANS, align="left", max_w=None, link=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = y0(top + size * 0.9)
    tx = x
    if align in ("right", "center") and max_w:
        tw = c.stringWidth(s, font, size)
        if align == "right":
            tx = x + max_w - tw
        else:
            tx = x + (max_w - tw) / 2
    c.drawString(tx, baseline, s)
    if link:
        tw = c.stringWidth(s, font, size)
        c.linkURL(link, (tx, baseline - 2, tx + tw, baseline + size), relative=0, thickness=0)


def wrap_lines(c, s, font, size, max_w):
    words = s.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if c.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def text_wrapped(c, s, x, top, max_w, size, colour, font=SANS, leading=None):
    if leading is None:
        leading = size * 1.5
    lines = wrap_lines(c, s, font, size, max_w)
    c.setFillColor(colour)
    c.setFont(font, size)
    ty = top
    for line in lines:
        c.drawString(x, y0(ty + size * 0.9), line)
        ty += leading
    return ty


def bulleted(c, items, x, top, max_w, size, colour, font=SANS, leading=None, gap=6):
    if leading is None:
        leading = size * 1.4
    ty = top
    for item in items:
        end_y = text_wrapped(c, "•  " + item, x, ty, max_w, size, colour, font, leading)
        ty = end_y + gap
    return ty


def chrome(c, eyebrow, header_right="PROFITPULSE"):
    rect(c, 0, 0, W, H, C_WHITE)
    rect(c, 0, 0, 0.1 * IN, H, C_AMBER_D)
    rect(c, 0, 0, W, 1.0 * IN, C_BLACK)
    text(c, eyebrow, 0.35 * IN, 0.34 * IN, 12, C_OFFW, font=SANS_B)
    text(c, header_right, 8.5 * IN, 0.34 * IN, 13, C_AMBER_B, font=SANS_B,
         align="right", max_w=4.5 * IN)
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
         0.35 * IN, 7.1 * IN, 8, C_BLACK, font=SANS)
    text(c, DATE_STAMP, 10.0 * IN, 7.1 * IN, 8, C_BLACK, font=SANS,
         align="right", max_w=2.9 * IN)


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_SententiaConsulting_18Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Sententia Consulting | ProfitPulse Brief | 18 Aug 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")

# ════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ════════════════════════════════════════════════════════════════════════
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")

text(c, COMPANY, 0.35 * IN, 1.2 * IN, 40, C_BLACK, font=SERIF)
text(c, "Government risk and financial advisory firm, Barton, Canberra ACT",
     0.35 * IN, 2.0 * IN, 14, C_TEAL, font=SANS)

card_y = 2.4 * IN
card_h = 1.6 * IN
card_w = 2.3 * IN
card_gap = 0.15 * IN
x = 0.35 * IN

stat_cards = [
    ("$10M+", ["Revenue surpassed,", "FY2025"], "AFR Fast 100 2025"),
    ("60.2%", ["Three year revenue", "CAGR"], "AFR Fast 100, rank 45"),
    ("2020", ["Founded as a two", "person firm"], "Consultancy.com.au"),
    ("$45.3M", ["Lifetime federal", "contract value"], "Pollywatch, FY20 to FY26"),
    ("174", ["Federal contract", "notices won"], "Pollywatch register"),
]

for number, label_lines, source in stat_cards:
    rect(c, x, card_y, card_w, card_h, C_BLACK)
    rect(c, x, card_y, card_w, 0.06 * IN, C_TEAL)
    text(c, number, x + 0.15 * IN, card_y + 0.2 * IN, 26, C_AMBER_B, font=SERIF)
    ly = card_y + 0.72 * IN
    for line in label_lines:
        text(c, line, x + 0.15 * IN, ly, 11, C_OFFW, font=SANS)
        ly += 0.19 * IN
    text(c, source, x + 0.15 * IN, card_y + 1.36 * IN, 8, C_OFFW, font=SANS_I)
    x += card_w + card_gap

text(c, "KEY COMMERCIAL SIGNALS", 0.35 * IN, 4.3 * IN, 12, C_TEAL, font=SANS_B)

signals = [
    "Melbourne office opened Aug 2024, led by new state MD Tom Fazio (Consultancy.com.au)",
    "Brioni Bale promoted to Managing Director, four senior hires added in 2025 (Consultancy.com.au)",
    "Two agencies are over half of the $45.3M lifetime federal contract book (Pollywatch)",
    "FY2025 26 alone brought $7.15M in new Commonwealth contract awards (Pollywatch)",
    "Grew from 2 founders in 2020 to a near 20 strong team by 2023 (Consultancy.com.au)",
]
bulleted(c, signals, 0.35 * IN, 4.65 * IN, 12.3 * IN, 12, C_BLACK, font=SANS, gap=9)

c.showPage()

# ════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ════════════════════════════════════════════════════════════════════════
chrome(c, "THE OPPORTUNITY")

text(c, "Sententia Consulting: three commercial observations from ProfitPulse",
     0.35 * IN, 1.2 * IN, 18, C_BLACK, font=SERIF)

col_y = 1.75 * IN
col_h = 4.5 * IN
col_w = 4.1667 * IN
col_gap = 0.1 * IN
cx = 0.3 * IN

columns = [
    {
        "fill": C_TEAL, "ink": C_BLACK, "index": "01",
        "header": "Growth is outrunning infrastructure",
        "body": ("A 60.2 percent three year revenue CAGR took Sententia from two "
                 "founders to a near 20 strong team and a second state office in "
                 "under five years, per the 2025 AFR Fast 100 and Consultancy.com.au. "
                 "At this pace, capacity planning and capital allocation typically lag "
                 "delivery. A Strategic Growth Diagnostic maps revenue, capacity and "
                 "margin headroom into a costed 12 month plan before that gap widens."),
    },
    {
        "fill": C_BLACK, "ink": C_WHITE, "index": "02",
        "header": "Two agencies carry over half the book",
        "body": ("Public federal contract records show Services Australia and the "
                 "Department of Health and Aged Care together account for more than "
                 "half of Sententia's $45.3 million lifetime Commonwealth contract "
                 "value, per Pollywatch. A strong panel position, and a concentration "
                 "worth watching as a standing item, the kind of visibility a "
                 "Fractional CFO Partnership keeps in view every month."),
    },
    {
        "fill": C_GOLD, "ink": C_BLACK, "index": "03",
        "header": "Two offices means two engines to plan",
        "body": ("Canberra and Melbourne each carry their own delivery leadership "
                 "and their own billable mix. As headcount keeps growing across both "
                 "sites, per Consultancy.com.au reporting on recent hires, seeing "
                 "utilisation and revenue per consultant clearly across two offices "
                 "gets harder without dedicated tracking. A Workforce Capacity and "
                 "Utilisation Review builds that view."),
    },
]

for col in columns:
    rect(c, cx, col_y, col_w, col_h, col["fill"])
    text(c, col["index"], cx + 0.25 * IN, col_y + 0.2 * IN, 42, col["ink"], font=SERIF)
    text(c, col["header"], cx + 0.25 * IN, col_y + 1.1 * IN, 15, col["ink"], font=SANS_B)
    text_wrapped(c, col["body"], cx + 0.25 * IN, col_y + 1.75 * IN, col_w - 0.5 * IN,
                 10.5, col["ink"], font=SANS, leading=15.5)
    cx += col_w + col_gap

text(c, "These are observations offered in good faith. Sententia has built something",
     0.3 * IN, 6.4 * IN, 11, C_BLACK, font=SANS_I, align="center", max_w=12.7 * IN)
text(c, "impressive. The question is simply whether the financial architecture matches the ambition.",
     0.3 * IN, 6.58 * IN, 11, C_BLACK, font=SANS_I, align="center", max_w=12.7 * IN)

c.showPage()

# ════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ════════════════════════════════════════════════════════════════════════
chrome(c, "THE RECOMMENDATION")

left_x = 0.35 * IN
left_w = 7.4 * IN

text(c, "Strategic Growth Diagnostic", left_x, 1.18 * IN, 24, C_BLACK, font=SERIF)
text(c, "$5,000 one off", left_x, 1.7 * IN, 17, C_AMBER_D, font=SANS_B)
text_wrapped(c, ("Six weeks mapping revenue, capacity and margin headroom into a 12 month "
                 "growth plan with funding and capital allocation steps, sized to a two office model."),
             left_x, 2.14 * IN, left_w, 12, C_BLACK, font=SANS, leading=16)

rect(c, left_x, 2.95 * IN, left_w, 1.55 * IN, C_WHITE, stroke_colour=C_TEAL, stroke_w=1.5)
text(c, "STEP ONE, ANSWER A FEW QUICK QUESTIONS", left_x + 0.2 * IN, 3.1 * IN, 11, C_TEAL, font=SANS_B)
text(c, "See the solutions matched to your size and industry.", left_x + 0.2 * IN, 3.42 * IN, 11, C_BLACK, font=SANS)
text(c, "profit-pulse.com.au/services/find-your-fit", left_x + 0.2 * IN, 3.78 * IN, 14, C_TEAL, font=SANS_B,
     link="https://profit-pulse.com.au/services/find-your-fit/")

btn_x, btn_top, btn_w, btn_h = left_x, 4.7 * IN, 4.6 * IN, 0.55 * IN
rect(c, btn_x, btn_top, btn_w, btn_h, C_WHITE, stroke_colour=C_AMBER_D, stroke_w=1.5)
text(c, "Purchase the suggested product now to get started",
     btn_x, btn_top + 0.19 * IN, 11, C_AMBER_D, font=SANS_B, align="center", max_w=btn_w)
c.linkURL("https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
          (btn_x, y0(btn_top + btn_h), btn_x + btn_w, y0(btn_top)), relative=0, thickness=0)

text(c, "Prefer a conversation first?", left_x, 5.5 * IN, 11, C_BLACK, font=SANS)
text(c, "Book a complimentary discovery call", left_x, 5.84 * IN, 12, C_TEAL, font=SANS_B,
     link="https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

right_x = 8.05 * IN
right_w = 4.95 * IN
rect(c, right_x, 1.15 * IN, right_w, 5.55 * IN, C_BLACK)
rect(c, right_x, 1.15 * IN, right_w, 0.06 * IN, C_TEAL)

text(c, "Nitesh Roopa", right_x + 0.25 * IN, 1.4 * IN, 19, C_AMBER_B, font=SERIF)
text(c, "CA, Managing Partner, ProfitPulse", right_x + 0.25 * IN, 1.86 * IN, 12, C_WHITE, font=SANS)

cred_lines = [
    "16 years across 4 countries",
    "Over 52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
bulleted(c, cred_lines, right_x + 0.25 * IN, 2.4 * IN, right_w - 0.5 * IN, 11, C_OFFW, font=SANS, gap=8)

hline(c, right_x + 0.25 * IN, 4.2 * IN, right_w - 0.5 * IN, C_TEAL, 1.2)

contact_lines = [
    ("Profit-Pulse.com.au", C_OFFW),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL),
    ("+61 411 876 267", C_OFFW),
    ("linkedin.com/in/nitesh-roopa-77594163", C_TEAL),
]
cy = 4.4 * IN
for item_text, item_col in contact_lines:
    text(c, item_text, right_x + 0.25 * IN, cy, 11, item_col, font=SANS)
    cy += 0.24 * IN

c.save()
print(f"PDF saved: {out_path}")
