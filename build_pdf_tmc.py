"""
ProfitPulse Brief PDF Builder v3.3 house style
Target: TMC Fine Jewellers | Date: 12 Aug 2026
Mirrors build_brief_tmc.py exactly, direct PDF via reportlab.
Page 960pt x 540pt = 13.333in x 7.5in at 72dpi.
Brand colours only, no other colours anywhere.
"""

from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

W = 960.0
H = 540.0

BLACK     = HexColor("#000000")
TEAL      = HexColor("#01A296")
AMBER_B   = HexColor("#F8C806")
AMBER_D   = HexColor("#F6A102")
GOLD      = HexColor("#E3A712")
WHITE     = HexColor("#FFFFFF")
OFF_WHITE = HexColor("#E6E5DE")

SERIF = "Times-Bold"
SERIF_R = "Times-Roman"
SANS = "Helvetica"
SANS_B = "Helvetica-Bold"
SANS_I = "Helvetica-Oblique"

DATE_STAMP = "12 Aug 2026"
COMPANY = "TMC Fine Jewellers"


def y(top_in):
    return H - top_in * 72.0


def rect(c, x_in, top_in, w_in, h_in, colour):
    c.setFillColor(colour)
    c.rect(x_in * 72.0, y(top_in) - h_in * 72.0, w_in * 72.0, h_in * 72.0, fill=1, stroke=0)


def text(c, s, x_in, top_in, size, font=SANS, colour=BLACK, align="left", box_w_in=None):
    c.setFont(font, size)
    c.setFillColor(colour)
    xpt = x_in * 72.0
    ypt = y(top_in) - size
    if align == "right" and box_w_in:
        c.drawRightString(xpt + box_w_in * 72.0, ypt, s)
    elif align == "center" and box_w_in:
        c.drawCentredString(xpt + (box_w_in * 72.0) / 2, ypt, s)
    else:
        c.drawString(xpt, ypt, s)


def wrap_text(s, font, size, max_w_pt):
    words = s.split(" ")
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if stringWidth(trial, font, size) <= max_w_pt:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def para(c, s, x_in, top_in, w_in, size, font=SANS, colour=BLACK, leading_mult=1.22, space_after=8):
    max_w_pt = w_in * 72.0
    lines = wrap_text(s, font, size, max_w_pt)
    cy = top_in
    for ln in lines:
        text(c, ln, x_in, cy, size, font=font, colour=colour)
        cy += (size * leading_mult) / 72.0
    return cy + space_after / 72.0


def multi_para(c, paras, x_in, top_in, w_in, size, font=SANS, colour=BLACK, leading_mult=1.2, space_after=8):
    cy = top_in
    for p in paras:
        cy = para(c, p, x_in, cy, w_in, size, font=font, colour=colour, leading_mult=leading_mult, space_after=space_after)
    return cy


def house_chrome(c, eyebrow):
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    rect(c, 0, 0, 13.333, 1.0, BLACK)
    text(c, eyebrow, 0.35, 0.62, 12, font=SANS_B, colour=OFF_WHITE)
    text(c, "PROFITPULSE", 9.4, 0.62, 12, font=SANS_B, colour=WHITE, align="right", box_w_in=3.6)
    rect(c, 0, 0, 0.1, 7.5, AMBER_D)
    c.setStrokeColor(TEAL)
    c.setLineWidth(0.75)
    c.line(0.35 * 72, y(7.02), (0.35 + 12.6) * 72, y(7.02))
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         0.35, 7.28, 8, font=SANS, colour=BLACK)
    text(c, DATE_STAMP, 10.5, 7.28, 8, font=SANS, colour=BLACK, align="right", box_w_in=2.5)


c = canvas.Canvas("/tmp/claude-0/-home-user-Daily-Outreach/ca95556c-e243-5dd8-a8af-ec1c034b7243/scratchpad/pptxwork/tmc_brief.pdf",
                   pagesize=(W, H))

# ── SLIDE 1 ──────────────────────────────────────────────────────────────
house_chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")
text(c, COMPANY, 0.35, 1.6, 40, font=SERIF, colour=BLACK)
text(c, "Lab grown diamond and moissanite bridal jewellery house, New Farm, Brisbane QLD",
     0.35, 2.3, 13, font=SANS, colour=BLACK)

cards = [
    ("$12.7M", "FY2025 revenue", "Smart50 2025, rank 5"),
    ("95%",    "Revenue growth, FY25", "Smart50 2025 citation"),
    ("35",     "Team members", "Smart50 2025 citation"),
    ("4",      "Global showroom cities", "TMC showroom pages"),
    ("2020",   "Year founded", "TMC Our Story page"),
]
card_w, card_h, gap, x0, top0 = 2.3, 1.6, 0.15, 0.35, 2.6
x = x0
for number, label, source in cards:
    rect(c, x, top0, card_w, card_h, BLACK)
    rect(c, x, top0, card_w, 4/72.0, TEAL)
    text(c, number, x + 0.18, top0 + 0.52, 26, font=SERIF, colour=AMBER_B)
    lbl_lines = wrap_text(label, SANS, 11, (card_w - 0.36) * 72)
    ly = top0 + 0.92
    for ln in lbl_lines[:2]:
        text(c, ln, x + 0.18, ly, 11, font=SANS, colour=OFF_WHITE)
        ly += 0.17
    text(c, source, x + 0.18, top0 + 1.36, 8, font=SANS, colour=OFF_WHITE)
    x += card_w + gap

text(c, "KEY COMMERCIAL SIGNALS", 0.35, 4.75, 12, font=SANS_B, colour=TEAL)
signals = [
    "Smart50 2025 ranks TMC fifth nationally with 95 percent three year growth to $12.7M, per SmartCompany.",
    "Fourth global showroom opened in Armadale, Melbourne, 6 December 2025, after Brisbane, Sydney and a Mayfair, London pop up in June 2025, per Inside Retail.",
    "BFCM week 2025 was targeted to deliver over $3.75 million, against $2.6 million and 1,200 customers in the 2024 event, per SmartCompany.",
    "Brand renamed from The Moissanite Company to TMC Fine Jewellers as it broadened into lab grown diamonds, per Ragtrader.",
    "Founded 2020 by Makayla and Tom Donovan on a $10,000 initial investment, per Onya Magazine and company published material.",
]
multi_para(c, signals, 0.35, 5.18, 12.6, 11.5, font=SANS, colour=BLACK, leading_mult=1.18, space_after=9)

c.showPage()

# ── SLIDE 2 ──────────────────────────────────────────────────────────────
house_chrome(c, "THE OPPORTUNITY")
text(c, f"{COMPANY}: three commercial observations from ProfitPulse", 0.35, 1.42, 16, font=SERIF, colour=BLACK)

col_top, col_h, col_w, gap2, x0col = 1.65, 4.85, 4.24, 0.11, 0.35
col_x = [x0col, x0col + col_w + gap2, x0col + 2 * (col_w + gap2)]
fills = [TEAL, BLACK, GOLD]
text_colours = [BLACK, WHITE, BLACK]
sub_colours = [BLACK, TEAL, BLACK]

observations = [
    ("01", "Four showrooms, one working capital question",
     "Brisbane, Sydney, Melbourne and a London pop up now each hold moissanite and lab grown "
     "diamond stock ahead of sale. Every new showroom and every pre BFCM build locks up cash "
     "before it converts to revenue, and four sites multiply that exposure at once."),
    ("02", "One week now carries close to a third of the year",
     "The 2025 BFCM target of $3.75 million is planned six to eight months out, up from $2.6 "
     "million in 2024. That concentration is a strength when forecast well and a liquidity "
     "shock when the stock build outpaces the cash available to fund it."),
    ("03", "Growth has outrun the reporting cadence",
     "Four showrooms across three countries within about eighteen months is a rapid build for "
     "a five year old business. Matching that pace with a costed twelve month plan keeps the "
     "next site funded by design rather than by whichever quarter had cash to spare."),
]

for i, (idx, header, bodytext) in enumerate(observations):
    rect(c, col_x[i], col_top, col_w, col_h, fills[i])
    pad = 0.28
    text(c, idx, col_x[i] + pad, col_top + 0.62, 32, font=SERIF, colour=sub_colours[i])
    hy = para(c, header, col_x[i] + pad, col_top + 1.15, col_w - 2*pad, 15, font=SERIF, colour=text_colours[i],
              leading_mult=1.12, space_after=6)
    multi_para(c, [bodytext], col_x[i] + pad, col_top + 2.05, col_w - 2*pad, 11, font=SANS,
               colour=text_colours[i], leading_mult=1.28, space_after=0)

text(c, "These observations are offered in good faith. Makayla and Tom have built something rare",
     0.35, 6.62, 10.5, font=SANS_I, colour=BLACK)
text(c, "in five years. The question is simply whether the financial architecture keeps pace with the shopfronts.",
     0.35, 6.8, 10.5, font=SANS_I, colour=BLACK)

c.showPage()

# ── SLIDE 3 ──────────────────────────────────────────────────────────────
house_chrome(c, "THE RECOMMENDATION")
left_x, left_w = 0.35, 7.65
text(c, "Working Capital Unlock", left_x, 1.55, 27, font=SERIF, colour=BLACK)
text(c, "$7,500 one off", left_x, 2.02, 16, font=SANS_B, colour=TEAL)
multi_para(c, ["Maps cash trapped in showroom stock, workshop inventory and supplier terms across all "
               "four sites, with a prioritised plan to release it before the next major sale week."],
           left_x, 2.42, left_w, 11.5, font=SANS, colour=BLACK, leading_mult=1.2)

rect(c, left_x, 3.05, left_w, 1.55, OFF_WHITE)
c.setStrokeColor(TEAL)
c.setLineWidth(1)
c.rect(left_x * 72, y(3.05) - 1.55*72, left_w*72, 1.55*72, fill=0, stroke=1)
text(c, "Step one, answer a few quick questions", left_x + 0.25, 3.42, 13, font=SANS_B, colour=BLACK)
text(c, "See the solutions matched to your size and industry.", left_x + 0.25, 3.78, 11, font=SANS, colour=BLACK)
text(c, "profit-pulse.com.au/services/find-your-fit", left_x + 0.25, 4.18, 13, font=SANS_B, colour=TEAL)

rect(c, left_x, 4.78, left_w, 0.7, AMBER_D)
text(c, "Purchase the suggested product now to get started", left_x, 5.18, 13, font=SANS_B, colour=BLACK,
     align="center", box_w_in=left_w)

text(c, "Prefer a conversation first?", left_x, 5.86, 11, font=SANS, colour=BLACK)
text(c, "Book a complimentary discovery call", left_x, 6.2, 12, font=SANS_B, colour=TEAL)

right_x, right_w = 8.25, 4.73
rect(c, right_x, 1.15, right_w, 5.5, BLACK)
rect(c, right_x, 1.15, right_w, 4/72.0, TEAL)
pad = 0.3
text(c, "Nitesh Roopa", right_x + pad, 1.82, 19, font=SERIF, colour=AMBER_B)
text(c, "CA, Managing Partner, ProfitPulse", right_x + pad, 2.18, 12, font=SANS, colour=WHITE)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
multi_para(c, cred, right_x + pad, 2.68, right_w - 2*pad, 11, font=SANS, colour=OFF_WHITE, leading_mult=1.15, space_after=9)

c.setStrokeColor(TEAL)
c.setLineWidth(1)
c.line((right_x+pad)*72, y(4.25), (right_x+pad+right_w-2*pad)*72, y(4.25))

contact = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
multi_para(c, contact, right_x + pad, 4.68, right_w - 2*pad, 11, font=SANS, colour=OFF_WHITE, leading_mult=1.25, space_after=7)

c.showPage()
c.save()
print("PDF saved.")
