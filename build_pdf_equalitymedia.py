"""
ProfitPulse Brief PDF Builder, Version 3.3 house style
Target: Equality Media + Marketing | Date: 06 Aug 2026
Direct PDF generation mirroring the PPTX. Brand colours only. Zero dashes in copy.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

W = 960
H = 540

C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_MUTED     = C_BLACK
C_PANEL_BG  = C_OFF_WHITE
C_PANEL_DK  = C_BLACK

DATE_STAMP = "06 Aug 2026"
COMPANY = "Equality Media + Marketing"
COMPANY_SHORT = "EQUALITY MEDIA + MARKETING"


def rl_y(y_top):
    return H - y_top


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold", align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl_y(y_top + size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline, text)


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    if leading is None:
        leading = size * 1.35
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
    lines, current = [], ""
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
        c.drawString(x, rl_y(y + size), line)
        y += leading
    return y


def hline(c, x, y_top, w, colour, thickness=1.2):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def chrome(c, eyebrow, right_label):
    fill_rect(c, 0, 0, W, H, C_WHITE)
    fill_rect(c, 0, 0, W, 72, C_BLACK)
    txt(c, eyebrow, 26, 25, 13, C_OFF_WHITE, font="Helvetica-Bold")
    txt(c, right_label, 300, 25, 13, C_TEAL, font="Helvetica-Bold", align="right", max_width=634)
    txt(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
           "Profit-Pulse.com.au", 26, 505, 8, C_MUTED, font="Helvetica")
    txt(c, DATE_STAMP, 760, 505, 8, C_MUTED, font="Helvetica", align="right", max_width=175)
    fill_rect(c, 0, 0, 7, H, C_AMBER_B)


def stat_card(c, x, y_top, w, h, number, label_lines, source):
    fill_rect(c, x, y_top, w, h, C_BLACK)
    fill_rect(c, x, y_top, w, 3, C_TEAL)
    txt(c, number, x + 9, y_top + 11, 26, C_AMBER_B, font="Helvetica-Bold")
    ly = y_top + 52
    for line in label_lines:
        txt(c, line, x + 9, ly, 10, C_OFF_WHITE, font="Helvetica")
        ly += 14
    txt(c, source, x + 9, y_top + h - 20, 7, C_OFF_WHITE, font="Helvetica-Oblique")


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMediaMarketing_06Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Equality Media + Marketing, ProfitPulse Brief, 06 Aug 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Working Capital Unlock")

# ============================================================================
# SLIDE 1
# ============================================================================
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF", COMPANY_SHORT)
txt(c, COMPANY, 26, 88, 32, C_BLACK, font="Helvetica-Bold")
txt(c, "Independent full service advertising agency, Richmond, Melbourne VIC",
    26, 132, 12, C_BLACK, font="Helvetica-Oblique")

cards = [
    ("$13M", ["Annual revenue,", "rank 24 nationally"], "SmartCompany Smart50 2025"),
    ("49%", ["Three year revenue", "growth rate"], "SmartCompany Smart50 2025"),
    ("30", ["People on the", "team"], "Mumbrella, 2026"),
    ("#55", ["AFR Fast 100", "2025 national rank"], "AFR Fast 100 list, 2025"),
    ("2018", ["Year the agency", "was founded"], "Smart50 2024 and 2025"),
]
card_w, gap, card_top, card_h = 164, 11, 175, 116
for i, (num, label, src) in enumerate(cards):
    x = 26 + i * (card_w + gap)
    stat_card(c, x, card_top, card_w, card_h, num, label, src)

txt(c, "REVENUE REPORTED AT AWARD TIME", 26, 315, 11, C_TEAL, font="Helvetica-Bold")
chart_x, chart_y, chart_w, chart_h = 26, 345, 396, 120
base_y = chart_y + chart_h
bar_w = 90
max_val = 13.0
vals = [("Smart50 2024", 7.3), ("Smart50 2025", 13.0)]
bx = chart_x + 30
for label, val in vals:
    bh = (val / max_val) * (chart_h - 20)
    fill_rect(c, bx, base_y - bh, bar_w, bh, C_TEAL)
    txt(c, f"${val:.1f}M", bx, base_y - bh - 16, 12, C_BLACK, font="Helvetica-Bold",
        align="center", max_width=bar_w)
    txt(c, label, bx, base_y + 4, 9, C_BLACK, font="Helvetica",
        align="center", max_width=bar_w)
    bx += bar_w + 90
hline(c, chart_x, base_y, chart_w, C_TEAL, 1)
txt(c, "Source: SmartCompany Smart50 2024 and 2025 award citations",
    26, 480, 7, C_BLACK, font="Helvetica-Oblique")

txt(c, "KEY COMMERCIAL SIGNALS", 456, 315, 11, C_TEAL, font="Helvetica-Bold")
signals = [
    "Ranked 24th on the 2025 Smart50 list, up from 45th in 2024. (SmartCompany)",
    "Also placed 55th nationally on the AFR Fast 100 2025 growth list. (AFR Fast 100)",
    "Appointed ANZ media agency of record for skincare brand BIODERMA in 2025. (AdNews, B&T)",
    "Named 2026 AFR BOSS Best Place to Work winner, Media and Marketing category. (Mumbrella)",
    "Team has grown from 17 people in 2024 toward 30 by 2026. (SmartCompany, Mumbrella)",
    "Runs a four day, full pay 32 hour work week introduced in 2022. (Mumbrella, AdNews)",
]
sy = 340
for s in signals:
    ey = txt_wrapped(c, "•  " + s, 456, sy, 480, 11, C_BLACK, font="Helvetica", leading=14)
    sy = ey + 8

c.showPage()

# ============================================================================
# SLIDE 2
# ============================================================================
chrome(c, "THE OPPORTUNITY", COMPANY_SHORT)
txt(c, "Equality Media + Marketing: three commercial observations from ProfitPulse",
    26, 82, 13, C_TEAL, font="Helvetica-Bold")

col_top, col_h, col_w, col_gap = 112, 340, 290, 9
col_x = [26, 26 + col_w + col_gap, 26 + 2 * (col_w + col_gap)]
col_fills = [C_TEAL, C_BLACK, C_GOLD]
col_text = [C_BLACK, C_OFF_WHITE, C_BLACK]

observations = [
    ("01", "Growth is outrunning the finance function",
     "Revenue reported at Smart50 award time moved from 7.3 million dollars in 2024 to "
     "13 million in 2025, while the team has grown toward 30 people. That pace of hiring "
     "and client onboarding is a common point where cash timing, not profit, becomes the "
     "real constraint on how fast a business can safely take on new work."),
    ("02", "Media buying carries a working capital load",
     "Full service planning and buying, including the new ANZ mandate for BIODERMA, "
     "typically means paying media suppliers ahead of client settlement. As billings "
     "scale, the gap between paying media costs and collecting client revenue widens, "
     "and can quietly absorb cash the business assumes it still has on hand."),
    ("03", "A four day week raises the value of every hour",
     "Equality Time has delivered 85 percent staff retention, 14 percent growth in gross "
     "profit and a 23 percent rise in client work. Running full pay on a 32 hour week "
     "means each hour of capacity carries more weight, which makes a live view of where "
     "time turns into margin as important as the cash discipline needed to fund growth."),
]

for i, x in enumerate(col_x):
    fill_rect(c, x, col_top, col_w, col_h, col_fills[i])
    idx, header, para = observations[i]
    txt(c, idx, x + 16, col_top + 16, 26, col_text[i], font="Helvetica-Bold")
    txt_wrapped(c, header, x + 16, col_top + 60, col_w - 32, 14, col_text[i],
                font="Helvetica-Bold", leading=17)
    txt_wrapped(c, para, x + 16, col_top + 110, col_w - 32, 10.5, col_text[i],
                font="Helvetica", leading=14)

txt_wrapped(c, "These observations are offered in good faith. Equality Media + Marketing has "
               "built something genuinely impressive. The question is simply whether the "
               "financial architecture keeps pace with the growth.",
            26, 466, 900, 11, C_BLACK, font="Helvetica-Oblique", leading=15)

c.showPage()

# ============================================================================
# SLIDE 3
# ============================================================================
chrome(c, "THE RECOMMENDATION", COMPANY_SHORT)

lx, lw = 26, 560
txt(c, "Working Capital Unlock", lx, 88, 22, C_BLACK, font="Helvetica-Bold")
txt(c, "$4,450 one off, ProfitPulse verified price", lx, 122, 13, C_TEAL, font="Helvetica-Bold")
txt_wrapped(c, "Maps cash trapped in debtors, supplier payment timing and stock or work in "
               "progress across the business, then delivers a prioritised action list to "
               "release cash within weeks.",
            lx, 148, lw, 11.5, C_BLACK, font="Helvetica", leading=15)

fill_rect(c, lx, 226, lw, 108, C_PANEL_BG)
stroke_rect(c, lx, 226, lw, 108, C_TEAL, 1)
txt(c, "Step one, answer a few quick questions", lx + 14, 240, 13, C_BLACK, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry.", lx + 14, 264, 11,
    C_BLACK, font="Helvetica")
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 13)
c.drawString(lx + 14, rl_y(226 + 82), "profit-pulse.com.au/services/find-your-fit")
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          (lx + 14, rl_y(226 + 96), lx + 300, rl_y(226 + 78)), relative=0)

fill_rect(c, lx, 350, lw, 48, C_AMBER_D)
txt(c, "Purchase the suggested product now to get started", lx, 366, 13, C_BLACK,
    font="Helvetica-Bold", align="center", max_width=lw)
c.linkURL("https://buy.stripe.com/14AaEYb163Y4a9q4c73ks0y", (lx, rl_y(398), lx + lw, rl_y(350)), relative=0)

txt(c, "Prefer a conversation first?", lx, 418, 12, C_BLACK, font="Helvetica-Bold")
c.setFillColor(C_TEAL)
c.setFont("Helvetica", 12)
c.drawString(lx, rl_y(440), "Book a complimentary discovery call")
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (lx, rl_y(444), lx + 260, rl_y(426)), relative=0)

rx, rw = 610, 324
fill_rect(c, rx, 88, rw, 372, C_PANEL_DK)
fill_rect(c, rx, 88, rw, 3, C_TEAL)
txt(c, "Nitesh Roopa", rx + 18, 110, 19, C_AMBER_B, font="Helvetica-Bold")
txt(c, "CA, Managing Partner, ProfitPulse", rx + 18, 136, 12, C_WHITE, font="Helvetica")

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion,",
    "Cahora Bassa, Mozambique Government",
    "Total GRBT project value over AUD 10 billion",
]
cy = 172
for line in cred:
    txt(c, line, rx + 18, cy, 11, C_OFF_WHITE, font="Helvetica")
    cy += 18

hline(c, rx + 18, cy + 8, rw - 36, C_TEAL, 1.2)
cy += 26
contact = [
    ("Profit-Pulse.com.au", C_OFF_WHITE, "Helvetica"),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL, "Helvetica"),
    ("+61 411 876 267", C_OFF_WHITE, "Helvetica"),
    ("linkedin.com/in/nitesh-roopa-77594163", C_OFF_WHITE, "Helvetica"),
]
for text_, colour, font in contact:
    size = 12 if font == "Helvetica" and colour != C_OFF_WHITE else 10
    txt(c, text_, rx + 18, cy, size, colour, font=font)
    cy += 22

c.showPage()
c.save()
print(f"PDF saved: {out_path}")
