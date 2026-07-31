"""
ProfitPulse Brief PDF Builder
Target: Taskforce Australia | Date: 01 Aug 2026
Mirrors build_brief_taskforce.py geometry exactly (inches * 72 = points).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

IN = 72.0
W = 13.333 * IN
H = 7.5 * IN

C_BLACK = HexColor("#000000")
C_TEAL = HexColor("#01A296")
C_AMBER_B = HexColor("#F8C806")
C_AMBER_D = HexColor("#F6A102")
C_GOLD = HexColor("#E3A712")
C_WHITE = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_MID_GREY = HexColor("#666666")
C_DK_TXT = HexColor("#222222")
C_LINE_GREY = HexColor("#333333")
C_TEAL_TINT = HexColor("#F2FAF9")

DATE_STR = "01 Aug 2026"
COMPANY = "Taskforce Australia"


def y(top_in):
    return H - top_in * IN


def rect(c, l, t, w, h, colour):
    c.setFillColor(colour)
    c.rect(l * IN, y(t) - h * IN, w * IN, h * IN, fill=1, stroke=0)


def rect_stroke(c, l, t, w, h, fill_colour, stroke_colour, lw=1):
    c.setFillColor(fill_colour)
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(lw)
    c.rect(l * IN, y(t) - h * IN, w * IN, h * IN, fill=1, stroke=1)


def txt(c, s, l, t, size, colour, font="Helvetica", align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = y(t) - size * 0.8
    if align == "right" and max_w:
        tw = c.stringWidth(s, font, size)
        c.drawString((l + max_w) * IN - tw, baseline, s)
    elif align == "center" and max_w:
        tw = c.stringWidth(s, font, size)
        c.drawString(l * IN + (max_w * IN - tw) / 2, baseline, s)
    else:
        c.drawString(l * IN, baseline, s)


def txt_wrapped(c, s, l, t, max_w_in, size, colour, font="Helvetica", leading=None, bullet=False):
    if leading is None:
        leading = size * 1.3
    words = s.split()
    lines, cur = [], ""
    max_w_pt = max_w_in * IN
    prefix = "•  " if bullet else ""
    for wd in words:
        test = (cur + " " + wd).strip()
        if c.stringWidth(prefix + test, font, size) <= max_w_pt:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = wd
    if cur:
        lines.append(cur)
    cy = t
    for i, line in enumerate(lines):
        text_line = (prefix if i == 0 else ("   " if bullet else "")) + line
        txt(c, text_line, l, cy, size, colour, font=font)
        cy += leading / IN
    return cy


def hline(c, l, t, w, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    yy = y(t)
    c.line(l * IN, yy, (l + w) * IN, yy)


def header_band(c, eyebrow, right_label="PROFITPULSE"):
    rect(c, 0, 0, 0.1, 7.5, C_AMBER_D)
    rect(c, 0, 0, 13.333, 1.0, C_BLACK)
    txt(c, eyebrow, 0.35, 0.55, 12, C_OFF_WHITE, font="Helvetica-Bold")
    txt(c, right_label, 9.0, 0.55, 12, C_TEAL, font="Helvetica-Bold", align="right", max_w=3.98)


def footer(c):
    txt(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
        0.35, 7.28, 8, C_MID_GREY, font="Helvetica")
    txt(c, DATE_STR, 10.6, 7.28, 8, C_MID_GREY, font="Helvetica", align="right", max_w=2.38)


def stat_card(c, l, t, w, h, number, label_lines, source):
    rect(c, l, t, w, h, C_BLACK)
    rect(c, l, t, w, 4 / IN, C_TEAL)
    txt(c, number, l + 0.12, t + 0.42, 22, C_AMBER_B, font="Helvetica-Bold")
    ly = t + 0.78
    for line in label_lines:
        txt(c, line, l + 0.12, ly, 10.5, C_OFF_WHITE, font="Helvetica")
        ly += 0.19
    txt(c, source, l + 0.12, t + h - 0.14, 7.5, HexColor("#9a9a9a"), font="Helvetica-Oblique")


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle(f"{COMPANY} | ProfitPulse Brief | {DATE_STR}")
c.setAuthor("ProfitPulse")
c.setSubject("Product and Service Line Profitability")

# ════════════════════════════════════ SLIDE 1 ════════════════════════════════════
rect(c, 0, 0, 13.333, 7.5, C_WHITE)
header_band(c, "COMMERCIAL INTELLIGENCE BRIEF")

txt(c, COMPANY, 0.35, 1.55, 40, C_BLACK, font="Helvetica-Bold")
txt(c, "National rental property compliance and maintenance platform, Burnley VIC",
    0.35, 2.22, 12, C_DK_TXT, font="Helvetica-Oblique")

cards = [
    ("$12.8M", ["FY25 revenue"], "Smart50 2025 award, Nov 2025"),
    ("72%", ["Revenue growth,", "FY23 to FY25"], "Smart50 2023 and 2025"),
    ("19", ["People on", "the team"], "Smart50 2025 award"),
    ("2014", ["Founded,", "Burnley VIC"], "Company website"),
    ("140K+", ["RentSafe jobs", "completed"], "Taskforce Australia site"),
]
card_w, gap, start_x, card_top, card_h = 2.3, 0.15, 0.6, 2.4, 1.6
for i, (num, lbl, src) in enumerate(cards):
    l = start_x + i * (card_w + gap)
    stat_card(c, l, card_top, card_w, card_h, num, lbl, src)

txt(c, "REVENUE, THREE VERIFIED YEARS", 0.6, 4.55, 11, C_TEAL, font="Helvetica-Bold")
bars = [("FY23", 7.43), ("FY24", 9.7), ("FY25", 12.8)]
max_val = 12.8
bar_w = 0.9
bar_gap = 0.75
bx = 0.8
base_top = 6.55
area_h = 1.7
for label, val in bars:
    bh = area_h * (val / max_val)
    top_y = base_top - bh
    rect(c, bx, top_y, bar_w, bh, C_TEAL)
    txt(c, f"${val:.2f}M", bx - 0.05, top_y - 0.06, 11, C_BLACK, font="Helvetica-Bold")
    txt(c, label, bx, base_top + 0.24, 10, HexColor("#444444"), font="Helvetica", align="center", max_w=bar_w)
    bx += bar_w + bar_gap
txt(c, "Source: SmartCompany Smart50 award citations, 2023, 2024 and 2025",
    0.6, 6.95, 7.5, HexColor("#777777"), font="Helvetica-Oblique")

sig_x, sig_w = 6.6, 6.35
txt(c, "KEY COMMERCIAL SIGNALS", sig_x, 4.55, 11, C_TEAL, font="Helvetica-Bold")
signals = [
    "Ranked 37th nationally in Smart50 2025, its third straight Smart50 year since 2023. (SmartCompany)",
    "Revenue rose from $7.43M in FY23 to $12.8M in FY25 while the team held near 19 to 20 people. (Smart50 2023 to 2025)",
    "Named Victorian State Winner for Outstanding Growth, 2024 Telstra Best of Business Awards. (Telstra)",
    "RentSafe, launched 2021, has completed over 140,000 compliance jobs across 300 real estate offices. (Company site)",
    "Runs three distinct lines: RentSafe compliance, RentRepair maintenance, and manufacturer warranty servicing. (Company site)",
]
sy = 4.95
for sline in signals:
    sy = txt_wrapped(c, sline, sig_x, sy, sig_w, 10, C_DK_TXT, font="Helvetica", leading=13.5, bullet=True)
    sy += 0.08

footer(c)

# ════════════════════════════════════ SLIDE 2 ════════════════════════════════════
c.showPage()
rect(c, 0, 0, 13.333, 7.5, C_WHITE)
header_band(c, "THE OPPORTUNITY")
txt(c, f"{COMPANY}: three commercial observations from ProfitPulse",
    0.35, 1.35, 13, C_BLACK, font="Helvetica-Bold")

col_top, col_h, col_w, col_gap = 1.65, 4.75, 4.02, 0.12
col_fills = [C_TEAL, C_BLACK, C_GOLD]
col_txt = [C_WHITE, C_WHITE, C_BLACK]
headers = [
    "Two products, one blended margin",
    "Two very different customers",
    "Flat headcount, doubling revenue",
]
bodies = [
    "RentSafe compliance checks, RentRepair maintenance subscriptions, and manufacturer warranty servicing are priced and resourced differently, yet revenue is reported as one blended figure. Without a line by line margin view it is hard to tell which line is funding the growth and which is diluting it.",
    "RentSafe and RentRepair serve real estate agencies and property managers, who chase compliance and tenant satisfaction. Warranty servicing serves manufacturers, who chase fault rates and callback costs. These groups pay, churn and cost to serve very differently, and that unevenness is worth mapping.",
    "Revenue rose from $7.43 million in FY23 to $12.8 million in FY25 while the team held at 19 to 20 people, a genuine operating leverage story built on a network of over 5,500 tradespeople rather than headcount. The next stage of growth deserves a costed plan for the platform and working capital.",
]
for i in range(3):
    cx = 0.35 + i * (col_w + col_gap)
    rect(c, cx, col_top, col_w, col_h, col_fills[i])
    tc = col_txt[i]
    txt(c, f"0{i+1}", cx + 0.22, col_top + 0.62, 30, tc, font="Helvetica-Bold")
    txt_wrapped(c, headers[i], cx + 0.22, col_top + 1.15, col_w - 0.4, 15, tc, font="Helvetica-Bold", leading=18)
    txt_wrapped(c, bodies[i], cx + 0.22, col_top + 1.95, col_w - 0.44, 10.5, tc, font="Helvetica", leading=14.5)

txt(c, "These observations are offered in good faith. Taskforce has built a genuinely scaled",
    0.35, 6.62, 10.5, C_DK_TXT, font="Helvetica-Oblique", align="center", max_w=12.6)
txt(c, "national platform. The question is simply whether the margin picture is as clear as the growth curve.",
    0.35, 6.8, 10.5, C_DK_TXT, font="Helvetica-Oblique", align="center", max_w=12.6)

footer(c)

# ════════════════════════════════════ SLIDE 3 ════════════════════════════════════
c.showPage()
rect(c, 0, 0, 13.333, 7.5, C_WHITE)
header_band(c, "THE RECOMMENDATION")

left_x, left_w = 0.35, 7.55
txt(c, "Product and Service Line Profitability", left_x, 1.65, 22, C_BLACK, font="Helvetica-Bold")
txt(c, "$3,950 one off", left_x, 2.1, 16, HexColor("#B87A02"), font="Helvetica-Bold")
txt_wrapped(c,
    "A three week review ranking RentSafe, RentRepair, and warranty servicing by gross margin, "
    "contribution margin, and operational drag, so growth capital follows the line that earns it.",
    left_x, 2.45, left_w, 11.5, C_DK_TXT, font="Helvetica", leading=15.5)

rect_stroke(c, left_x, 3.15, left_w, 1.45, C_TEAL_TINT, C_TEAL, lw=1)
txt(c, "Step one, answer a few quick questions", left_x + 0.2, 3.42, 13, C_TEAL, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry.",
    left_x + 0.2, 3.75, 11, C_DK_TXT, font="Helvetica")
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 13)
c.drawString((left_x + 0.2) * IN, y(4.15), "profit-pulse.com.au/services/find-your-fit")
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          ((left_x + 0.2) * IN, y(4.2), (left_x + 4.4) * IN, y(3.95)), relative=0)

rect(c, left_x, 4.75, left_w, 0.65, C_AMBER_B)
txt(c, "Purchase the suggested product now to get started",
    left_x, 5.13, 13, C_BLACK, font="Helvetica-Bold", align="center", max_w=left_w)
c.linkURL("https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
          (left_x * IN, y(5.4), (left_x + left_w) * IN, y(4.75)), relative=0)

txt(c, "Prefer a conversation first?", left_x, 5.85, 11, C_DK_TXT, font="Helvetica")
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 12)
c.drawString(left_x * IN, y(6.2), "Book a complimentary discovery call")
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (left_x * IN, y(6.25), (left_x + 3.6) * IN, y(6.0)), relative=0)

right_x, right_w = 8.15, 4.83
rect(c, right_x, 1.25, right_w, 5.35, C_BLACK)
txt(c, "NITESH ROOPA", right_x + 0.25, 1.72, 17, C_AMBER_B, font="Helvetica-Bold")
txt(c, "CA, Managing Partner, ProfitPulse", right_x + 0.25, 2.08, 12, C_WHITE, font="Helvetica")
hline(c, right_x + 0.25, 2.28, right_w - 0.5, C_TEAL, lw=1.5)

creds = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3B, Cahora Bassa",
    "Total GRBT project value over AUD 10B",
]
cy = 2.6
for cl in creds:
    txt(c, "•  " + cl, right_x + 0.25, cy, 10.5, C_OFF_WHITE, font="Helvetica")
    cy += 0.3

hline(c, right_x + 0.25, cy + 0.06, right_w - 0.5, C_LINE_GREY, lw=1)
cy += 0.3
contacts = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for cl in contacts:
    txt(c, cl, right_x + 0.25, cy, 10.5, C_OFF_WHITE, font="Helvetica")
    cy += 0.28

footer(c)

c.save()
print("PDF saved:", out_path)
