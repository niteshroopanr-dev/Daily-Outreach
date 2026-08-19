"""
ProfitPulse Brief PDF Builder
Target: My Wealth Solutions | Date: 20 Aug 2026
Direct PDF generation, 3 slides, mirrors build_brief.py layout exactly.
Brand colours only. Zero dashes. Page 960pt x 540pt (13.333in x 7.5in @72dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

W = 960.0
H = 540.0

C_BLACK   = HexColor("#000000")
C_TEAL    = HexColor("#01A296")
C_AMBER_B = HexColor("#F8C806")
C_AMBER_D = HexColor("#F6A102")
C_GOLD    = HexColor("#E3A712")
C_WHITE   = HexColor("#FFFFFF")
C_OFFWH   = HexColor("#E6E5DE")

COMPANY = "My Wealth Solutions"
DATE_STR = "20 Aug 2026"
FOOTER_PREPARED = ("Prepared by Nitesh Roopa CA, Managing Partner and Founder, "
                    "ProfitPulse, Profit-Pulse.com.au")
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_G1_COMMAND = "https://buy.stripe.com/bJe00kc5agKQepG0ZV3ks1x"
BOOKING_LINK = ("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/"
                 "?ismsaljsauthenabled=true")

# ── Fixed chrome geometry (points, 72pt/in) ────────────────────────────────
STRIPE_W = 0.1 * 72
HEADER_H = 1.0 * 72
FOOTER_Y = 7.05 * 72     # top-origin screen y of footer rule
MARGIN_L = 0.35 * 72
MARGIN_R = 0.35 * 72
CONTENT_W = W - MARGIN_L - MARGIN_R


def rl_y(top_y):
    return H - top_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold", align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl_y(y_top + size)
    if align == "right" and max_w:
        x = x + max_w - c.stringWidth(text, font, size)
    elif align == "center" and max_w:
        x = x + (max_w - c.stringWidth(text, font, size)) / 2
    c.drawString(x, baseline, text)


def wrap_lines(c, text, max_w, size, font):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if c.stringWidth(t, font, size) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    if leading is None:
        leading = size * 1.32
    lines = wrap_lines(c, text, max_w, size, font)
    c.setFillColor(colour)
    c.setFont(font, size)
    y = y_top
    for ln in lines:
        c.drawString(x, rl_y(y + size), ln)
        y += leading
    return y


def hline(c, x, y_top, w, colour, thickness=1):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def bullets(c, lines, x, y_top, max_w, size, colour, font="Helvetica", leading_gap=4):
    y = y_top
    for line in lines:
        wrapped = wrap_lines(c, "•  " + line, max_w, size, font)
        c.setFillColor(colour)
        c.setFont(font, size)
        first = True
        for wl in wrapped:
            c.drawString(x, rl_y(y + size), wl)
            y += size * 1.25
            first = False
        y += leading_gap
    return y


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MyWealthSolutions_20Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle(f"{COMPANY} | ProfitPulse Brief | {DATE_STR}")
c.setAuthor("ProfitPulse")
c.setSubject("Operational Intelligence Review")


def chrome(c, eyebrow, right_label, right_colour=C_TEAL):
    fill_rect(c, 0, 0, W, H, C_WHITE)
    fill_rect(c, 0, 0, STRIPE_W, H, C_AMBER_B)
    fill_rect(c, 0, 0, W, HEADER_H, C_BLACK)
    txt(c, eyebrow, MARGIN_L, 32, 12, C_OFFWH, font="Helvetica-Bold")
    txt(c, right_label, W - MARGIN_R - 280, 32, 12, right_colour, font="Helvetica-Bold",
        align="right", max_w=280)
    hline(c, MARGIN_L, FOOTER_Y, CONTENT_W, HexColor("#CCCCCC"), 0.75)
    txt(c, FOOTER_PREPARED, MARGIN_L, 512, 8, HexColor("#666666"), font="Helvetica")
    txt(c, DATE_STR, W - MARGIN_R - 180, 512, 8, HexColor("#666666"), font="Helvetica",
        align="right", max_w=180)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ════════════════════════════════════════════════════════════════════════
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF", COMPANY)

txt(c, COMPANY, MARGIN_L, 88, 34, C_BLACK, font="Helvetica-Bold")
txt(c, "Financial planning and wealth advisory group, Newstead, Brisbane QLD",
    MARGIN_L, 138, 13, HexColor("#333333"), font="Helvetica")

CARD_TOP = 2.4 * 72
CARD_H = 1.6 * 72
N_CARDS = 6
GAP = 0.15 * 72
CARD_W = (CONTENT_W - GAP * (N_CARDS - 1)) / N_CARDS

cards = [
    ("$9.9M", ["FY2025 revenue,", "Smart50 profile"], "SmartCompany Smart50 2025"),
    ("47%",   ["Revenue growth", "rate, Smart50"], "SmartCompany Smart50 2025"),
    ("#25",   ["Smart50 2025", "rank of 50"], "SmartCompany Smart50 2025"),
    ("55",    ["People across", "six offices"], "Company website, 2026"),
    ("2011",  ["Founded in", "Newstead, QLD"], "Company website, our story"),
    ("3x",    ["GPS Wealth Practice", "of the Year winner"], "GPS Wealth conference recaps"),
]

for i, (num, lbl, src) in enumerate(cards):
    x = MARGIN_L + i * (CARD_W + GAP)
    fill_rect(c, x, CARD_TOP, CARD_W, CARD_H, C_BLACK)
    fill_rect(c, x, CARD_TOP, CARD_W, 3, C_TEAL)
    txt(c, num, x + 9, CARD_TOP + 12, 24, C_AMBER_B, font="Helvetica-Bold")
    ly = CARD_TOP + 50
    for l in lbl:
        txt(c, l, x + 9, ly, 10, C_OFFWH, font="Helvetica")
        ly += 14
    txt(c, src, x + 9, CARD_TOP + 96, 7, C_TEAL, font="Helvetica-Oblique")

SIG_TOP = 4.28 * 72
txt(c, "KEY COMMERCIAL SIGNALS", MARGIN_L, SIG_TOP, 12, C_TEAL, font="Helvetica-Bold")

signals = [
    "Ranked 25th nationally, 2025 Smart50 Awards. Source: SmartCompany Smart50 2025 profile.",
    "Grew to six offices: Newstead, Sydney CBD, Neutral Bay, Gold Coast, Melbourne, Townsville. Source: company website.",
    "Named GPS Wealth Practice of the Year three years running. Source: company blog, GPS Wealth conference recaps.",
    "Recognised in the AFR Fast 100 and the 2025 ifa Excellence Awards. Source: company website, ifa Excellence Awards 2025.",
    "Founded 2011 by Guy Freeman and Ben Budge, both trained together in financial planning. Source: company Our Story page.",
    "Rated a top 1 percent Australian financial planner, Brisbane Outer Suburbs 2026. Source: Quality Business Awards Australia.",
]
bullets(c, signals, MARGIN_L, SIG_TOP + 28, CONTENT_W, 11, HexColor("#222222"))

c.showPage()

# ════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ════════════════════════════════════════════════════════════════════════
chrome(c, "THE OPPORTUNITY", COMPANY)

txt(c, f"{COMPANY}: three commercial observations from ProfitPulse",
    MARGIN_L, 86, 15, C_BLACK, font="Helvetica-Bold")

COL_TOP = 1.7 * 72
COL_H = 4.55 * 72
COL_GAP = 0.15 * 72
COL_W = (CONTENT_W - COL_GAP * 2) / 3

observations = [
    ("01", C_TEAL, C_BLACK,
     "Six offices, one performance picture, or not yet",
     "The practice has grown from one Newstead office in 2011 to six locations "
     "across Queensland, New South Wales, and Victoria, lifting revenue 47 "
     "percent to 9.9 million dollars in the year that earned Smart50 rank 25. "
     "Each office carries its own referral splits, servicing cost, and "
     "licensee fees. Without a location by location margin view, a quietly "
     "subsidised office can look identical to a strong one."),
    ("02", C_BLACK, C_WHITE,
     "Fifty five people is a company, not a practice",
     "Headcount has grown to 55 alongside revenue, a point where informal, "
     "founder led oversight usually starts to strain. Three consecutive GPS "
     "Wealth Practice of the Year wins confirm client outcomes are strong. "
     "What those awards do not confirm is whether cost base and workforce "
     "capacity are still returning a proportional margin as the team has "
     "grown."),
    ("03", C_GOLD, C_BLACK,
     "The award record is proof, the numbers are the next step",
     "An AFR Fast 100 mention, an ifa Excellence Awards nod, and three GPS "
     "Wealth trophies make this one of the more credentialed growth stories "
     "in Australian financial advice. That record carries real weight with a "
     "bank, a licensee, or a future partner. Businesses that convert growth "
     "stories into value are the ones holding clean, location level numbers "
     "behind the headline."),
]

for i, (idx, fill, textcol, header, para) in enumerate(observations):
    x = MARGIN_L + i * (COL_W + COL_GAP)
    fill_rect(c, x, COL_TOP, COL_W, COL_H, fill)
    txt(c, idx, x + 16, COL_TOP + 14, 30, (C_OFFWH if fill == C_BLACK else C_BLACK),
        font="Helvetica-Bold")
    txt_wrapped(c, header, x + 16, COL_TOP + 68, COL_W - 32, 14, textcol,
                font="Helvetica-Bold", leading=18)
    txt_wrapped(c, para, x + 16, COL_TOP + 132, COL_W - 32, 10.5, textcol,
                font="Helvetica", leading=15)

CLOSE_TOP = COL_TOP + COL_H + 10
txt_wrapped(c,
    "These observations are offered in good faith. My Wealth Solutions has "
    "built something genuinely impressive in fifteen years. The question is "
    "simply whether the financial architecture across six offices is keeping "
    "pace with the ambition.",
    MARGIN_L, CLOSE_TOP, CONTENT_W, 11, HexColor("#333333"), font="Helvetica-Oblique", leading=15)

c.showPage()

# ════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ════════════════════════════════════════════════════════════════════════
chrome(c, "THE RECOMMENDATION", "PROFITPULSE")

LEFT_X = MARGIN_L
LEFT_W = 7.4 * 72
RIGHT_X = 8.05 * 72
RIGHT_W = 4.93 * 72

txt(c, "Operational Intelligence Review", LEFT_X, 86, 22, C_BLACK, font="Helvetica-Bold")
txt(c, "$6,500 one off", LEFT_X, 128, 17, C_AMBER_D, font="Helvetica-Bold")
txt_wrapped(c,
    "A six week review across four lenses: customer concentration and "
    "profitability, product and service line margin, workforce capacity, "
    "and operational bottlenecks, mapped office by office and adviser by "
    "adviser.",
    LEFT_X, 160, LEFT_W, 11.5, HexColor("#222222"), font="Helvetica", leading=16)

step_top = 3.28 * 72
stroke_rect(c, LEFT_X, step_top, LEFT_W, 1.35 * 72, C_TEAL, 1)
txt(c, "Step one, answer a few quick questions", LEFT_X + 14, step_top + 14, 13, C_TEAL,
    font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry.", LEFT_X + 14, step_top + 48,
    11, HexColor("#222222"), font="Helvetica")
q_y = step_top + 82
txt(c, QUESTIONNAIRE_CLEAN, LEFT_X + 14, q_y, 13, C_TEAL, font="Helvetica-Bold")
c.linkURL(QUESTIONNAIRE_URL, (LEFT_X + 14, rl_y(q_y + 16), LEFT_X + 14 + 300, rl_y(q_y)),
          relative=0)

btn_top = 4.85 * 72
btn_w = 4.6 * 72
btn_h = 0.55 * 72
stroke_rect(c, LEFT_X, btn_top, btn_w, btn_h, C_AMBER_D, 1.5)
btn_label = "Purchase the suggested product now to get started"
txt(c, btn_label, LEFT_X, btn_top + btn_h / 2 - 5, 11, C_AMBER_D, font="Helvetica-Bold",
    align="center", max_w=btn_w)
c.linkURL(STRIPE_G1_COMMAND, (LEFT_X, rl_y(btn_top + btn_h), LEFT_X + btn_w, rl_y(btn_top)),
          relative=0)

txt(c, "Prefer a conversation first?", LEFT_X, 5.65 * 72, 11, HexColor("#222222"), font="Helvetica")
book_y = 5.96 * 72
txt(c, "Book a complimentary discovery call", LEFT_X, book_y, 12, C_TEAL, font="Helvetica-Bold")
c.linkURL(BOOKING_LINK, (LEFT_X, rl_y(book_y + 16), LEFT_X + 260, rl_y(book_y)), relative=0)

# Right column credibility panel
fill_rect(c, RIGHT_X, 1.15 * 72, RIGHT_W, 5.75 * 72, C_BLACK)
txt(c, "NITESH ROOPA", RIGHT_X + 18, 1.35 * 72, 16, C_AMBER_B, font="Helvetica-Bold")
txt(c, "CA, Managing Partner, ProfitPulse", RIGHT_X + 18, 1.78 * 72, 12, C_WHITE, font="Helvetica")
hline(c, RIGHT_X + 18, 2.18 * 72, RIGHT_W - 36, C_TEAL, 1.5)

cred = [
    "16 years of experience across 4 countries",
    "Over 52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique",
    "Over AUD 10 billion Gympie Road Bypass Tunnel, total project value",
]
cy = bullets(c, cred, RIGHT_X + 18, 2.35 * 72, RIGHT_W - 40, 10.5, C_OFFWH)

hline(c, RIGHT_X + 18, 4.35 * 72, RIGHT_W - 36, HexColor("#333333"), 1)

contact = [
    ("Profit-Pulse.com.au", C_OFFWH),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL),
    ("+61 411 876 267", C_OFFWH),
    ("linkedin.com/in/nitesh-roopa-77594163", C_TEAL),
]
cy2 = 4.55 * 72
for label, colour in contact:
    txt(c, label, RIGHT_X + 18, cy2, 11, colour, font="Helvetica")
    cy2 += 22

c.save()
print(f"PDF saved: {out_path}")
