"""
ProfitPulse Brief PDF Builder
Target: National Media | Date: 13 Jun 2026
Section 6 house style: white body, amber left stripe, black header band.
Page: 960 x 540 pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.units import inch

W = 960
H = 540

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WH   = HexColor("#E6E5DE")
C_MID_GRY  = HexColor("#888888")
C_DRK_GRY  = HexColor("#444444")
C_LIGHT_BG = HexColor("#F4F4F4")

QUESTIONNAIRE_URL = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_URL        = "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D"
BOOKING_URL       = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def rl(y):
    """Screen y (top=0) to ReportLab y (bottom=0)."""
    return H - y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl(y_top + h), w, h, fill=1, stroke=0)


def fill_stroke_rect(c, x, y_top, w, h, fill, stroke, lw=1):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(lw)
    c.rect(x, rl(y_top + h), w, h, fill=1, stroke=1)


def hline(c, x, y_top, w, colour, thickness=1.2):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, rl(y_top), x + w, rl(y_top))


def t(c, text, x, y_top, size, colour, font="Helvetica", align="left", max_w=None):
    """Single line text. y_top is top of text area in screen coords."""
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl(y_top + size * 1.1)
    if align == "right" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, baseline, text)


def tw(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    """Wrapped text. Returns next y_top after last line."""
    if leading is None:
        leading = size * 1.45
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
    lines, cur = [], ""
    for word in words:
        test = (cur + " " + word).strip()
        if c.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    y = y_top
    for line in lines:
        c.drawString(x, rl(y + size * 1.1), line)
        y += leading
    return y


def link_rect(c, x, y_top, w, h, url):
    """Attach a clickable hyperlink to a rectangle area."""
    c.linkURL(url, (x, rl(y_top + h), x + w, rl(y_top)), relative=0)


def header_band(c, eyebrow, company):
    fill_rect(c, 0, 0, W, 68, C_BLACK)
    t(c, eyebrow, 22, 16, 11, C_OFF_WH, font="Helvetica-Bold")
    t(c, company, 22, 16, 13, C_TEAL, font="Helvetica-Bold",
      align="right", max_w=W - 44)


def footer_line(c, date="13 Jun 2026"):
    hline(c, 22, 514, W - 44, C_TEAL, 0.8)
    t(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, "
         "ProfitPulse, Profit-Pulse.com.au",
      22, 520, 8, C_DRK_GRY, font="Helvetica")
    t(c, date, 22, 520, 8, C_DRK_GRY, font="Helvetica",
      align="right", max_w=W - 44)


def stat_card(c, x, y, cw, ch, number, label1, label2, source):
    """Black tile with teal top accent, large number, two-line label, source."""
    fill_rect(c, x, y, cw, ch, C_BLACK)
    fill_rect(c, x, y, cw, 5, C_TEAL)
    t(c, number, x + 8, y + 8, 20, C_OFF_WH, font="Helvetica-Bold")
    if label1:
        t(c, label1, x + 8, y + 35, 9, C_OFF_WH, font="Helvetica")
    if label2:
        t(c, label2, x + 8, y + 48, 9, C_MID_GRY, font="Helvetica")
    t(c, source, x + 8, y + 63, 7, C_TEAL, font="Helvetica-Oblique")


# ================================================================
# OUTPUT
# ================================================================
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_NationalMedia_13Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("National Media | ProfitPulse Brief | 13 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Product and Service Line Profitability")


# ================================================================
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# ================================================================
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)    # amber left stripe
header_band(c, "COMMERCIAL INTELLIGENCE BRIEF", "NATIONAL MEDIA")

# Large company name
t(c, "NATIONAL MEDIA", 22, 74, 34, C_BLACK, font="Helvetica-Bold")
# Descriptor
t(c, "B2B Trade Exhibitions and Events Producer, Bundall, Gold Coast QLD",
  22, 116, 12, C_TEAL, font="Helvetica")

# Stat cards (6 across)
CARD_W, CARD_H = 152, 82
CARD_GAP = 7
CARD_Y = 134
cx = 22
cards = [
    ("$18.4M",  "Revenue",       "FY2025",                "Smart50 2025 citation"),
    ("42%",     "Three Year",    "Average Growth",        "Smart50 2025 citation"),
    ("Rank 28", "Smart50",       "Australia 2025",        "SmartCompany Smart50 2025"),
    ("48",      "Team Members",  "",                      "Smart50 2025 citation"),
    ("9+",      "Live Event",    "Brands",                "nationalmedia.com.au"),
    ("550+",    "Combined WHS",  "Show Exhibitors",       "nationalmedia.com.au"),
]
for num, l1, l2, src in cards:
    stat_card(c, cx, CARD_Y, CARD_W, CARD_H, num, l1, l2, src)
    cx += CARD_W + CARD_GAP

# Key Commercial Signals
sy = 228
t(c, "KEY COMMERCIAL SIGNALS", 22, sy, 10, C_TEAL, font="Helvetica-Bold")
hline(c, 22, sy + 14, W - 44, C_TEAL, 1)

signals = [
    ("FutureBuild Australia launched 11 to 13 June 2026 at ICC Sydney: new event brand debut confirmed",
     "nationalmedia.com.au"),
    ("Mark Harvey presented at SISO CEO Summit on Organic Growth via Launches: expansion strategy confirmed publicly",
     "LinkedIn public profile"),
    ("Workplace Health and Safety Show: 550 plus exhibitors across three consecutive state editions in NSW, QLD, and VIC",
     "nationalmedia.com.au"),
    ("Food and Hospitality Week bundles four separate trade shows: Restaurant and Foodservice, Pizza Pasta and Italian, Cafe and Coffee, and RESTECH",
     "nationalmedia.com.au"),
    ("Smart50 2025 rank 28 of 50: revenue verified at $18.4 million with 42 percent three year average growth",
     "SmartCompany Smart50 2025 award citation"),
    ("Founded 1993: over three decades of operations from Bundall, Gold Coast QLD across multiple Australian industry sectors",
     "nationalmedia.com.au"),
]
row_y = sy + 20
for sig_text, sig_src in signals:
    t(c, sig_text, 22, row_y, 9, C_BLACK, font="Helvetica")
    t(c, "Source: " + sig_src, 22, row_y + 12, 7, C_MID_GRY, font="Helvetica-Oblique")
    row_y += 28

footer_line(c)
c.showPage()


# ================================================================
# PAGE 2: THE OPPORTUNITY
# ================================================================
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
header_band(c, "THE OPPORTUNITY", "NATIONAL MEDIA")

t(c, "National Media: Three commercial observations from ProfitPulse",
  22, 72, 10, C_DRK_GRY, font="Helvetica-Oblique")

# Three columns
COL_W = (W - 7) // 3   # ~317 pt each
COL_Y = 88
COL_H = 390
col_fills = [C_TEAL, C_BLACK, C_GOLD]
col_x_arr = [7, 7 + COL_W, 7 + 2 * COL_W]
txt_colours = [C_BLACK, C_OFF_WH, C_BLACK]

obs_data = [
    ("01", "Nine event brands, one financial picture",
     ("National Media runs nine or more distinct event brands across food and hospitality, "
      "workplace health and safety, construction, and accommodation technology. "
      "With $18.4 million in revenue across this portfolio, the aggregate 42 percent "
      "growth headline looks strong. The question that matters at this stage is whether "
      "that growth is distributed evenly or concentrated in a few events carrying the rest. "
      "In most multi brand businesses, two or three products generate the bulk of the real "
      "margin while others absorb shared overhead without full visibility.")),
    ("02", "Each new launch commits capital before revenue arrives",
     ("Mark Harvey's strategy of organic growth via new event launches, stated publicly "
      "at the SISO CEO Summit, means the portfolio is actively expanding. FutureBuild "
      "Australia at ICC Sydney in June 2026 is the most recent addition. Each launch commits "
      "venue deposits, exhibitor acquisition spend, and team time before the first stand is "
      "sold. Without a clear view of which existing events produce the strongest returns, "
      "capital allocation for new launches defaults to experience and instinct rather "
      "than evidence.")),
    ("03", "The next 42 percent requires deliberate capital allocation",
     ("With nine or more event brands running across three states and an active launch "
      "strategy, the financial complexity of the portfolio is rising. Deciding which shows "
      "to scale, which to price more assertively, and which to rationalise requires an "
      "objective ranking of every event by gross margin and contribution. The Product and "
      "Service Line Profitability review produces exactly that: every brand ranked by "
      "financial performance in three weeks, built from management accounts, delivered "
      "as a clear action list.")),
]

for i, (idx, hdr, para) in enumerate(obs_data):
    cx2 = col_x_arr[i]
    fill_rect(c, cx2, COL_Y, COL_W, COL_H, col_fills[i])
    tc2 = txt_colours[i]
    t(c, idx, cx2 + 14, COL_Y + 12, 26, tc2, font="Helvetica-Bold")
    tw(c, hdr, cx2 + 14, COL_Y + 48, COL_W - 28, 12, tc2, font="Helvetica-Bold", leading=17)
    tw(c, para, cx2 + 14, COL_Y + 90, COL_W - 28, 9.5, tc2, font="Helvetica", leading=14)

# Warm line below columns
warm = ("These observations are offered in the spirit of a genuine commercial conversation. "
        "National Media has built something impressive over three decades. "
        "The question is simply whether the financial intelligence behind the portfolio "
        "matches the ambition of the growth strategy.")
tw(c, warm, 22, 486, W - 44, 9, C_DRK_GRY, font="Helvetica-Oblique", leading=13)

footer_line(c)
c.showPage()


# ================================================================
# PAGE 3: THE RECOMMENDATION
# ================================================================
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
header_band(c, "THE RECOMMENDATION", "NATIONAL MEDIA")

# LEFT COLUMN: recommendation block
LX  = 22
LW  = 570
LY0 = 76

t(c, "Product and Service Line Profitability",
  LX, LY0, 20, C_BLACK, font="Helvetica-Bold")
t(c, "$3,950 one off",
  LX, LY0 + 28, 15, C_AMBER_D, font="Helvetica-Bold")

desc = ("A three week project ranking every event brand by gross margin, "
        "contribution margin, and operational drag. For a nine brand events "
        "business growing at 42 percent, this identifies which shows to scale, "
        "which to reprice, and which to rationalise. "
        "Verified ProfitPulse price. The questionnaire confirms the exact fit.")
tw(c, desc, LX, LY0 + 54, LW, 10, C_BLACK, font="Helvetica", leading=15)

# Step one block
fill_rect(c, LX, LY0 + 130, LW, 94, C_LIGHT_BG)
t(c, "STEP ONE: ANSWER A FEW QUICK QUESTIONS",
  LX + 10, LY0 + 140, 10, C_TEAL, font="Helvetica-Bold")
t(c, "See the solutions matched to your size and industry.",
  LX + 10, LY0 + 158, 10, C_BLACK, font="Helvetica")
# Clean questionnaire address (no UTM on brief per Rule 6.4)
t(c, "profit-pulse.com.au/full-suite-of-products",
  LX + 10, LY0 + 175, 11, C_TEAL, font="Helvetica")
link_rect(c, LX + 10, LY0 + 173, 400, 16, QUESTIONNAIRE_URL)

# CTA button (Stripe link hidden behind text)
fill_rect(c, LX, LY0 + 232, LW, 34, C_AMBER_D)
t(c, "Purchase the suggested product now to get started",
  LX, LY0 + 240, 13, C_BLACK, font="Helvetica-Bold",
  align="center", max_w=LW)
link_rect(c, LX, LY0 + 232, LW, 34, STRIPE_URL)

# Conversation alternative
t(c, "Prefer a conversation first?", LX, LY0 + 278, 10, C_BLACK, font="Helvetica")
t(c, "Book a complimentary discovery call",
  LX, LY0 + 295, 11, C_TEAL, font="Helvetica-Bold")
link_rect(c, LX, LY0 + 293, 280, 16, BOOKING_URL)

# RIGHT COLUMN: credibility panel
RX = 610
RW = 330
RY = 76
RH = 430

fill_stroke_rect(c, RX, RY, RW, RH, C_BLACK, HexColor("#222222"), 1)

t(c, "NITESH ROOPA", RX + 16, RY + 16, 17, C_AMBER_B, font="Helvetica-Bold")
t(c, "CA, Managing Partner", RX + 16, RY + 42, 11, C_WHITE, font="Helvetica")
t(c, "ProfitPulse", RX + 16, RY + 62, 17, C_TEAL, font="Helvetica-Bold")
hline(c, RX + 16, RY + 92, RW - 32, C_TEAL, 1.2)

creds = [
    ("Profit-Pulse.com.au",                              C_OFF_WH, "Helvetica",    10),
    ("Nitesh@Profit-Pulse.com.au",                       C_TEAL,   "Helvetica",    10),
    ("+61 411 876 267",                                  C_OFF_WH, "Helvetica",    10),
    ("",                                                 C_OFF_WH, "Helvetica",     5),
    ("16 years across 4 countries",                      C_MID_GRY,"Helvetica",     9),
    ("52 deals executed and managed",                    C_MID_GRY,"Helvetica",     9),
    ("Largest single deal USD 1.3 billion Cahora Bassa", C_MID_GRY,"Helvetica",     9),
    ("Total GRBT project value over AUD 10 billion",     C_MID_GRY,"Helvetica",     9),
    ("",                                                 C_OFF_WH, "Helvetica",     5),
    ("linkedin.com/in/nitesh-roopa-77594163",            C_DRK_GRY,"Helvetica",     8),
]

cy = RY + 102
for txt_str, col, fnt, sz in creds:
    if txt_str:
        t(c, txt_str, RX + 16, cy, sz, col, font=fnt)
    cy += sz * 1.6 if txt_str else sz * 2

footer_line(c)
c.save()
print(f"PDF saved: {out_path}")
