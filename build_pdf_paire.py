"""
ProfitPulse Brief PDF Builder v3.3 house style
Target: Paire | Date: 25 Aug 2026
Direct PDF generation mirroring build_brief_paire.py (the PPTX) slide for slide,
coordinate for coordinate, so the PDF matches the deck exactly.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi, widescreen).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import simpleSplit

IN = 72.0  # points per inch
W = 13.333 * IN
H = 7.5 * IN

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WHITE= HexColor("#E6E5DE")
C_LGREY    = HexColor("#F4F4F2")
C_MGREY    = HexColor("#AAAAAA")
C_DGREY1   = HexColor("#444444")
C_DGREY2   = HexColor("#666666")
C_DGREY3   = HexColor("#777777")
C_DGREY4   = HexColor("#9A9A9A")
C_TEXT     = HexColor("#222222")

SERIF_B = "Times-Bold"
SANS    = "Helvetica"
SANS_B  = "Helvetica-Bold"
SANS_I  = "Helvetica-Oblique"

DATE_STR = "25 Aug 2026"

QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL   = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_C2_COMMAND   = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_LINK        = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def rl_y(y_top):
    """Convert top origin y (points) to ReportLab bottom origin y."""
    return H - y_top


def rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def text(c, s, x, y_top, size, colour, font=SANS, align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl_y(y_top + size * 0.92)
    draw_x = x
    if align in ("right", "center") and max_w:
        tw = c.stringWidth(s, font, size)
        if align == "right":
            draw_x = x + max_w - tw
        else:
            draw_x = x + (max_w - tw) / 2
    c.drawString(draw_x, baseline, s)


def wrapped(c, s, x, y_top, max_w, size, colour, font=SANS, leading=None, align="left"):
    if leading is None:
        leading = size * 1.28
    lines = simpleSplit(s, font, size, max_w)
    y = y_top
    for line in lines:
        text(c, line, x, y, size, colour, font=font, align=align, max_w=max_w)
        y += leading
    return y


def multiline(c, lines, x, y_top, size, colour, font=SANS, gap=6, leading=None):
    if leading is None:
        leading = size * 1.15
    y = y_top
    for item in lines:
        if isinstance(item, str):
            s, sz, col, f = item, size, colour, font
        else:
            s = item.get("text", "")
            sz = item.get("size", size)
            col = item.get("colour", colour)
            f = item.get("font", font)
        text(c, s, x, y, sz, col, font=f)
        y += leading + gap
    return y


def house_chrome(c, eyebrow):
    rect(c, 0, 0, 0.1 * IN, H, C_AMBER_D)
    rect(c, 0.1 * IN, 0, W - 0.1 * IN, 1.0 * IN, C_BLACK)
    text(c, eyebrow, 0.35 * IN, 0.32 * IN, 12, C_OFF_WHITE, font=SANS_B)
    text(c, "PROFITPULSE", 9.5 * IN, 0.32 * IN, 12, C_TEAL, font=SANS_B, align="right", max_w=3.6 * IN)
    rect(c, 0.35 * IN, 7.02 * IN, 12.6 * IN, 0.75, C_MGREY)
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         0.35 * IN, 7.08 * IN, 8, C_DGREY2, font=SANS)
    text(c, DATE_STR, 10.5 * IN, 7.08 * IN, 8, C_DGREY2, font=SANS, align="right", max_w=2.5 * IN)


def stat_card(c, x, y_top, w, h, number, label_lines, source):
    rect(c, x, y_top, w, h, C_BLACK)
    rect(c, x, y_top, w, 3, C_TEAL)
    text(c, number, x + 0.12 * IN, y_top + 0.12 * IN, 27, C_AMBER_B, font=SERIF_B)
    multiline(c, label_lines, x + 0.12 * IN, y_top + 0.62 * IN, 10, C_OFF_WHITE, font=SANS, gap=1, leading=12)
    text(c, source, x + 0.12 * IN, y_top + h - 0.32 * IN, 7, C_DGREY4, font=SANS_I)


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_25Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Paire | ProfitPulse Brief | 25 Aug 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Working Capital Unlock")

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ═════════════════════════════════════════════════════════════════════════
rect(c, 0, 0, W, H, C_WHITE)
house_chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")

text(c, "Paire", 0.35 * IN, 1.12 * IN, 42, C_BLACK, font=SERIF_B)
text(c, "Direct to consumer sock and essentials brand, South Melbourne VIC",
     0.35 * IN, 1.9 * IN, 13, C_DGREY1, font=SANS)

cards = [
    ("$10.0M", ["Revenue, FY2025"], "SmartCompany Smart50, Nov 2025"),
    ("15/50",  ["Smart50 2025 rank"], "SmartCompany, Nov 2025"),
    ("300K+",  ["Global customers"], "Inside Retail, Aug 2026"),
    ("2020",   ["Founded, garage", "start in Melbourne"], "Power Retail profile"),
    ("60+",    ["Products across", "four categories"], "Inside Retail, Aug 2026"),
    ("26",     ["Team, full time", "and casual staff"], "SmartCompany Smart50 profile"),
]
card_w = 1.98 * IN
gap_w = 0.15 * IN
x = 0.35 * IN
y_cards = 2.35 * IN
card_h = 1.5 * IN
for number, label_lines, source in cards:
    stat_card(c, x, y_cards, card_w, card_h, number, label_lines, source)
    x += card_w + gap_w

# Chart
chart_x, chart_y, chart_w, chart_h = 0.35 * IN, 4.12 * IN, 3.9 * IN, 2.55 * IN
rect(c, chart_x, chart_y, chart_w, chart_h, C_LGREY)
text(c, "REVENUE, FY22 TO FY25", chart_x + 0.18 * IN, chart_y + 0.12 * IN, 10, C_TEAL, font=SANS_B)

base_y = chart_y + 2.05 * IN
max_bar_h = 1.35 * IN
bar1_h = max_bar_h * (1.7 / 10.0)
bar1_x, bar1_w = chart_x + 1.0 * IN, 0.85 * IN
rect(c, bar1_x, base_y - bar1_h, bar1_w, bar1_h, C_TEAL)
text(c, "$1.7M", bar1_x - 0.2 * IN, base_y - bar1_h - 0.32 * IN, 10, C_BLACK, font=SANS_B, align="center", max_w=1.25 * IN)
text(c, "FY2022", bar1_x - 0.2 * IN, base_y + 0.06 * IN, 9, C_DGREY2, font=SANS, align="center", max_w=1.25 * IN)

bar2_h = max_bar_h
bar2_x, bar2_w = chart_x + 2.3 * IN, 0.85 * IN
rect(c, bar2_x, base_y - bar2_h, bar2_w, bar2_h, C_AMBER_D)
text(c, "$10.0M", bar2_x - 0.2 * IN, base_y - bar2_h - 0.32 * IN, 10, C_BLACK, font=SANS_B, align="center", max_w=1.25 * IN)
text(c, "FY2025", bar2_x - 0.2 * IN, base_y + 0.06 * IN, 9, C_DGREY2, font=SANS, align="center", max_w=1.25 * IN)

rect(c, chart_x + 0.18 * IN, base_y, chart_w - 0.36 * IN, 1, C_MGREY)
text(c, "Source: SmartCompany Smart50 2025 growth profile", chart_x + 0.18 * IN,
     chart_y + chart_h - 0.3 * IN, 7, C_DGREY3, font=SANS_I)

sig_x, sig_y, sig_w = 4.55 * IN, 4.12 * IN, 8.43 * IN
text(c, "KEY COMMERCIAL SIGNALS", sig_x, sig_y, 11, C_TEAL, font=SANS_B)
signals = [
    "Revenue grew from $1.7M (FY22) to $10.0M (FY25), Smart50 2025 rank 15 of 50.",
    "Converted pop up trials into two permanent stores, QV Melbourne and South Melbourne.",
    "Entered the United States market in 2026, its first launch outside ANZ.",
    "Also trading into Singapore and Malaysia; range now exceeds 60 SKUs.",
    "Declined Shark Tank Australia offers in November 2024, staying founder funded.",
    "Planning further Melbourne and Sydney stores within the next 12 to 18 months.",
]
multiline(c, signals, sig_x, sig_y + 0.4 * IN, 11, C_TEXT, font=SANS, gap=8, leading=13)
text(c, "Sources: SmartCompany Smart50 2025 series, Inside Retail Aug 2026, FashionUnited Aug 2026",
     sig_x, 6.55 * IN, 7, C_DGREY3, font=SANS_I)

c.showPage()

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ═════════════════════════════════════════════════════════════════════════
rect(c, 0, 0, W, H, C_WHITE)
house_chrome(c, "THE OPPORTUNITY")
text(c, "Paire: three commercial observations from ProfitPulse", 0.35 * IN, 1.02 * IN, 13, C_BLACK, font=SANS_B)

col_y, col_h, col_w, col_gap = 1.55 * IN, 4.0 * IN, 4.14 * IN, 0.15 * IN
col_x = [0.35 * IN, 0.35 * IN + col_w + col_gap, 0.35 * IN + 2 * (col_w + col_gap)]
fills = [C_TEAL, C_BLACK, C_GOLD]
text_colours = [C_BLACK, C_WHITE, C_BLACK]
headers = [
    "Cash is funding four fronts at once",
    "SKU growth has outpaced margin visibility",
    "Growth is entirely founder funded",
]
bodies = [
    ("Paire now trades in five markets at once: Australia, the US, Singapore and "
     "Malaysia online, plus two Melbourne stores. Each market and each store needs "
     "its own stock position before a single sale converts back to cash. Funding "
     "this many fronts simultaneously, on top of a wider domestic rollout already "
     "planned, is a working capital question before it is anything else. "
     "A Working Capital Unlock maps exactly where that cash is trapped."),
    ("Sixty plus products now sit across socks, underwear, activewear and outerwear, "
     "sold through direct online, two flagship stores and three overseas markets. At "
     "$1.7 million the product mix was simple to read by eye. At $10.0 million across "
     "this many channels, blended revenue growth can mask lines and markets that are "
     "actually thin on margin once freight, duty and store overhead are allocated "
     "properly against each one."),
    ("Paire walked away from Shark Tank offers in November 2024 rather than give up "
     "2.5 percent of the business. That discipline means every store opening and "
     "every new market this year is funded from trading cash, not investor cash, "
     "which raises the cost of any dollar left sitting idle in stock or receivables, "
     "and makes disciplined cash management the difference between funding the next "
     "market and stalling before it."),
]
for i in range(3):
    rect(c, col_x[i], col_y, col_w, col_h, fills[i])
    text(c, f"0{i+1}", col_x[i] + 0.22 * IN, col_y + 0.18 * IN, 34, text_colours[i], font=SERIF_B)
    text(c, headers[i], col_x[i] + 0.22 * IN, col_y + 0.95 * IN, 15, text_colours[i], font=SERIF_B)
    wrapped(c, bodies[i], col_x[i] + 0.22 * IN, col_y + 1.75 * IN, col_w - 0.44 * IN, 11, text_colours[i], font=SANS, leading=14.5)

wrapped(c, ("These are observations offered in good faith. Paire has built something genuinely "
            "impressive from a garage start. The question is simply whether the cash engine is "
            "built to match the ambition of five markets at once."),
        0.35 * IN, 5.95 * IN, 12.6 * IN, 11.5, C_DGREY1, font=SANS_I, leading=15)

c.showPage()

# ═════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ═════════════════════════════════════════════════════════════════════════
rect(c, 0, 0, W, H, C_WHITE)
house_chrome(c, "THE RECOMMENDATION")

left_x, left_w = 0.35 * IN, 7.6 * IN
right_x, right_w = 8.25 * IN, 4.73 * IN

text(c, "Working Capital Unlock", left_x, 1.15 * IN, 25, C_BLACK, font=SERIF_B)
text(c, "$6,000 one off", left_x, 1.68 * IN, 15, C_TEAL, font=SANS_B)
wrapped(c, ("A four week project that maps cash trapped in stock, receivables, supplier terms "
            "and banking facilities across every Paire channel and market, then hands over a "
            "prioritised release plan. Typical clients release 8 to 15 percent of revenue back "
            "into the business without raising outside capital."),
        left_x, 2.08 * IN, left_w, 11, C_TEXT, font=SANS, leading=14)

rect(c, left_x, 3.25 * IN, left_w, 1.35 * IN, C_LGREY)
text(c, "Step one, answer a few quick questions", left_x + 0.2 * IN, 3.38 * IN, 12, C_BLACK, font=SANS_B)
text(c, "See the solutions matched to your size and industry.", left_x + 0.2 * IN, 3.72 * IN, 10.5, HexColor("#333333"), font=SANS)
text(c, QUESTIONNAIRE_CLEAN, left_x + 0.2 * IN, 4.05 * IN, 12, C_TEAL, font=SANS_B)
c.linkURL(QUESTIONNAIRE_URL, (left_x + 0.2 * IN, rl_y(4.05 * IN + 15), left_x + 0.2 * IN + 3.5 * IN, rl_y(4.05 * IN)), relative=0)

rect(c, left_x, 4.78 * IN, left_w, 0.62 * IN, C_TEAL)
text(c, "Purchase the suggested product now to get started", left_x, 4.78 * IN + 0.22 * IN, 13, C_BLACK,
     font=SANS_B, align="center", max_w=left_w)
c.linkURL(STRIPE_C2_COMMAND, (left_x, rl_y(4.78 * IN + 0.62 * IN), left_x + left_w, rl_y(4.78 * IN)), relative=0)

text(c, "Prefer a conversation first?", left_x, 5.62 * IN, 11, HexColor("#333333"), font=SANS)
text(c, "Book a complimentary discovery call", left_x, 5.92 * IN, 11, C_TEAL, font=SANS_B)
c.linkURL(BOOKING_LINK, (left_x, rl_y(5.92 * IN + 15), left_x + 3.3 * IN, rl_y(5.92 * IN)), relative=0)

rect(c, right_x, 1.15 * IN, right_w, 4.75 * IN, C_BLACK)
text(c, "Nitesh Roopa", right_x + 0.25 * IN, 1.35 * IN, 18, C_AMBER_B, font=SERIF_B)
text(c, "CA, Managing Partner, ProfitPulse", right_x + 0.25 * IN, 1.78 * IN, 11.5, C_WHITE, font=SANS)
rect(c, right_x + 0.25 * IN, 2.18 * IN, right_w - 0.5 * IN, 1.5, C_TEAL)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
multiline(c, cred, right_x + 0.25 * IN, 2.35 * IN, 10.5, C_OFF_WHITE, font=SANS, gap=6, leading=13)

rect(c, right_x + 0.25 * IN, 3.95 * IN, right_w - 0.5 * IN, 1.5, C_TEAL)
contact = [
    {"text": "Profit-Pulse.com.au", "colour": C_OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "colour": C_TEAL},
    {"text": "+61 411 876 267", "colour": C_OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "colour": C_TEAL},
]
multiline(c, contact, right_x + 0.25 * IN, 4.12 * IN, 10.5, C_OFF_WHITE, font=SANS, gap=6, leading=13)

rect(c, right_x + 0.25 * IN, 5.15 * IN, right_w - 0.5 * IN, 1.5, C_TEAL)
wrapped(c, "Fractional CFO for growth stage SMEs: cash flow, investor ready reporting and growth strategy.",
        right_x + 0.25 * IN, 5.3 * IN, right_w - 0.5 * IN, 10, C_OFF_WHITE, font=SANS_I, leading=13)

c.showPage()
c.save()
print(f"PDF saved: {out_path}")
