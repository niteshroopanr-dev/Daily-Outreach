"""
ProfitPulse Brief PDF Builder - Version 3
Target: Taskforce Australia | Date: 20 Jun 2026
Direct PDF generation, 3 slides, brand colours, white background, zero dashes.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.utils import simpleSplit

W, H = 960, 540

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFFWHITE = HexColor("#E6E5DE")
C_MID      = HexColor("#888888")
C_DARK     = HexColor("#444444")
C_PANEL    = HexColor("#F5F5F5")
C_LT       = HexColor("#CCCCCC")

DATE = "20 Jun 2026"
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/full-suite-of-products"
QUESTIONNAIRE_URL   = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_G3_COMMAND   = "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D"
BOOKING_URL         = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

OUT_PATH = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_20Jun2026.pdf"
c = canvas.Canvas(OUT_PATH, pagesize=(W, H))
c.setTitle("Taskforce Australia | ProfitPulse Brief | 20 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Product and Service Line Profitability")


def rl(screen_y):
    return H - screen_y


def rect(x, y_top, w, h, fill, stroke=None, sw=1):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
        c.rect(x, rl(y_top + h), w, h, fill=1, stroke=1)
    else:
        c.setStrokeColor(fill)
        c.rect(x, rl(y_top + h), w, h, fill=1, stroke=0)


def txt(text, x, y_top, size, colour, font="Helvetica", align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl(y_top + size * 0.85)
    if align == "right" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, baseline, text)


def wrapped(text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    if leading is None:
        leading = size * 1.5
    lines = simpleSplit(text, font, size, max_w)
    y = y_top
    c.setFillColor(colour)
    c.setFont(font, size)
    for line in lines:
        c.drawString(x, rl(y + size * 0.85), line)
        y += leading
    return y


def hline(x, y_top, w, colour, thickness=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, rl(y_top), x + w, rl(y_top))


def add_stripe_header(eyebrow, right_label):
    rect(0, 0, 8, H, C_AMBER_D)
    rect(8, 0, W - 8, 68, C_BLACK)
    txt(eyebrow, 20, 14, 12, C_OFFWHITE, font="Helvetica-Bold")
    txt(right_label, 756, 14, 11, C_TEAL, font="Helvetica-Bold", align="right", max_w=186)


def add_footer():
    hline(8, H - 20, W - 8, C_LT, 0.8)
    txt("Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
        20, H - 16, 8, C_DARK, font="Helvetica")
    txt(DATE, 720, H - 16, 8, C_DARK, font="Helvetica", align="right", max_w=222)


def stat_card(x, y_top, w, h, number, label1, label2, source):
    rect(x, y_top, w, h, C_BLACK)
    rect(x, y_top, w, 4, C_TEAL)
    txt(number, x + 8, y_top + 8, 24, C_WHITE, font="Helvetica-Bold")
    txt(label1,  x + 8, y_top + 38, 9, C_OFFWHITE, font="Helvetica")
    if label2:
        txt(label2, x + 8, y_top + 51, 8.5, HexColor("#999999"), font="Helvetica")
    txt(source,  x + 8, y_top + h - 13, 7, C_MID, font="Helvetica-Oblique")


def add_link(url, x, y_top, w, h):
    c.linkURL(url, (x, rl(y_top + h), x + w, rl(y_top)), relative=0)


# =============================================================================
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# =============================================================================
rect(0, 0, W, H, C_WHITE)
add_stripe_header("COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

txt("Taskforce Australia", 20, 76, 38, C_BLACK, font="Times-Bold")
txt("Property maintenance and compliance technology platform  |  Burnley, Melbourne VIC",
    20, 122, 10.5, C_TEAL, font="Helvetica")

# Stat cards (6 cards, 0..5)
CARDS = [
    ("$12.8M",  "Revenue FY2025",       "Smart50 award",          "Source: Smart50 2025"),
    ("31%",     "YoY growth FY2025",    "vs $9.7M in FY2024",     "Source: Smart50 2025"),
    ("19",      "Team members",         "Full time 2025",         "Source: Smart50 2025"),
    ("140K+",   "Jobs deployed",        "via RentSafe platform",  "Source: taskforce.com.au"),
    ("300+",    "Real estate offices",  "on platform 2024",       "Source: taskforce.com.au"),
    ("2014",    "Year founded",         "Burnley Melbourne VIC",  "Source: Smart50 2025"),
]
card_area_w = W - 20 - 8
card_gap = 6
card_w = (card_area_w - 5 * card_gap) / 6
card_h = 99
card_y = 138

for i, (num, l1, l2, src) in enumerate(CARDS):
    cx = 20 + i * (card_w + card_gap)
    stat_card(cx, card_y, card_w, card_h, num, l1, l2, src)

# Key commercial signals label
txt("KEY COMMERCIAL SIGNALS", 20, 252, 10, C_TEAL, font="Helvetica-Bold")

SIGNALS = [
    "Telstra Best of Business 2024, Victorian State Winner for Outstanding Growth  (telstra.com.au, eliteagent.com)",
    "Most Innovative Proptech, 2024 Proptech Awards Sydney  (industry press 2024)",
    "Smart50 rank improved each year: listed 2023, rank 46 in 2024, rank 37 in 2025  (SmartCompany Smart50 award citations)",
    "Doubled agency clients in 12 months to over 300 real estate offices  (Telstra Best of Business 2024 citation)",
    "RentSafe deployed across 20 consumer brands and 180 major real estate brands since launch 2021  (taskforce.com.au)",
]
sy = 269
for sig in SIGNALS:
    lines = simpleSplit(sig, "Helvetica", 9.5, 565)
    for line in lines:
        txt(line, 28, sy, 9.5, C_BLACK, font="Helvetica")
        sy += 14
    sy += 4

# Revenue chart
txt("REVENUE GROWTH (AUD MILLION)", 614, 252, 10, C_TEAL, font="Helvetica-Bold")
BAR_DATA = [(7.43, "FY2023", "Smart50 2023"), (9.7, "FY2024", "Smart50 2024"), (12.8, "FY2025", "Smart50 2025")]
BAR_X0   = 620
BAR_W    = 78
BAR_GAP  = 44
BAR_BASE = 484
BAR_AREA = 210

for i, (rev, yr, src) in enumerate(BAR_DATA):
    bx = BAR_X0 + i * (BAR_W + BAR_GAP)
    bh = int(BAR_AREA * rev / 12.8)
    by = BAR_BASE - bh
    rect(bx, by, BAR_W, bh, C_TEAL)
    txt(f"${rev}M", bx, by - 16, 9, C_BLACK, font="Helvetica-Bold", align="center", max_w=BAR_W)
    txt(yr,   bx, BAR_BASE + 5,  9,  C_DARK, font="Helvetica", align="center", max_w=BAR_W)
    txt(src,  bx, BAR_BASE + 19, 7,  C_MID,  font="Helvetica-Oblique", align="center", max_w=BAR_W)

hline(BAR_X0, BAR_BASE, BAR_W * 3 + BAR_GAP * 2, C_DARK, 1)

add_footer()


# =============================================================================
# PAGE 2: THE OPPORTUNITY
# =============================================================================
c.showPage()
rect(0, 0, W, H, C_WHITE)
add_stripe_header("THE OPPORTUNITY", "PROFITPULSE")

txt("Taskforce Australia   |   Three commercial observations from ProfitPulse",
    20, 78, 10.5, C_DARK, font="Helvetica")

COL_TOP = 100
COL_H   = 395
COL_GAP = 8
COL_W   = (W - 20 - 8 - 2 * COL_GAP) / 3

OBS = [
    (C_TEAL, C_BLACK,
     "01", "Three streams, one margin question",
     ("Taskforce runs three distinct service lines: RentSafe compliance checks, RentRepair maintenance, "
      "and consumer brand maintenance partnerships. At $12.8 million with 19 staff the business runs lean. "
      "But sustained growth across three streams raises a critical question: which stream delivers margin "
      "after platform costs, tradesperson coordination and compliance overhead are fully allocated? "
      "A product profitability engagement maps that in three weeks, ranking every revenue line by gross "
      "margin, contribution margin and operational drag, and giving the founders clarity on where the "
      "next hire, the next investment, and the next pitch should go.")),
    (C_BLACK, C_WHITE,
     "02", "Platform scale without a financial map creates risk",
     ("Three consecutive Smart50 appearances, each at a higher rank than the last, and a Telstra "
      "Outstanding Growth award in Victoria for 2024 confirm Taskforce as one of the fastest moving "
      "property tech businesses in Australia. But rapid platform expansion typically outpaces the "
      "financial architecture. With 300 plus real estate offices, 20 consumer brands and more than "
      "5,000 tradespeople on the network, the margin profile of each segment is almost certainly uneven. "
      "Scaling the least profitable activity hardest is a common outcome when the margin picture "
      "is not clear. ProfitPulse maps this in one focused three week engagement.")),
    (C_GOLD, C_BLACK,
     "03", "Where does the next dollar go at $12.8 million?",
     ("Revenue grew from $7.43 million in FY2023 to $12.8 million in FY2025, a 72 percent cumulative "
      "increase over two years verified across three independent Smart50 award citations. "
      "The momentum is real. The strategic question is capital allocation: which product line deserves "
      "the next hire, the next marketing dollar, the next technology investment? "
      "A Product and Service Line Profitability engagement produces a ranked view of every revenue line "
      "by gross margin, contribution margin and drag, giving the founders a defensible basis for every "
      "growth decision in the next 12 months.")),
]

for i, (fill, tc, idx, head, body) in enumerate(OBS):
    cx = 20 + i * (COL_W + COL_GAP)
    rect(cx, COL_TOP, COL_W, COL_H, fill)
    txt(idx,  cx + 12, COL_TOP + 12, 28, tc, font="Times-Bold")
    wrapped(head, cx + 12, COL_TOP + 50, COL_W - 24, 11,
            tc, font="Helvetica-Bold", leading=16)
    wrapped(body, cx + 12, COL_TOP + 116, COL_W - 24, 9.5,
            tc, font="Helvetica", leading=14.5)

wrapped(("These are observations offered in good faith. "
         "Taskforce has built something genuinely impressive. "
         "The question is simply whether the financial architecture is now keeping pace with the ambition."),
        20, 504, W - 28, 9, C_DARK, font="Helvetica-Oblique", leading=13)

add_footer()


# =============================================================================
# PAGE 3: THE RECOMMENDATION AND HOW TO START
# =============================================================================
c.showPage()
rect(0, 0, W, H, C_WHITE)
add_stripe_header("THE RECOMMENDATION", "PROFITPULSE")

# ── LEFT COLUMN ───────────────────────────────────────────────────────────────
LX = 20
LW = 550

txt("Product and Service Line Profitability",
    LX, 78, 17, C_BLACK, font="Times-Bold")

txt("$3,950 one off", LX, 104, 13, C_TEAL, font="Helvetica-Bold")

wrapped(("A three week project ranking every product or service line by gross margin, "
         "contribution margin and operational drag. "
         "Gives Taskforce Australia clarity on where to scale, where to fix and where to stop."),
        LX, 126, LW, 10.5, C_BLACK, font="Helvetica", leading=15)

hline(LX, 172, LW, C_TEAL, 1.5)

txt("STEP ONE", LX, 181, 10, C_TEAL, font="Helvetica-Bold")
txt("Answer a few quick questions", LX, 200, 12, C_BLACK, font="Helvetica-Bold")
txt("See the solutions matched to your size and industry.",
    LX, 220, 10, C_BLACK, font="Helvetica")

# Questionnaire address (clean text, hyperlinked to clean URL)
txt(QUESTIONNAIRE_CLEAN, LX, 238, 10, C_TEAL, font="Helvetica")
add_link(QUESTIONNAIRE_URL, LX, 232, LW, 18)

hline(LX, 264, LW, C_LT, 0.8)

# Purchase CTA (Stripe link behind words, address never visible)
txt("Purchase the suggested product now to get started",
    LX, 274, 11.5, C_AMBER_D, font="Helvetica-Bold")
add_link(STRIPE_G3_COMMAND, LX, 267, LW, 20)

txt("Product and Service Line Profitability   |   $3,950 one off",
    LX, 298, 9.5, C_DARK, font="Helvetica")

hline(LX, 320, LW, C_LT, 0.8)

txt("Prefer a conversation first?", LX, 330, 10, C_BLACK, font="Helvetica")
txt("Book a complimentary discovery call", LX, 349, 10.5, C_TEAL, font="Helvetica-Bold")
add_link(BOOKING_URL, LX, 342, LW, 20)
txt("bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/",
    LX, 370, 8.5, C_DARK, font="Helvetica")

hline(LX, 390, LW, C_LT, 0.6)
txt("Also identified: Strategic Growth Diagnostic   |   Customer Concentration and Profitability Map",
    LX, 399, 8.5, C_MID, font="Helvetica")

# ── RIGHT COLUMN: About Nitesh ────────────────────────────────────────────────
RX = 590
RW = 352
rect(RX - 8, 75, RW + 16, 436, C_PANEL)

txt("Nitesh Roopa", RX, 86, 18, C_BLACK, font="Times-Bold")
txt("CA, Managing Partner", RX, 114, 11, C_DARK, font="Helvetica")
txt("ProfitPulse", RX, 133, 13, C_TEAL, font="Helvetica-Bold")
hline(RX, 154, RW - 20, C_TEAL, 1.5)

CREDS = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed across career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa hydro, Mozambique",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "QIC 2023 to 2025: Finance and Commercial Lead,",
    "  AUD 10B Gympie Road Bypass Tunnel detailed business case",
    "Nedbank CIB 2015 to 2022: Energy Finance,",
    "  Principal and Equity Finance roles",
    "PwC South Africa 2010 to 2014: CA traineeship,",
    "  Top 40 listed clients including SABMiller and Sun International",
]
cy_r = 164
for line in CREDS:
    txt(line, RX, cy_r, 9, C_DARK, font="Helvetica")
    cy_r += 13

hline(RX, cy_r + 4, RW - 20, C_LT, 0.8)
cy_r += 16

CONTACTS = [
    ("Profit-Pulse.com.au",                   C_DARK),
    ("Nitesh@Profit-Pulse.com.au",             C_TEAL),
    ("+61 411 876 267",                        C_DARK),
    ("linkedin.com/in/nitesh-roopa-77594163",  C_DARK),
]
for ctext, ccol in CONTACTS:
    txt(ctext, RX, cy_r, 9.5, ccol, font="Helvetica")
    cy_r += 15

add_footer()

# =============================================================================
c.save()
print(f"PDF saved: {OUT_PATH}")
