"""
ProfitPulse Brief PDF Builder
Target: Sententia Consulting | Date for: 19 Jun 2026
Section 6 house style: white body, amber left bar, black header band.
Three slides. No dashes. No tier names in prospect-facing content.
Page: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch

W, H = 960, 540   # points

# ── Brand colours ─────────────────────────────────────────────────────────────
C_BLACK   = HexColor("#000000")
C_TEAL    = HexColor("#01A296")
C_AMB_B   = HexColor("#F8C806")
C_AMB_D   = HexColor("#F6A102")
C_GOLD    = HexColor("#E3A712")
C_WHITE   = HexColor("#FFFFFF")
C_OFF_WH  = HexColor("#E6E5DE")
C_MID_GY  = HexColor("#888888")
C_DRK_GY  = HexColor("#444444")
C_LT_GY   = HexColor("#CCCCCC")
C_TEAL_DK = HexColor("#041A18")

HLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def ry(screen_y):
    """Convert top-origin screen y to ReportLab bottom-origin y."""
    return H - screen_y


def fill(c, x, yt, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, ry(yt + h), w, h, fill=1, stroke=0)


def stroke(c, x, yt, w, h, sc, lw=1):
    c.setStrokeColor(sc)
    c.setLineWidth(lw)
    c.setFillColor(HexColor("#00000000"))
    c.rect(x, ry(yt + h), w, h, fill=0, stroke=1)


def hline(c, x, yt, w, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.line(x, ry(yt), x + w, ry(yt))


def single(c, text, x, yt, size, colour, font="Helvetica", align="left", max_w=None):
    """Draw one line of text. yt is top of the text in screen coords."""
    c.setFillColor(colour)
    c.setFont(font, size)
    base_y = ry(yt + size * 1.25)
    if align == "right" and max_w is not None:
        tw = c.stringWidth(text, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w is not None:
        tw = c.stringWidth(text, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, base_y, text)


def wrap_text(c, text, x, yt, max_w, size, colour, font="Helvetica", leading=None):
    """Draw wrapped text. Returns next yt after last line."""
    if leading is None:
        leading = size * 1.5
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
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
    y = yt
    for line in lines:
        c.drawString(x, ry(y + size * 1.25), line)
        y += leading
    return y


def link_rect(c, x, yt, w, h, url):
    """Add a clickable URI over a rectangle area."""
    c.linkURL(url, (x, ry(yt + h), x + w, ry(yt)))


def stat_card(c, x, yt, w, h, number, label1, label2, source):
    """Draw a stat card: black tile with teal top stripe."""
    STRIPE = 4
    fill(c, x, yt, w, STRIPE, C_TEAL)
    fill(c, x, yt + STRIPE, w, h - STRIPE, C_BLACK)
    PAD = 8
    single(c, number, x + PAD, yt + STRIPE + 2, 22, C_AMB_B, font="Helvetica-Bold")
    single(c, label1, x + PAD, yt + STRIPE + 30, 8.5, C_OFF_WH)
    single(c, label2, x + PAD, yt + STRIPE + 42, 8.5, C_OFF_WH)
    single(c, source,  x + PAD, yt + STRIPE + 55, 7,   C_MID_GY, font="Helvetica-Oblique")


# ── Output ────────────────────────────────────────────────────────────────────
OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_SententiaConsulting_19Jun2026.pdf"
c = canvas.Canvas(OUT, pagesize=(W, H))
c.setTitle("Sententia Consulting | ProfitPulse Brief | 19 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMB_D)   # amber left bar

# Header band
HDR_H = 63
fill(c, 7, 0, W - 7, HDR_H, C_BLACK)
single(c, "COMMERCIAL INTELLIGENCE BRIEF", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold")
single(c, "PROFITPULSE", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold",
       align="right", max_w=920)

# Company name
single(c, "Sententia Consulting", 20, 72, 34, C_BLACK, font="Helvetica-Bold")

# Descriptor
single(c, "Government and public sector management consultancy  |  Canberra ACT",
       20, 113, 11, C_TEAL)

# Stat cards
CARD_Y  = 130
CARD_H  = 88
cards = [
    ("Over $10M",  "Annual revenue",     "surpassed",          "AFR Fast 100 2025 #45"),
    ("60.20%",     "Three year CAGR",    "compound annual",    "AFR Fast 100 2025"),
    ("#45",        "AFR Fast 100",       "rank 2025",          "AFR Fast 100 2025"),
    ("2020",       "Year founded",       "two person start",   "Company history"),
    ("50+",        "Government and",     "sector clients",     "sententiaconsulting.com.au"),
    ("#28",        "AFR Fast Starters",  "rank 2022",          "AFR Fast Starters 2022"),
]
TOTAL_W = W - 40  # 20 left + 20 right margin
CARD_W  = int(TOTAL_W / 6) - 2
for i, (num, l1, l2, src) in enumerate(cards):
    cx = 20 + i * (CARD_W + 2)
    stat_card(c, cx, CARD_Y, CARD_W, CARD_H, num, l1, l2, src)

# KEY COMMERCIAL SIGNALS
SIG_LABEL_Y = 229
single(c, "KEY COMMERCIAL SIGNALS", 20, SIG_LABEL_Y, 10, C_TEAL, font="Helvetica-Bold")

signals = [
    ("Revenue tripled in under five years via 240% jump in government contract value",
     "Source: consultancy.com.au"),
    ("AFR Fast 100 2025 rank 45 with 60.20% CAGR, following AFR Fast Starters 2022 rank 28",
     "Source: AFR award citations"),
    ("National expansion: Melbourne office open, dedicated Managing Director appointed",
     "Source: consultancy.com.au"),
    ("Major recruitment drive including professionals from Protiviti and other firms",
     "Source: consultancy.com.au"),
    ("Active 2025 Graduate Program with applications open, sustained team scaling",
     "Source: sententiaconsulting.com.au"),
    ("50+ clients across federal, state, territory government and sector",
     "Source: sententiaconsulting.com.au"),
]

sy = SIG_LABEL_Y + 20
for sig, src in signals:
    fill(c, 20, sy + 3, 5, 5, C_TEAL)
    single(c, sig, 32, sy, 10.5, C_BLACK)
    single(c, src, 32, sy + 14, 8, C_MID_GY, font="Helvetica-Oblique")
    sy += 34

# Footer
hline(c, 20, 497, W - 40, C_LT_GY, 0.5)
single(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
          "Profit-Pulse.com.au",
       20, 503, 7.5, C_DRK_GY)
single(c, "19 Jun 2026", 20, 503, 7.5, C_DRK_GY, align="right", max_w=920)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMB_D)

fill(c, 7, 0, W - 7, HDR_H, C_BLACK)
single(c, "THE OPPORTUNITY", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold")
single(c, "SENTENTIA CONSULTING", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold",
       align="right", max_w=920)

single(c, "Sententia Consulting  |  Three commercial observations from ProfitPulse",
       20, 71, 11, C_BLACK)

# Three columns
COL_Y   = 90
COL_H   = 390
COL_W   = int((W - 20) / 3)  # 313 each approx
col_fills  = [C_TEAL, C_BLACK, C_GOLD]
txt_col    = [C_BLACK, C_WHITE, C_BLACK]
num_col    = [C_WHITE, C_AMB_B, C_BLACK]
hdr_col    = [C_WHITE, C_AMB_B, C_BLACK]

obs = [
    ("01",
     "Financial architecture has not kept pace with the growth",
     "Sententia has tripled revenue on government contract wins. "
     "The question a senior finance professional asks next is whether the business "
     "knows which contracts produce the best margin after full cost allocation. "
     "At $10 million in government consulting revenue, the difference between a "
     "high margin engagement and a breakeven one is rarely visible without a "
     "deliberate diagnostic. The Strategic Growth Diagnostic maps exactly that."),
    ("02",
     "Geographic expansion introduces new structural cost questions",
     "A Melbourne office with its own Managing Director creates real overhead "
     "and delivery complexity. Without a financial model for the Melbourne "
     "contribution alongside Canberra, expansion can dilute overall margin before "
     "it scales. Understanding how the national footprint changes the cost base, "
     "and planning the revenue ramp required to absorb it, is the work that "
     "belongs in the next financial planning cycle."),
    ("03",
     "Workforce utilisation is the lever that scales with headcount growth",
     "As Sententia grows its team through senior hires and a graduate cohort, "
     "revenue per person and billable utilisation become the efficiency measures "
     "that define profitability. In a professional services firm at this scale, "
     "a five percentage point improvement in utilisation across 30 people can "
     "move margin materially. The Workforce Capacity and Utilisation Review "
     "is the instrument for this."),
]

PAD = 10
for i, (num, hdr, body) in enumerate(obs):
    cx = 20 + i * COL_W
    fill(c, cx, COL_Y, COL_W - 2, COL_H, col_fills[i])

    single(c, num, cx + PAD, COL_Y + 12, 30, num_col[i], font="Helvetica-Bold")
    wrap_text(c, hdr, cx + PAD, COL_Y + 54, COL_W - PAD * 2 - 2,
              10.5, hdr_col[i], font="Helvetica-Bold", leading=15)
    wrap_text(c, body, cx + PAD, COL_Y + 120, COL_W - PAD * 2 - 2,
              9.5, txt_col[i], font="Helvetica", leading=14)

# Warm line
single(c, "These observations are offered in good faith. Sententia Consulting has "
          "built something genuinely impressive. The question is simply whether "
          "the financial architecture matches the ambition.",
       20, 492, 8.5, C_DRK_GY, font="Helvetica-Oblique")

hline(c, 20, 519, W - 40, C_LT_GY, 0.5)
single(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
       20, 524, 7.5, C_DRK_GY)
single(c, "19 Jun 2026", 20, 524, 7.5, C_DRK_GY, align="right", max_w=920)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3: THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMB_D)

fill(c, 7, 0, W - 7, HDR_H, C_BLACK)
single(c, "THE RECOMMENDATION", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold")
single(c, "SENTENTIA CONSULTING", 20, 18, 10, C_OFF_WH, font="Helvetica-Bold",
       align="right", max_w=920)

# ── LEFT COLUMN (20 to 620) ───────────────────────────────────────────────────
LX, LW = 20, 580

# Service name
single(c, "Strategic Growth Diagnostic", LX, 72, 24, C_BLACK, font="Helvetica-Bold")

# Price (no tier name)
single(c, "$5,000 one off", LX, 104, 14, C_TEAL, font="Helvetica-Bold")

# Description
wrap_text(c,
    "A six week engagement mapping revenue, capacity, and margin headroom, "
    "then producing a 12 month growth plan with funding and capital allocation "
    "steps across three scenarios. For Sententia, this clarifies which contracts "
    "earn the best margin and builds the financial model for the next growth phase.",
    LX, 126, LW, 11, C_BLACK, font="Helvetica", leading=16)

hline(c, LX, 198, LW, C_LT_GY, 0.5)

# Step one
single(c, "STEP ONE", LX, 207, 10, C_TEAL, font="Helvetica-Bold")
single(c, "Answer a few quick questions", LX, 227, 13, C_BLACK, font="Helvetica-Bold")
single(c, "See the solutions matched to your size and industry.",
       LX, 248, 11, C_BLACK)

CLEAN_URL = "https://profit-pulse.com.au/full-suite-of-products"
single(c, "profit-pulse.com.au/full-suite-of-products", LX, 268, 12, C_TEAL, font="Helvetica-Bold")
link_rect(c, LX, 262, LW, 22, CLEAN_URL)

# Purchase CTA
fill(c, LX, 300, LW, 35, C_TEAL)
single(c, "Purchase the suggested product now to get started",
       LX, 307, 12, C_WHITE, font="Helvetica-Bold",
       align="center", max_w=LW)
STRIPE_URL = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"
link_rect(c, LX, 300, LW, 35, STRIPE_URL)

hline(c, LX, 348, LW, C_LT_GY, 0.5)

single(c, "Prefer a conversation first?", LX, 356, 11, C_BLACK)
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
single(c, "Book a complimentary discovery call", LX, 376, 11, C_TEAL, font="Helvetica-Bold")
link_rect(c, LX, 370, LW, 22, BOOKING_URL)

# ── RIGHT COLUMN credibility panel (630 to 940) ───────────────────────────────
RX, RW = 630, 310
fill(c, RX, 68, RW, 430, C_BLACK)

single(c, "NITESH ROOPA", RX + 14, 76, 15, C_AMB_B, font="Helvetica-Bold")
single(c, "CA, Managing Partner", RX + 14, 100, 10.5, C_WHITE)
single(c, "ProfitPulse", RX + 14, 118, 14, C_TEAL, font="Helvetica-Bold")
hline(c, RX + 14, 142, RW - 28, C_TEAL, 1.5)

creds = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3B Cahora Bassa",
    "QLD GRBT project value: AUD 10B",
    "CA qualification: SAICA South Africa",
]
cy = 150
for line in creds:
    single(c, line, RX + 14, cy, 9.5, C_MID_GY)
    cy += 14

hline(c, RX + 14, cy + 4, RW - 28, HexColor("#333333"), 0.5)
cy += 12

contacts = [
    ("Profit-Pulse.com.au",          C_OFF_WH),
    ("Nitesh@Profit-Pulse.com.au",   C_TEAL),
    ("+61 411 876 267",              C_OFF_WH),
]
for ct, cc in contacts:
    single(c, ct, RX + 14, cy, 10.5, cc)
    cy += 15

cy += 4
single(c, "linkedin.com/in/nitesh-roopa-77594163", RX + 14, cy, 8.5, C_MID_GY)
cy += 14
single(c, "Brisbane, Australia", RX + 14, cy, 8.5, C_MID_GY)

# Footer
hline(c, 20, 519, W - 40, C_LT_GY, 0.5)
single(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
       20, 524, 7.5, C_DRK_GY)
single(c, "19 Jun 2026", 20, 524, 7.5, C_DRK_GY, align="right", max_w=920)

# ── Save ─────────────────────────────────────────────────────────────────────
c.save()
print(f"PDF saved: {OUT}")
