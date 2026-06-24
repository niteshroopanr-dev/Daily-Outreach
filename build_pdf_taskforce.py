"""
ProfitPulse Brief PDF Builder
Target: Taskforce Australia | Date: 25 Jun 2026
White background house style. Brand colours only. Zero dashes.
Page: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor
from reportlab.lib.units import inch

W, H = 960, 540

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_W    = HexColor("#E6E5DE")
C_MID_GRY  = HexColor("#888888")
C_DRK_GRY  = HexColor("#444444")
C_LT_GRY   = HexColor("#666666")
C_TEAL_LT  = HexColor("#F0FBFA")
C_NEAR_BLK = HexColor("#111111")


def ry(y_top):
    return H - y_top


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, ry(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, stroke_colour, lw=1):
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(lw)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, ry(y_top + h), w, h, fill=0, stroke=1)


def fill_stroke_rect(c, x, y_top, w, h, fill_c, stroke_c, lw=1):
    c.setFillColor(fill_c)
    c.setStrokeColor(stroke_c)
    c.setLineWidth(lw)
    c.rect(x, ry(y_top + h), w, h, fill=1, stroke=1)


def hline(c, x, y_top, w, colour, thickness=1):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, ry(y_top), x + w, ry(y_top))


def t(c, text, x, y_top, size, colour, font="Helvetica", align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = ry(y_top + size)
    if align == "right" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, baseline, text)


def tw(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
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
        c.drawString(x, ry(y + size), line)
        y += leading
    return y


def link(c, text, x, y_top, size, colour, url, font="Helvetica"):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = ry(y_top + size)
    tw_val = c.stringWidth(text, font, size)
    c.drawString(x, baseline, text)
    c.linkURL(url, (x, baseline - 2, x + tw_val, baseline + size + 2), relative=0)


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_25Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Taskforce Australia | ProfitPulse Brief | 25 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")


# =============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# =============================================================================
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)         # left amber stripe
fill_rect(c, 7, 0, W - 7, 72, C_BLACK)       # header band

t(c, "COMMERCIAL INTELLIGENCE BRIEF", 18, 16, 11, C_OFF_W, font="Helvetica-Bold")
t(c, "TASKFORCE AUSTRALIA", 560, 16, 11, C_TEAL, font="Helvetica-Bold",
  align="right", max_w=383)
t(c, "25 Jun 2026", 18, 46, 9, C_MID_GRY)

t(c, "Taskforce Australia", 18, 82, 38, C_BLACK, font="Helvetica-Bold")
t(c, "Technology enabled property maintenance and compliance platform, Burnley, Melbourne VIC",
  18, 132, 12, C_TEAL, font="Helvetica")
hline(c, 18, 150, 924, C_TEAL, 1.5)

# ── STAT CARDS ────────────────────────────────────────────────────────────────
CARDS = [
    ("$12.8M", "Revenue  FY2025",     "Smart50 Nov 2025"),
    ("31%",    "Year on year growth", "Smart50 Nov 2025"),
    ("19",     "Full time team",      "Smart50 Nov 2025"),
    ("5,500",  "Tradespeople",        "Smart50 Nov 2025"),
    ("#37",    "Smart50 2025 rank",   "Smart50 Nov 2025"),
    ("500",    "Agency clients",      "taskforce.com.au"),
]
CARD_W, CARD_H, CARD_TOP = 147, 100, 157
CARD_GAP = 4
cx = 18
for num, lbl, src in CARDS:
    fill_rect(c, cx, CARD_TOP, CARD_W, CARD_H, C_BLACK)
    fill_rect(c, cx, CARD_TOP, CARD_W, 5, C_TEAL)
    t(c, num, cx + 8, CARD_TOP + 12, 22, C_AMBER_B, font="Helvetica-Bold")
    t(c, lbl, cx + 8, CARD_TOP + 64, 9, C_OFF_W)
    t(c, src, cx + 8, CARD_TOP + 80, 7, C_TEAL, font="Helvetica-Oblique")
    cx += CARD_W + CARD_GAP

# ── KEY COMMERCIAL SIGNALS ────────────────────────────────────────────────────
t(c, "KEY COMMERCIAL SIGNALS", 18, 268, 10, C_TEAL, font="Helvetica-Bold")
hline(c, 18, 282, 535, C_TEAL, 1)

SIGNALS = [
    ("Three consecutive Smart50 placements (2023, 2024, 2025), national ranking improving each year.",
     "SmartCompany Smart50 series 2023 to 2025."),
    ("Revenue grew from $7.43M to $12.8M in two years, a cumulative 72 percent increase.",
     "SmartCompany Smart50 2023 and Smart50 2025."),
    ("Housing division publicly named strongest and most profitable; social housing expansion underway.",
     "SmartCompany Smart50 2025."),
    ("Founders publicly committed to 50 to 60 percent CAGR next three years via organically funded growth.",
     "SmartCompany Smart50 2025."),
    ("Four point growth agenda: AI platform investment and staff equity plan create active capital decisions.",
     "SmartCompany Smart50 2025."),
    ("RentSafe platform serves 500 real estate agencies with automated property safety compliance.",
     "taskforce.com.au and PropTechPRO company profile."),
]
sy = 290
for sig_text, sig_src in SIGNALS:
    tw(c, sig_text, 18, sy, 535, 9, C_BLACK, font="Helvetica", leading=13)
    t(c, "Source: " + sig_src, 18, sy + 13, 7, C_MID_GRY, font="Helvetica-Oblique")
    sy += 36

# ── REVENUE GROWTH CHART ──────────────────────────────────────────────────────
t(c, "REVENUE GROWTH  (AUD)", 564, 268, 10, C_TEAL, font="Helvetica-Bold")
hline(c, 564, 282, 378, C_TEAL, 1)

CHART_BASE = 490
CHART_MAX_H = 198
BAR_W = 88
BAR_GAP = 28
BARS = [("FY2023", 7.43, C_GOLD), ("FY2024", 9.7, C_AMBER_D), ("FY2025", 12.8, C_TEAL)]
bx = 582
for label, val, colour in BARS:
    bh = int(val / 12.8 * CHART_MAX_H)
    fill_rect(c, bx, CHART_BASE - bh, BAR_W, bh, colour)
    t(c, f"${val}M", bx, CHART_BASE - bh - 16, 9, C_BLACK,
      font="Helvetica-Bold", align="center", max_w=BAR_W)
    t(c, label, bx, CHART_BASE + 5, 9, C_DRK_GRY,
      align="center", max_w=BAR_W)
    bx += BAR_W + BAR_GAP

t(c, "Source: SmartCompany Smart50 award citations 2023, 2024, and 2025",
  564, 508, 7, C_TEAL, font="Helvetica-Oblique")

# Footer
hline(c, 7, 524, W - 7, C_TEAL, 1)
t(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
  18, 528, 8, C_LT_GRY)
t(c, "25 Jun 2026", 720, 528, 8, C_LT_GRY, align="right", max_w=222)


# =============================================================================
# SLIDE 2: THE OPPORTUNITY
# =============================================================================
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
fill_rect(c, 7, 0, W - 7, 72, C_BLACK)

t(c, "THE OPPORTUNITY", 18, 16, 11, C_OFF_W, font="Helvetica-Bold")
t(c, "TASKFORCE AUSTRALIA", 560, 16, 11, C_TEAL, font="Helvetica-Bold",
  align="right", max_w=383)
t(c, "Three commercial observations from ProfitPulse", 18, 82, 13, C_DRK_GRY)

OBS = [
    {
        "idx": "01",
        "header": "Exceptional operating leverage sets the foundation",
        "body": (
            "Taskforce generates $12.8M in annual revenue with 19 full time staff, producing "
            "approximately $674K in revenue per employee. This figure is verifiable across three "
            "consecutive Smart50 award citations. A business operating at this ratio has a clearly "
            "productive core. The next question is not whether the model works but what the financial "
            "architecture should look like to scale it intentionally and without overstretching the "
            "team or the balance sheet."
        ),
    },
    {
        "idx": "02",
        "header": "Four growth vectors competing for the same capital pool",
        "body": (
            "The four point growth agenda covers social housing expansion, a staff equity plan, "
            "AI platform development, and organic growth in the core compliance business. All four "
            "represent genuine opportunities backed by strong market tailwinds in rental compliance "
            "and community housing. Without a structured capital allocation model and a scenario tested "
            "12 month plan, the risk is that resources are spread thinly across all four vectors "
            "rather than sequenced against the highest return lever at each stage."
        ),
    },
    {
        "idx": "03",
        "header": "A 50% CAGR ambition needs a financial plan to match the ambition",
        "body": (
            "Jason Bright has publicly committed to a 50 to 60 percent CAGR over the next three "
            "years, implying revenue approaching $30M by FY2028 from a base of $12.8M. That trajectory "
            "requires a clear view of margin headroom, capacity constraints, and funding requirements "
            "before capital is deployed. A costed 12 month growth plan with three scenario models "
            "converts that ambition into a fundable and executable path. This is the specific "
            "capability the Strategic Growth Diagnostic delivers."
        ),
    },
]

COL_FILLS = [C_TEAL, C_BLACK, C_GOLD]
IDX_COLS  = [C_WHITE, C_AMBER_B, C_BLACK]
HDR_COLS  = [C_WHITE, C_AMBER_B, C_BLACK]
BOD_COLS  = [C_WHITE, C_OFF_W, C_BLACK]

COL_TOP = 104
COL_H   = 384
COL_W   = 295
COL_GAP = 16

cx = 18
for i, obs in enumerate(OBS):
    fill_rect(c, cx, COL_TOP, COL_W, COL_H, COL_FILLS[i])
    t(c, obs["idx"], cx + 12, COL_TOP + 12, 26, IDX_COLS[i], font="Helvetica-Bold")
    tw(c, obs["header"], cx + 12, COL_TOP + 56, COL_W - 24, 11, HDR_COLS[i],
       font="Helvetica-Bold", leading=14)
    tw(c, obs["body"], cx + 12, COL_TOP + 102, COL_W - 24, 9.5, BOD_COLS[i],
       font="Helvetica", leading=13)
    cx += COL_W + COL_GAP

tw(c,
   ("These observations are offered in good faith. Taskforce Australia has built something "
    "genuinely impressive. The question is simply whether the financial architecture matches the ambition."),
   18, 498, 924, 8.5, C_LT_GRY, font="Helvetica-Oblique", leading=12)

hline(c, 7, 524, W - 7, C_TEAL, 1)
t(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
  18, 528, 8, C_LT_GRY)
t(c, "25 Jun 2026", 720, 528, 8, C_LT_GRY, align="right", max_w=222)


# =============================================================================
# SLIDE 3: THE RECOMMENDATION
# =============================================================================
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
fill_rect(c, 7, 0, W - 7, 72, C_BLACK)

t(c, "THE RECOMMENDATION", 18, 16, 11, C_OFF_W, font="Helvetica-Bold")
t(c, "TASKFORCE AUSTRALIA", 560, 16, 11, C_TEAL, font="Helvetica-Bold",
  align="right", max_w=383)

# ── LEFT COLUMN ───────────────────────────────────────────────────────────────
lx, lw = 18, 525

t(c, "Strategic Growth Diagnostic", lx, 82, 24, C_BLACK, font="Helvetica-Bold")
t(c, "$5,000 one off", lx, 118, 15, C_TEAL, font="Helvetica-Bold")

tw(c,
   ("A six week engagement that maps your revenue, capacity, and margin headroom, then produces "
    "a 12 month growth plan with funding and capital allocation steps spelled out. "
    "Three scenario modelling. Designed for a business with a clear growth target "
    "that needs the financial architecture to support it."),
   lx, 140, lw, 10.5, HexColor("#333333"), font="Helvetica", leading=14)

hline(c, lx, 202, lw, C_TEAL, 1)

# Step 1 box
fill_stroke_rect(c, lx, 208, lw, 106, C_TEAL_LT, C_TEAL, 1)
t(c, "Step one: answer a few quick questions", lx + 10, 218, 13, C_TEAL, font="Helvetica-Bold")
t(c, "See the solutions matched to your size and industry.", lx + 10, 248, 10.5, C_BLACK)
link(c, "profit-pulse.com.au/full-suite-of-products",
     lx + 10, 272, 10.5, C_TEAL,
     "https://profit-pulse.com.au/full-suite-of-products",
     font="Helvetica-Bold")

# Purchase CTA
fill_rect(c, lx, 322, lw, 37, C_AMBER_D)
link(c, "Purchase the suggested product now to get started",
     lx + 10, 329, 11, C_BLACK,
     "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
     font="Helvetica-Bold")

# Booking
t(c, "Prefer a conversation first?", lx, 372, 10.5, C_DRK_GRY)
link(c, "Book a complimentary discovery call",
     lx + 228, 372, 10.5, C_TEAL,
     "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# ── RIGHT COLUMN (credibility) ────────────────────────────────────────────────
rx, rw2 = 564, 378
fill_rect(c, rx, 80, rw2, 408, C_BLACK)

t(c, "Nitesh Roopa", rx + 14, 90, 20, C_AMBER_B, font="Helvetica-Bold")
t(c, "CA, Managing Partner", rx + 14, 120, 12, C_WHITE)
t(c, "ProfitPulse", rx + 14, 142, 18, C_TEAL, font="Helvetica-Bold")
hline(c, rx + 14, 170, rw2 - 28, C_TEAL, 1.5)

CREDS = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed across career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa Hydro",
    "Total GRBT Queensland project value: over AUD 10 billion",
]
cy = 182
for cline in CREDS:
    t(c, cline, rx + 14, cy, 9.5, C_MID_GRY)
    cy += 16

hline(c, rx + 14, cy + 6, rw2 - 28, HexColor("#333333"), 1)
cy += 18

CONTACTS = [
    ("Profit-Pulse.com.au",                     C_OFF_W),
    ("Nitesh@Profit-Pulse.com.au",              C_TEAL),
    ("+61 411 876 267",                         C_OFF_W),
    ("linkedin.com/in/nitesh-roopa-77594163",   C_MID_GRY),
]
for ct, cc in CONTACTS:
    t(c, ct, rx + 14, cy, 9.5, cc)
    cy += 15

hline(c, 7, 524, W - 7, C_TEAL, 1)
t(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
  18, 528, 8, C_LT_GRY)
t(c, "25 Jun 2026", 720, 528, 8, C_LT_GRY, align="right", max_w=222)

c.save()
print(f"PDF saved: {out_path}")
