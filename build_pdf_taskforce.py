"""
ProfitPulse Brief PDF Builder V3
Target: Taskforce Australia | Date: 01 Jul 2026
White-background V3 house style. Brand colours only. Zero dashes.
Page: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

W = 960
H = 540

# Brand colours (Rule 3)
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")


def rl(y):
    """Convert top-origin screen-y to ReportLab bottom-origin y."""
    return H - y


def fill(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, scol, lw=1):
    c.setStrokeColor(scol)
    c.setLineWidth(lw)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, rl(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold",
        align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl(y_top + size)
    if align == "right" and max_width is not None:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width is not None:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline, text)


def txt_wrap(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    """Draw wrapped text. Returns final y_top after last line."""
    if leading is None:
        leading = size * 1.48
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
    lines = []
    cur = ""
    for w_ in words:
        test = (cur + " " + w_).strip()
        if c.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    y = y_top
    for ln in lines:
        c.drawString(x, rl(y + size), ln)
        y += leading
    return y


def hline(c, x, y_top, w, colour, lw=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.line(x, rl(y_top), x + w, rl(y_top))


def link(c, url, x, y_top, w, h):
    """Add a clickable hyperlink rectangle."""
    c.linkURL(url, (x, rl(y_top + h), x + w, rl(y_top)), relative=0)


def footer_slide(c, date="01 Jul 2026"):
    hline(c, 7, 504, 946, C_BLACK, 1)
    txt(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
        18, 512, 9, C_BLACK, font="Helvetica")
    txt(c, date, 720, 512, 9, C_BLACK, font="Helvetica",
        align="right", max_width=232)


# ══════════════════════════════════════════════════════════════════
out = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Jul2026.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
c.setTitle("Taskforce Australia | ProfitPulse Brief | 01 Jul 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")


# ════════════════════════════════════════════════════════════════
# SLIDE 1  COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_B)
fill(c, 7, 0, 953, 68, C_BLACK)

txt(c, "COMMERCIAL INTELLIGENCE BRIEF", 20, 16, 11, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "PROFITPULSE", 720, 16, 11, C_OFF_WHITE, font="Helvetica-Bold",
    align="right", max_width=232)

# Company name (Times-Roman approximates Cambria/Georgia serif)
txt(c, "Taskforce Australia", 20, 74, 42, C_BLACK, font="Times-BoldItalic")

# Descriptor
txt(c, "Technology enabled property maintenance and compliance platform"
       "     Burnley, Melbourne VIC",
    20, 124, 11, C_TEAL, font="Helvetica-Oblique")

# Teal rule
hline(c, 13, 142, 940, C_TEAL, 2)

# ── Stat cards ───────────────────────────────────────────────────
CARDS = [
    ("$12.8M",  "Revenue", "FY2025",      "Smart50 2025, SmartCompany"),
    ("31%",     "Growth Rate", "Smart50 2025", "SmartCompany Smart50 2025"),
    ("19",      "Team", "Members",         "Smart50 2025, SmartCompany"),
    ("5,500",   "Tradie", "Network",       "Smart50 2025, SmartCompany"),
    ("Rank 37", "Smart50", "2025",         "SmartCompany Smart50 2025"),
]
cw   = 183
ch   = 97
cgap = 6
cx_  = 7
cy_  = 148

for big, lbl1, lbl2, src in CARDS:
    fill(c, cx_, cy_, cw, ch, C_BLACK)
    fill(c, cx_, cy_, cw, 3, C_TEAL)
    txt(c, big, cx_ + 8, cy_ + 10, 24, C_AMBER_B, font="Times-BoldItalic")
    txt(c, lbl1, cx_ + 8, cy_ + 56, 9.5, C_OFF_WHITE, font="Helvetica")
    txt(c, lbl2, cx_ + 8, cy_ + 68, 9.5, C_OFF_WHITE, font="Helvetica")
    txt(c, src,  cx_ + 8, cy_ + 82, 7, C_OFF_WHITE, font="Helvetica-Oblique")
    cx_ += cw + cgap

# ── Key Commercial Signals ────────────────────────────────────
SIGNALS = [
    "Smart50 consecutive listings: 2023, rank 46 in 2024, rank 37 in 2025  |  Source: SmartCompany Smart50 2023/2024/2025",
    "Revenue grew from $9.7M to $12.8M year on year (31 per cent)  |  Source: SmartCompany Smart50 2024 and 2025",
    "RentRepair subscription launched 2023, from $59 per month  |  Source: Taskforce Australia website",
    "HousingSolutions: preferred supplier to community housing organisations  |  Source: Smart50 2025 and Taskforce website",
    "Real+ partnership announced March 2025  |  Source: Taskforce Australia website (March 2025)",
    "Telstra Best of Business Awards Victorian State Winner  |  Source: Elite Agent publication",
]
sig_x = 18
txt(c, "KEY COMMERCIAL SIGNALS", sig_x, 258, 10, C_TEAL, font="Helvetica-Bold")
sy = 272
for sig in SIGNALS:
    txt_wrap(c, sig, sig_x, sy, 560, 8.5, C_BLACK, font="Helvetica", leading=13)
    sy += 14

# ── Revenue Bar Chart ─────────────────────────────────────────
chart_x = 605
txt(c, "Revenue Growth", chart_x, 258, 10, C_TEAL, font="Helvetica-Bold")

baseline = 468
max_h    = 190
bw_      = 108

bh24 = int((9.7 / 12.8) * max_h)
bx24 = chart_x + 10
fill(c, bx24, baseline - bh24, bw_, bh24, C_AMBER_D)
txt(c, "$9.7M",  bx24, baseline - bh24 - 18, 9, C_BLACK, font="Helvetica-Bold",
    align="center", max_width=bw_)
txt(c, "FY2024", bx24, baseline + 5,          8, C_BLACK, font="Helvetica",
    align="center", max_width=bw_)

bh25 = max_h
bx25 = bx24 + bw_ + 28
fill(c, bx25, baseline - bh25, bw_, bh25, C_TEAL)
txt(c, "$12.8M", bx25, baseline - bh25 - 18, 9, C_BLACK, font="Helvetica-Bold",
    align="center", max_width=bw_)
txt(c, "FY2025", bx25, baseline + 5,          8, C_BLACK, font="Helvetica",
    align="center", max_width=bw_)

hline(c, chart_x + 5, baseline, 350, C_BLACK, 1.5)
txt(c, "Source: SmartCompany Smart50 2024 and 2025 award citations",
    chart_x, baseline + 22, 7.5, C_BLACK, font="Helvetica-Oblique")

footer_slide(c)


# ════════════════════════════════════════════════════════════════
# SLIDE 2  THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_B)
fill(c, 7, 0, 953, 68, C_BLACK)
txt(c, "THE OPPORTUNITY", 20, 16, 11, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "PROFITPULSE", 720, 16, 11, C_OFF_WHITE, font="Helvetica-Bold",
    align="right", max_width=232)
txt(c, "Taskforce Australia   |   Three commercial observations from ProfitPulse",
    20, 76, 11, C_BLACK, font="Helvetica")

# Three observation columns
COL_TOP  = 98
COL_BOT  = 485
COL_H    = COL_BOT - COL_TOP
COL_W    = int((960 - 7) / 3)   # 317 each

OBS_DATA = [
    {
        "fill": C_TEAL,
        "txt":  C_WHITE,
        "idx":  "01",
        "head": "Three products, one financial architecture",
        "body": (
            "Taskforce Australia now operates three distinct revenue streams: "
            "compliance checks on a transactional model, the RentRepair subscription "
            "at $59 per month and above (launched 2023), and HousingSolutions on "
            "contract. Each carries a different margin profile, a different customer "
            "acquisition cost, and a different claim on the team of nineteen managing "
            "$12.8 million in revenue. Allocating that team across three concurrent "
            "streams without a clear view of which generates the best return per hour "
            "of capacity is the central financial risk of the next growth phase. "
            "ProfitPulse's Strategic Growth Diagnostic maps exactly this in six weeks."
        ),
    },
    {
        "fill": C_BLACK,
        "txt":  C_OFF_WHITE,
        "idx":  "02",
        "head": "The software investment needs an ROI number first",
        "body": (
            "The publicly stated plan to expand in house software capabilities is a "
            "capital decision that compounds. Every dollar committed before the revenue "
            "base can support it is unavailable for scaling RentRepair subscriber "
            "acquisition or winning the next social housing contract. At $12.8 million "
            "revenue with nineteen employees, the sequencing question is a financial "
            "modelling exercise. Three scenario modelling of the software first versus "
            "growth first paths produces the cash flow and margin comparison needed to "
            "commit with confidence. ProfitPulse builds that model as part of the "
            "Strategic Growth Diagnostic."
        ),
    },
    {
        "fill": C_GOLD,
        "txt":  C_BLACK,
        "idx":  "03",
        "head": "Subscription revenue carries a higher value multiple",
        "body": (
            "The shift to RentRepair is more than a product decision. Recurring "
            "subscription revenue at scale commands a substantially higher enterprise "
            "value multiple than the same dollar of transactional revenue. If "
            "RentRepair grows to represent a meaningful share of total revenue, "
            "the business carries a very different value story for any future investor "
            "or acquirer. The financial reporting structure needs to reflect this "
            "division now, so the recurring versus transactional split is visible and "
            "defensible when the founders choose to test the market or invite "
            "institutional interest."
        ),
    },
]

for i, obs in enumerate(OBS_DATA):
    col_x = 7 + COL_W * i
    fill(c, col_x, COL_TOP, COL_W, COL_H, obs["fill"])

    # Index
    txt(c, obs["idx"], col_x + 12, COL_TOP + 12, 34, obs["txt"], font="Times-BoldItalic")

    # Rule under index
    rule_col = C_WHITE if obs["fill"] != C_WHITE else C_TEAL
    hline(c, col_x + 12, COL_TOP + 60, COL_W - 24, rule_col, 1)

    # Observation header
    txt_wrap(c, obs["head"],
             col_x + 12, COL_TOP + 66,
             COL_W - 24, 12, obs["txt"],
             font="Helvetica-Bold", leading=16)

    # Body text
    txt_wrap(c, obs["body"],
             col_x + 12, COL_TOP + 112,
             COL_W - 24, 9.5, obs["txt"],
             font="Helvetica", leading=13.5)

# Warm closing line
txt_wrap(c,
         "These are observations offered in good faith. Taskforce Australia has "
         "built something genuinely impressive. The question is simply whether "
         "the financial architecture now matches the ambition.",
         18, COL_BOT + 8, 920, 10, C_BLACK,
         font="Helvetica-Oblique", leading=14)

footer_slide(c)


# ════════════════════════════════════════════════════════════════
# SLIDE 3  THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
fill(c, 0, 0, 7, H, C_AMBER_B)
fill(c, 7, 0, 953, 68, C_BLACK)
txt(c, "THE RECOMMENDATION", 20, 16, 11, C_OFF_WHITE, font="Helvetica-Bold")
txt(c, "PROFITPULSE", 720, 16, 11, C_OFF_WHITE, font="Helvetica-Bold",
    align="right", max_width=232)

# ── LEFT COLUMN ───────────────────────────────────────────────
LX  = 20
LW  = 575
ly  = 80

# Service name
txt(c, "Strategic Growth Diagnostic", LX, ly, 28, C_BLACK, font="Times-BoldItalic")
ly += 36

# Price (no tier name per Rule 6)
txt(c, "$5,000 one off   ·   ProfitPulse verified price",
    LX, ly, 12, C_TEAL, font="Helvetica-Bold")
ly += 20

# Description
ly = txt_wrap(c,
    "A six week engagement mapping revenue, capacity, and margin headroom "
    "across all three revenue streams, producing a twelve month growth plan "
    "with capital allocation steps and three scenario modelling. Designed for "
    "Taskforce's current moment: three concurrent products, a software "
    "investment decision pending, and a team of nineteen managing $12.8M.",
    LX, ly + 6, LW, 10, C_BLACK, font="Helvetica", leading=14)
ly += 10

hline(c, LX, ly, LW, C_TEAL, 1)
ly += 10

# Step one block
txt(c, "Step one: answer a few quick questions", LX, ly, 12, C_BLACK, font="Helvetica-Bold")
ly += 18
txt(c, "See the solutions matched to your size and industry.",
    LX, ly, 10, C_BLACK, font="Helvetica")
ly += 16

# Clean URL (no UTM on brief, hyperlinked to clean address)
txt(c, "profit-pulse.com.au/full-suite-of-products",
    LX, ly, 11, C_TEAL, font="Helvetica-Bold")
url_w = c.stringWidth("profit-pulse.com.au/full-suite-of-products", "Helvetica-Bold", 11)
link(c, "https://profit-pulse.com.au/full-suite-of-products",
     LX, ly, url_w, 12)
ly += 20

hline(c, LX, ly, LW, C_BLACK, 0.75)
ly += 10

# Direct CTA (Stripe link behind text, address never shown)
txt(c, "Purchase the suggested product now to get started",
    LX, ly, 12, C_AMBER_D, font="Helvetica-Bold")
cta_w = c.stringWidth("Purchase the suggested product now to get started",
                       "Helvetica-Bold", 12)
link(c, "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
     LX, ly, cta_w, 14)
ly += 22

hline(c, LX, ly, LW, C_BLACK, 0.75)
ly += 10

# Booking link
txt(c, "Prefer a conversation first?", LX, ly, 10, C_BLACK, font="Helvetica")
ly += 15
txt(c, "Book a complimentary discovery call",
    LX, ly, 11, C_TEAL, font="Helvetica-Bold")
bk_w = c.stringWidth("Book a complimentary discovery call", "Helvetica-Bold", 11)
link(c,
     "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/"
     "?ismsaljsauthenabled=true",
     LX, ly, bk_w, 13)

# ── RIGHT COLUMN: Credibility ─────────────────────────────────
RX  = 614
RW  = 330
fill(c, RX, 76, RW, 418, C_BLACK)

ry = 84
txt(c, "Nitesh Roopa", RX + 14, ry, 20, C_AMBER_B, font="Times-BoldItalic")
ry += 28
txt(c, "CA, Managing Partner", RX + 14, ry, 11, C_WHITE, font="Helvetica")
ry += 16
txt(c, "ProfitPulse", RX + 14, ry, 14, C_TEAL, font="Helvetica-Bold")
ry += 22
hline(c, RX + 14, ry, RW - 28, C_TEAL, 1.5)
ry += 12

CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3 billion, Cahora Bassa Hydro",
    "Total GRBT project value: over AUD 10 billion",
]
for crd in CREDS:
    txt(c, crd, RX + 14, ry, 9.5, C_OFF_WHITE, font="Helvetica")
    ry += 14

ry += 6
hline(c, RX + 14, ry, RW - 28, C_TEAL, 1)
ry += 12

CONTACT = [
    ("Profit-Pulse.com.au",                   C_OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au",            C_TEAL),
    ("+61 411 876 267",                       C_OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", C_OFF_WHITE),
]
for ctxt, ccol in CONTACT:
    txt(c, ctxt, RX + 14, ry, 10, ccol, font="Helvetica")
    ry += 15

footer_slide(c)

c.save()
print(f"PDF saved: {out}")
