"""
ProfitPulse Brief PDF Builder
Target: AdUnion | Date: 09 Jun 2026
Direct PDF generation, 3 slides, brand colours only, zero dashes.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.units import inch
pt = 1  # 1 point = 1 unit in ReportLab
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ── Page dimensions ─────────────────────────────────────────────────────────
W = 960  # pt  (13.333 in * 72)
H = 540  # pt  (7.5   in * 72)

# ── Brand colours ────────────────────────────────────────────────────────────
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_DARK_PNL  = HexColor("#111111")
C_TEAL_DRK  = HexColor("#041A18")
C_MID_GREY  = HexColor("#888888")
C_DRK_GREY  = HexColor("#444444")
C_PANEL2    = HexColor("#1A1A1A")
C_AMBER_BG  = HexColor("#1A1500")

# ── Helpers ──────────────────────────────────────────────────────────────────
def rl_y(screen_y):
    """Convert top-origin y to ReportLab bottom-origin y."""
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    """Fill a rectangle. y_top is top-edge in screen coords."""
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, stroke_colour, line_width=1):
    """Draw a stroked rectangle outline. y_top is top-edge in screen coords."""
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def fill_stroke_rect(c, x, y_top, w, h, fill_colour, stroke_colour, line_width=1):
    c.setFillColor(fill_colour)
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold",
        align="left", max_width=None):
    """Draw a single line of text. y_top is baseline's top position in screen coords."""
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline_y = rl_y(y_top + size)  # approximate baseline
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline_y, text)


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica",
                leading=None, align="left"):
    """
    Draw text wrapped to max_w pt. Returns the y position after the last line.
    y_top is in screen coords.
    """
    if leading is None:
        leading = size * 1.45
    c.setFillColor(colour)
    c.setFont(font, size)

    words = text.split()
    lines = []
    current = ""
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
    return y  # next available y (screen coords)


def hline(c, x, y_top, w, colour, thickness=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def vline(c, x, y_top, h, colour, thickness=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, rl_y(y_top), x, rl_y(y_top + h))


# ── Output path ──────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_AdUnion_09Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("AdUnion | ProfitPulse Brief | 09 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Customer Concentration and Profitability Map")


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE OPENING
# ════════════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_BLACK)           # full black background
fill_rect(c, 0, 0, 5, H, C_TEAL)            # left accent bar 5pt wide

# ProfitPulse brand
txt(c, "PROFITPULSE", 20, 36, 10, C_TEAL, font="Helvetica-Bold")
hline(c, 20, 52, 180, C_TEAL, 1.5)

# Company name: AdUnion
txt(c, "AdUnion", 20, 80, 62, C_AMBER_B, font="Helvetica-Bold")

# Opportunity lines
txt(c, "Three years of compounding growth.", 20, 162, 24, C_WHITE, font="Helvetica")
txt(c, "The next stage belongs to margin clarity.", 20, 195, 24, C_OFF_WHITE, font="Helvetica")

# Teal accent line below
hline(c, 20, 228, 380, C_TEAL, 1.5)

# Verified metrics
txt(c, "$9.3M revenue  •  81% three year growth  •  10 people",
    20, 244, 14, C_OFF_WHITE, font="Helvetica")
txt(c, "Smart50 2025 #8   •   AFR Fast 100 2025 #26   •   Deloitte Tech Fast 50 2025 #31",
    20, 265, 12, C_TEAL, font="Helvetica")

# Right side large year watermark
txt(c, "MELBOURNE", 680, 110, 10, C_TEAL, font="Helvetica-Bold", align="right", max_width=260)
txt(c, "2016", 680, 130, 46, HexColor("#222222"), font="Helvetica-Bold",
    align="right", max_width=260)
txt(c, "Founded", 680, 185, 10, HexColor("#444444"), font="Helvetica",
    align="right", max_width=260)

# Industry descriptor
txt(c, "Streaming Media Advertising Agency", 20, 295, 12,
    HexColor("#555555"), font="Helvetica-Oblique")

# Date footer
txt(c, "09 Jun 2026", 710, 520, 10, C_DRK_GREY, font="Helvetica",
    align="right", max_width=230)
txt(c, "Profit-Pulse.com.au", 20, 520, 10, C_DRK_GREY, font="Helvetica")


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE INSIGHT
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_BLACK)
fill_rect(c, 0, 0, 5, H, C_TEAL)           # teal left bar
fill_rect(c, 5, 0, W - 5, 52, C_DARK_PNL)  # header bar

txt(c, "THE INSIGHT", 20, 14, 13, C_AMBER_B, font="Helvetica-Bold")
txt(c, "AdUnion  |  09 Jun 2026", 550, 14, 12, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=390)

# ── LEFT COLUMN: Verified data ───────────────────────────────────────────────
left_x = 20
col_w  = 280

txt(c, "Verified facts", left_x, 66, 10, C_TEAL, font="Helvetica-Bold")
hline(c, left_x, 80, col_w, C_TEAL, 1)

data = [
    ("Revenue (FY2024)",       "$9.3 million"),
    ("Three year growth",      "81% average"),
    ("Revenue since 2023",     "Tripled"),
    ("Team size",              "10 people"),
    ("Smart50 2025",           "Rank 8 of 50"),
    ("AFR Fast 100 2025",      "Rank 26"),
    ("Deloitte Tech Fast 50",  "Rank 31"),
    ("Founded",                "2016"),
    ("Location",               "Cremorne, Melbourne VIC"),
    ("Managing Director",      "Robert Ong"),
    ("MD role note",           "Co-founder, leads AdUnion"),
]

y = 94
row_h = 30
for label, val in data:
    fill_rect(c, left_x, y, col_w, row_h - 2, HexColor("#161616"))
    txt(c, label,   left_x + 6, y + 6,  9, C_MID_GREY, font="Helvetica")
    txt(c, val,     left_x + 6, y + 18, 10, C_OFF_WHITE, font="Helvetica-Bold")
    y += row_h

txt(c, "Source: SmartCompany Smart50 2025 and public industry publications",
    left_x, y + 8, 8, HexColor("#555555"), font="Helvetica-Oblique")

# ── Vertical divider ─────────────────────────────────────────────────────────
vline(c, 310, 60, 455, C_TEAL, 1)

# ── RIGHT COLUMN: Wedge ──────────────────────────────────────────────────────
rx = 322
rw = 620

txt(c, "The commercial observation", rx, 66, 10, C_TEAL, font="Helvetica-Bold")
hline(c, rx, 80, rw, C_TEAL, 1)

wedge_paras = [
    ("AdUnion has compounded at 81 percent over three years and now "
     "manages streaming campaigns for Samsung, Tubi, Tangerine Telecom, "
     "and a growing roster across retail, travel, and financial services. "
     "With ten people generating $9.3 million in revenue, productivity is "
     "exceptional and the growth record is verifiable across three "
     "independent Australian lists.",
     C_OFF_WHITE),

    ("The question that follows rapid, diversified client growth is: which "
     "accounts drive the real margin after the full cost of service, team "
     "time, and platform investment are allocated? In most agencies at this "
     "stage, the answer is uneven. The clients that look largest by revenue "
     "are rarely the ones that yield the most after full cost allocation.",
     C_OFF_WHITE),

    ("The Customer Concentration and Profitability Map answers that question "
     "in three weeks: every client ranked by revenue, gross margin "
     "contribution, and effort to serve. The output is a clear action list: "
     "grow these, reprice these, reset these. At this scale and growth "
     "velocity, that map is the most leveraged financial insight available.",
     C_OFF_WHITE),
]

wy = 93
for para_text, para_colour in wedge_paras:
    wy = txt_wrapped(c, para_text, rx, wy, rw, 12, para_colour,
                     font="Helvetica", leading=18)
    wy += 14  # paragraph gap

# ── Bottom service bar ───────────────────────────────────────────────────────
fill_rect(c, 5, 478, W - 5, 52, C_TEAL_DRK)
fill_rect(c, 5, 478, 28, 52, C_TEAL)

txt(c, "Matched Service: Customer Concentration and Profitability Map  (G2)",
    45, 487, 13, C_AMBER_B, font="Helvetica-Bold")
txt(c, "Command tier  •  $3,950 one off  •  ProfitPulse verified price  "
        "•  Questionnaire confirms exact fit",
    45, 508, 11, C_OFF_WHITE, font="Helvetica")

txt(c, "Profit-Pulse.com.au", 20, 520, 9, C_DRK_GREY, font="Helvetica")
txt(c, "09 Jun 2026", 710, 520, 9, C_DRK_GREY, font="Helvetica",
    align="right", max_width=230)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE CALL TO ACTION
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_BLACK)
fill_rect(c, 0, 0, 5, H, C_TEAL)
fill_rect(c, 5, 0, W - 5, 52, C_DARK_PNL)

txt(c, "FIND THE RIGHT FIX FOR YOUR BUSINESS",
    20, 14, 13, C_AMBER_B, font="Helvetica-Bold")
txt(c, "AdUnion", 710, 14, 13, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=230)

# ── LEFT side: CTAs ──────────────────────────────────────────────────────────
cx = 20
cw = 580

# PRIMARY CTA box
fill_stroke_rect(c, cx, 62, cw, 148, C_TEAL_DRK, C_TEAL, 1)

txt(c, "Step 1: Answer a few quick questions",
    cx + 12, 72, 14, C_AMBER_B, font="Helvetica-Bold")
txt_wrapped(c,
    "See the solutions matched to your size and industry, each with a "
    "direct purchase option.",
    cx + 12, 96, cw - 24, 12, C_OFF_WHITE, font="Helvetica", leading=17)
txt_wrapped(c,
    "profit-pulse.com.au/full-suite-of-products?"
    "utm_source=outreach&utm_medium=pptx"
    "&utm_campaign=nightly_outreach&utm_content=adunion",
    cx + 12, 132, cw - 24, 10, C_TEAL, font="Helvetica", leading=14)
txt(c, "QUESTIONNAIRE  (primary front door)",
    cx + 12, 192, 10, C_TEAL, font="Helvetica-Bold")

# Separator
hline(c, cx, 218, cw, C_TEAL, 1)
txt(c, "Or start directly with the matched service",
    cx + 12, 226, 10, C_MID_GREY, font="Helvetica-Oblique")

# SECONDARY CTA box
fill_stroke_rect(c, cx, 244, cw, 108, C_AMBER_BG, C_AMBER_D, 1)

txt(c, "Customer Concentration and Profitability Map",
    cx + 12, 254, 13, C_AMBER_D, font="Helvetica-Bold")
txt(c, "Command tier  •  $3,950 one off  •  ProfitPulse verified price",
    cx + 12, 276, 11, C_OFF_WHITE, font="Helvetica")
txt(c, "Questionnaire confirms exact tier before purchase if preferred.",
    cx + 12, 293, 10, C_MID_GREY, font="Helvetica-Oblique")
txt(c, "buy.stripe.com/14AbJ21qw2U0ftK0ZV3ks1A",
    cx + 12, 318, 10, C_AMBER_D, font="Helvetica")

# Booking link
hline(c, cx, 360, cw, HexColor("#333333"), 1)
txt(c, "Prefer a conversation first?",
    cx + 12, 370, 11, C_OFF_WHITE, font="Helvetica")
txt_wrapped(c,
    "Book a complimentary discovery call: "
    "bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/",
    cx + 12, 390, cw - 24, 10, C_TEAL, font="Helvetica", leading=14)

# Supporting services note
hline(c, cx, 430, cw, HexColor("#222222"), 1)
txt(c, "Supporting services identified: A4 Strategic Growth Diagnostic  |  D4 Budgeting and Forecasting Setup",
    cx + 12, 440, 9, HexColor("#555555"), font="Helvetica")

# ── RIGHT side: Signature block ──────────────────────────────────────────────
sx = 625
sw = 315

fill_stroke_rect(c, sx, 62, sw, 380, HexColor("#0A0A0A"), HexColor("#222222"), 1)

# Name and credentials
txt(c, "NITESH ROOPA", sx + 16, 78, 18, C_AMBER_B, font="Helvetica-Bold")
txt(c, "CA, Managing Partner", sx + 16, 106, 12, C_WHITE, font="Helvetica")
txt(c, "ProfitPulse", sx + 16, 126, 18, C_TEAL, font="Helvetica-Bold")
hline(c, sx + 16, 154, sw - 32, C_TEAL, 1.5)

sig_items = [
    ("Profit-Pulse.com.au",            C_OFF_WHITE, "Helvetica"),
    ("Nitesh@Profit-Pulse.com.au",     C_TEAL,      "Helvetica"),
    ("+61 411 876 267",                C_OFF_WHITE, "Helvetica"),
]
sy = 164
for item_text, item_col, item_font in sig_items:
    txt(c, item_text, sx + 16, sy, 11, item_col, font=item_font)
    sy += 18

hline(c, sx + 16, sy + 4, sw - 32, HexColor("#333333"), 1)
sy += 16

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "CA qualification: SAICA South Africa",
    "QIC 2023 to 2025: Finance and Commercial",
    "Lead, AUD 10B Gympie Road Bypass Tunnel",
    "Nedbank CIB 2015 to 2022: Energy Finance",
    "and Principal and Equity Finance",
]
for line in cred_lines:
    txt(c, line, sx + 16, sy, 10, C_MID_GREY, font="Helvetica")
    sy += 15

hline(c, sx + 16, sy + 6, sw - 32, HexColor("#333333"), 1)
sy += 18
txt(c, "linkedin.com/in/nitesh-roopa-77594163",
    sx + 16, sy, 9, HexColor("#555555"), font="Helvetica")

# Footer
txt(c, "Profit-Pulse.com.au", 20, 520, 9, C_DRK_GREY, font="Helvetica")
txt(c, "09 Jun 2026", 710, 520, 9, C_DRK_GREY, font="Helvetica",
    align="right", max_width=230)

# ── Save ──────────────────────────────────────────────────────────────────────
c.save()
print(f"PDF saved: {out_path}")
