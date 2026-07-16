"""
ProfitPulse Brief PDF Builder, house style v3.3.
Target: Equality Media + Marketing | Date for: 17 Jul 2026
Direct PDF generation, 3 slides, brand colours only, zero dashes.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import stringWidth
import brief_content as bc

W = 960.0
H = 540.0
IN = 72.0  # points per inch

BLACK = HexColor("#000000")
TEAL = HexColor("#01A296")
AMBER_B = HexColor("#F8C806")
AMBER_D = HexColor("#F6A102")
GOLD = HexColor("#E3A712")
WHITE = HexColor("#FFFFFF")
OFF_WHITE = HexColor("#E6E5DE")

SERIF_B = "Times-Bold"
SERIF = "Times-Roman"
SANS = "Helvetica"
SANS_B = "Helvetica-Bold"
SANS_I = "Helvetica-Oblique"


def ry(y_top):
    """Convert top-origin y to ReportLab bottom-origin y."""
    return H - y_top


def rect(c, x, y_top, w, h, fill, stroke=None, lw=1):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(lw)
    c.rect(x, ry(y_top + h), w, h, fill=1, stroke=1 if stroke else 0)


def text(c, s, x, y_top, size, colour, font=SANS, align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = ry(y_top) - size * 0.85
    xx = x
    if align in ("right", "center") and max_width is not None:
        tw = stringWidth(s, font, size)
        if align == "right":
            xx = x + max_width - tw
        else:
            xx = x + (max_width - tw) / 2
    c.drawString(xx, baseline, s)
    return baseline


def wrap(s, font, size, max_width):
    words = s.split()
    lines = []
    cur = ""
    for w in words:
        trial = (cur + " " + w).strip()
        if stringWidth(trial, font, size) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def para(c, s, x, y_top, max_width, size, colour, font=SANS, leading=None, align="left"):
    leading = leading or size * 1.32
    lines = wrap(s, font, size, max_width)
    y = y_top
    for ln in lines:
        text(c, ln, x, y, size, colour, font=font, align=align, max_width=max_width)
        y += leading
    return y  # y_top for next element


def link(c, x, y_top, w, h, url):
    c.linkURL(url, (x, ry(y_top + h), x + w, ry(y_top)), relative=0, thickness=0)


# ── Fixed chrome ─────────────────────────────────────────────────────────────
STRIPE_W = 0.1 * IN
HEADER_H = 1.0 * IN
FOOTER_Y = 7.05 * IN
CONTENT_X0 = 0.35 * IN
CONTENT_X1 = 13.0 * IN
CONTENT_W = CONTENT_X1 - CONTENT_X0


def chrome(c, eyebrow, right_label="PROFITPULSE"):
    rect(c, 0, 0, W, H, WHITE)
    rect(c, 0, 0, W, HEADER_H, BLACK)
    text(c, eyebrow, CONTENT_X0, 0.32 * IN, 12, OFF_WHITE, font=SANS_B)
    text(c, right_label, CONTENT_X1 - 3.5 * IN, 0.32 * IN, 12, OFF_WHITE, font=SANS_B,
         align="right", max_width=3.5 * IN)
    rect(c, 0, 0, STRIPE_W, H, AMBER_D)


def footer(c):
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
         CONTENT_X0, FOOTER_Y, 8, BLACK, font=SANS)
    text(c, bc.DATE_FOR, CONTENT_X1 - 2.5 * IN, FOOTER_Y, 8, BLACK, font=SANS,
         align="right", max_width=2.5 * IN)


OUT_PATH = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMedia_17Jul2026.pdf"
c = canvas.Canvas(OUT_PATH, pagesize=(W, H))

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")

text(c, bc.COMPANY, CONTENT_X0, 1.15 * IN, 34, BLACK, font=SERIF_B)
text(c, bc.LOCATION_LINE, CONTENT_X0, 1.68 * IN, 12.5, BLACK, font=SANS)

# Stat cards row
n_cards = len(bc.STAT_CARDS)
gap = 0.15 * IN
card_w = (CONTENT_W - gap * (n_cards - 1)) / n_cards
card_top = 2.12 * IN
card_h = 1.35 * IN
for i, card in enumerate(bc.STAT_CARDS):
    cx = CONTENT_X0 + i * (card_w + gap)
    rect(c, cx, card_top, card_w, card_h, BLACK)
    rect(c, cx, card_top, card_w, 4, TEAL)  # teal top edge stripe
    text(c, card["number"], cx + 10, card_top + 14, 27, AMBER_B, font=SANS_B)
    lbl_lines = wrap(card["label"], SANS, 10, card_w - 20)
    ly = card_top + 52
    for ln in lbl_lines[:2]:
        text(c, ln, cx + 10, ly, 10, OFF_WHITE, font=SANS)
        ly += 13
    text(c, card["source"], cx + 10, card_top + card_h - 16, 7.5, OFF_WHITE, font=SANS_I)

# Simple chart: verified revenue trajectory, 3 points
chart_title_y = 270
text(c, "REVENUE TRAJECTORY, VERIFIED ($ MILLION)", CONTENT_X0, chart_title_y, 11, TEAL, font=SANS_B)
bar_area_top = chart_title_y + 34
bar_area_h = 44
max_val = max(v for _, v in bc.CHART_DATA)
bar_gap = 0.5 * IN
bar_w = 1.1 * IN
for i, (label, val) in enumerate(bc.CHART_DATA):
    bx = CONTENT_X0 + i * (bar_w + bar_gap)
    bh = bar_area_h * (val / max_val)
    rect(c, bx, bar_area_top + (bar_area_h - bh), bar_w, bh, TEAL)
    text(c, f"${val:g}M", bx, bar_area_top + (bar_area_h - bh) - 13, 11, BLACK, font=SANS_B)
category_y = bar_area_top + bar_area_h + 14
for i, (label, val) in enumerate(bc.CHART_DATA):
    bx = CONTENT_X0 + i * (bar_w + bar_gap)
    text(c, label, bx, category_y, 10, BLACK, font=SANS)
source_y = category_y + 15
text(c, bc.CHART_SOURCE, CONTENT_X0, source_y, 7.5, BLACK, font=SANS_I)

# Key commercial signals
sig_top = source_y + 20
text(c, "KEY COMMERCIAL SIGNALS", CONTENT_X0, sig_top, 11, TEAL, font=SANS_B)
sy = sig_top + 20
for s in bc.KEY_SIGNALS:
    text(c, "•", CONTENT_X0, sy, 10.5, BLACK, font=SANS_B)
    text(c, s, CONTENT_X0 + 12, sy, 10.5, BLACK, font=SANS)
    sy += 15

footer(c)
c.showPage()

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════
chrome(c, "THE OPPORTUNITY")
text(c, f"{bc.COMPANY}: Three commercial observations from ProfitPulse",
     CONTENT_X0, 1.18 * IN, 14, BLACK, font=SERIF_B)

col_gap = 0.15 * IN
col_w = (CONTENT_W - col_gap * 2) / 3
col_top = 1.55 * IN
col_h = 4.75 * IN
FILL_MAP = {"TEAL": TEAL, "BLACK": BLACK, "GOLD": GOLD}
for i, obs in enumerate(bc.OBSERVATIONS):
    cx = CONTENT_X0 + i * (col_w + col_gap)
    fill = FILL_MAP[obs["fill"]]
    rect(c, cx, col_top, col_w, col_h, fill)
    text_colour = WHITE if obs["fill"] == "BLACK" else BLACK
    muted = OFF_WHITE if obs["fill"] == "BLACK" else BLACK
    pad = 16
    text(c, obs["index"], cx + pad, col_top + 20, 34, text_colour, font=SERIF_B)
    hy = col_top + 70
    hy_end = para(c, obs["header"], cx + pad, hy, col_w - pad * 2, 13.5, text_colour, font=SANS_B, leading=17)
    by = hy_end + 8
    para(c, obs["body"], cx + pad, by, col_w - pad * 2, 10.5, text_colour, font=SANS, leading=14)

para(c, bc.CLOSING_LINE, CONTENT_X0, col_top + col_h + 16, CONTENT_W, 10.5,
     BLACK, font=SANS_I, leading=14)
footer(c)
c.showPage()

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
chrome(c, "THE RECOMMENDATION")

left_x = CONTENT_X0
left_w = 7.45 * IN
div_x = left_x + left_w + 0.1 * IN
right_x = div_x + 0.15 * IN
right_w = CONTENT_X1 - right_x

y = 1.25 * IN
text(c, bc.PRIMARY_SERVICE_NAME, left_x, y, 23, BLACK, font=SERIF_B)
y += 36
text(c, f"{bc.PRIMARY_SERVICE_PRICE} {bc.PRIMARY_SERVICE_CADENCE}", left_x, y, 15.5, TEAL, font=SANS_B)
y += 30
y = para(c, bc.PRIMARY_SERVICE_DESC, left_x, y, left_w, 11.5, BLACK, font=SANS, leading=16)
y += 30

# Step one block
step_h = 1.3 * IN
rect(c, left_x, y, left_w, step_h, OFF_WHITE, stroke=TEAL, lw=1)
text(c, "Step one, answer a few quick questions", left_x + 16, y + 20, 13.5, TEAL, font=SANS_B)
text(c, "See the solutions matched to your size and industry.", left_x + 16, y + 48, 11.5,
     BLACK, font=SANS)
text(c, bc.QUESTIONNAIRE_CLEAN, left_x + 16, y + 80, 13, TEAL, font=SANS_B)
link(c, left_x + 16, y + 76, stringWidth(bc.QUESTIONNAIRE_CLEAN, SANS_B, 13) + 4, 18,
     bc.QUESTIONNAIRE_URL_HTML)
y += step_h + 26

# Direct purchase CTA
cta = "Purchase the suggested product now to get started"
cta_h = 0.62 * IN
rect(c, left_x, y, left_w, cta_h, AMBER_D)
text(c, cta, left_x, y + 22, 13, BLACK, font=SANS_B, align="center", max_width=left_w)
link(c, left_x, y, left_w, cta_h, bc.PRIMARY_SERVICE_STRIPE)
y += cta_h + 26

text(c, "Prefer a conversation first?", left_x, y, 11.5, BLACK, font=SANS)
y += 22
book_txt = "Book a complimentary discovery call"
text(c, book_txt, left_x, y, 12.5, TEAL, font=SANS_B)
link(c, left_x, y - 2, stringWidth(book_txt, SANS_B, 12.5) + 4, 16, bc.BOOKING_LINK)

# Divider
rect(c, div_x, 1.25 * IN, 1.2, 5.5 * IN, TEAL)

# Right column: credibility panel
ry_ = 1.25 * IN
text(c, "NITESH ROOPA", right_x, ry_, 17, BLACK, font=SERIF_B)
ry_ += 26
text(c, "CA, Managing Partner, ProfitPulse", right_x, ry_, 12, BLACK, font=SANS)
ry_ += 44
CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
for cline in CREDS:
    text(c, "•", right_x, ry_, 11, TEAL, font=SANS_B)
    ry_ = para(c, cline, right_x + 14, ry_, right_w - 14, 11, BLACK, font=SANS, leading=15)
    ry_ += 20
ry_ += 12
rect(c, right_x, ry_, right_w, 1.2, TEAL)
ry_ += 30
CONTACT = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for cl in CONTACT:
    text(c, cl, right_x, ry_, 11.5, BLACK, font=SANS)
    ry_ += 24

footer(c)
c.showPage()

c.save()
print(f"PDF built at {OUT_PATH}")
