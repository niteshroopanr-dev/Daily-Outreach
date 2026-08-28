"""
ProfitPulse Brief PDF Builder
Target: Attekus | Date: 29 Aug 2026
Direct PDF generation, 3 slides, matches Brief_Attekus_29Aug2026.pptx layout exactly.
Brand colours only. Zero dashes. Page size 960pt x 540pt (13.333in x 7.5in at 72dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color

IN = 72  # points per inch

W = 13.333 * IN
H = 7.5 * IN

C_BLACK   = HexColor("#000000")
C_TEAL    = HexColor("#01A296")
C_AMBER_B = HexColor("#F8C806")
C_AMBER_D = HexColor("#F6A102")
C_GOLD    = HexColor("#E3A712")
C_WHITE   = HexColor("#FFFFFF")
C_OFFWHITE= HexColor("#E6E5DE")

SERIF = "Times-Bold"
SERIF_R = "Times-Roman"
SANS = "Helvetica"
SANS_B = "Helvetica-Bold"

QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_URL = "https://buy.stripe.com/cNi9AUd9e9iodlC8sn3ks0k"
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
DATE_STAMP = "29 Aug 2026"


def rl_y(screen_y):
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def bordered_rect(c, x, y_top, w, h, fill_colour, line_colour, line_width=1):
    c.setFillColor(fill_colour)
    c.setStrokeColor(line_colour)
    c.setLineWidth(line_width)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=1)


def txt(c, text, x, y_top, size, colour, font=SANS_B, align="left", max_width=None, link=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline_y = rl_y(y_top + size)
    draw_x = x
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        draw_x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        draw_x = x + (max_width - tw) / 2
    c.drawString(draw_x, baseline_y, text)
    if link:
        tw = c.stringWidth(text, font, size)
        c.linkURL(link, (draw_x, baseline_y - 2, draw_x + tw, baseline_y + size), relative=0, thickness=0)


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font=SANS, leading=None, align="left"):
    if leading is None:
        leading = size * 1.3
    c.setFillColor(colour)
    c.setFont(font, size)
    words = text.split()
    lines, current = [], ""
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
    return y


def hline(c, x, y_top, w, colour, thickness=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def footer(c):
    txt(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
        0.4 * IN, 7.14 * IN, 9, C_OFFWHITE, font=SANS)
    txt(c, DATE_STAMP, 10.6 * IN, 7.14 * IN, 9, C_OFFWHITE, font=SANS, align="right", max_width=2.3 * IN)


def base_page(c):
    fill_rect(c, 0, 0, W, H, C_BLACK)
    fill_rect(c, 0, 0, 0.09 * IN, H, C_AMBER_D)


out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Attekus_29Aug2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Attekus, ProfitPulse Brief, 29 Aug 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Capital Allocation Review")

# ============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ============================================================================
base_page(c)

txt(c, "COMMERCIAL INTELLIGENCE BRIEF", 0.4 * IN, 0.32 * IN, 12, C_GOLD, font=SANS_B)
txt(c, "PROFITPULSE", 9.5 * IN, 0.32 * IN, 12, C_TEAL, font=SANS_B, align="right", max_width=3.4 * IN)

txt(c, "Attekus", 0.4 * IN, 0.72 * IN, 54, C_AMBER_B, font=SERIF)

txt(c, "Cloud based booking and event software for councils, Brisbane based",
    0.4 * IN, 1.72 * IN, 16, C_WHITE, font=SANS)

hline(c, 0.4 * IN, 2.12 * IN, 12.5 * IN, C_TEAL, 1.5)

stats = [
    ("$9.2M", ["FY2025 revenue", "reported"], "getlatka.com, Sept 2025"),
    ("$5M",   ["First external", "capital raised"], "QIC Ventures, Oct 2024"),
    ("20%",   ["Of ANZ councils", "on Bookable"], "Five V Capital profile"),
    ("42",    ["People across", "the business"], "getlatka.com profile"),
    ("2017",  ["Year founded,", "Brisbane"], "Startup Daily report"),
    ("WINNER",["ANZ High Growth", "Award 2025"], "Lord Mayor Business Awards"),
]
tile_w = 1.98 * IN
tile_h = 1.85 * IN
gap = 0.11 * IN
x0 = 0.4 * IN
y0 = 2.34 * IN
for i, (num, label_lines, source) in enumerate(stats):
    x = x0 + i * (tile_w + gap)
    bordered_rect(c, x, y0, tile_w, tile_h, C_BLACK, C_OFFWHITE, 0.75)
    fill_rect(c, x, y0, tile_w, 0.07 * IN, C_TEAL)
    txt(c, num, x + 0.12 * IN, y0 + 0.24 * IN, 28, C_AMBER_B, font=SERIF)
    ly = y0 + 0.86 * IN
    for line in label_lines:
        txt(c, line, x + 0.12 * IN, ly, 11, C_OFFWHITE, font=SANS)
        ly += 0.19 * IN
    txt(c, source, x + 0.12 * IN, y0 + 1.52 * IN, 8, C_GOLD, font=SANS)

txt(c, "KEY COMMERCIAL SIGNALS", 0.4 * IN, 4.42 * IN, 13, C_TEAL, font=SANS_B)

signals = [
    "Raised AUD 5 million first ever external round, led by Five V Capital, October 2024.",
    "QIC Ventures joined as co investor in the same institutional funding round.",
    "Named 2025 ANZ High Growth Business at Brisbane Lord Mayor Business Awards.",
    "Bookable platform now used by close to 20 percent of councils across Australia and NZ.",
    "New capital earmarked for UK council entry plus ANZ education and government sectors.",
    "CEO Peter Suchting joined in 2023, leading beyond founder only leadership.",
]
sy = 4.80 * IN
for s in signals:
    txt(c, "•  " + s, 0.4 * IN, sy, 13, C_OFFWHITE, font=SANS)
    sy += 0.315 * IN

footer(c)

# ============================================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================================
c.showPage()
base_page(c)

txt(c, "THE OPPORTUNITY", 0.4 * IN, 0.32 * IN, 24, C_AMBER_B, font=SERIF)
txt(c, "Attekus, three commercial observations from ProfitPulse",
    0.4 * IN, 0.84 * IN, 14, C_WHITE, font=SANS)
hline(c, 0.4 * IN, 1.22 * IN, 12.5 * IN, C_TEAL, 1.5)

columns = [
    {
        "idx": "01", "header": "New Capital, New Scrutiny", "fill": C_TEAL, "text_col": C_BLACK,
        "tag": "ALIGNED TO THE RECOMMENDED SERVICE",
        "body": ("Five V Capital and QIC Ventures backed Attekus with its first ever "
                 "external raise in October 2024, funding entry into the United Kingdom "
                 "council market plus ANZ education and government agencies. A Capital "
                 "Allocation Review ranks each expansion path by expected return before "
                 "further capital is committed."),
    },
    {
        "idx": "02", "header": "A Board That Expects More", "fill": C_BLACK, "text_col": C_OFFWHITE,
        "tag": None,
        "body": ("Peter Suchting joined as Chief Executive Officer in 2023, and two "
                 "institutional shareholders now sit alongside the founders. Reporting "
                 "built for founder only oversight rarely satisfies venture capital "
                 "standards. An Annual Plan and Board Pack lifts that reporting rhythm "
                 "to the standard Five V Capital and QIC Ventures now expect."),
    },
    {
        "idx": "03", "header": "Three Markets, One Plan", "fill": C_GOLD, "text_col": C_BLACK,
        "tag": None,
        "body": ("Attekus already serves close to 20 percent of councils across "
                 "Australia and New Zealand. The new capital targets three further "
                 "markets at once: the United Kingdom, education, and government "
                 "agencies. A Strategic Growth Diagnostic sequences that expansion "
                 "into a costed 12 month plan."),
    },
]

col_w = 4.03 * IN
col_gap = 0.19 * IN
col_y = 1.5 * IN
col_h = 4.55 * IN
for i, col in enumerate(columns):
    x = 0.4 * IN + i * (col_w + col_gap)
    if col["fill"] == C_BLACK:
        bordered_rect(c, x, col_y, col_w, col_h, C_BLACK, C_OFFWHITE, 0.75)
    else:
        fill_rect(c, x, col_y, col_w, col_h, col["fill"])
    txt(c, col["idx"], x + 0.22 * IN, col_y + 0.22 * IN, 32, col["text_col"], font=SERIF)
    txt_wrapped(c, col["header"], x + 0.22 * IN, col_y + 0.9 * IN, col_w - 0.44 * IN, 17,
                col["text_col"], font=SERIF, leading=21)
    body_top = col_y + 1.55 * IN
    if col["tag"]:
        txt(c, col["tag"], x + 0.22 * IN, col_y + 1.38 * IN, 9, col["text_col"], font=SANS_B)
        body_top = col_y + 1.65 * IN
    txt_wrapped(c, col["body"], x + 0.22 * IN, body_top, col_w - 0.44 * IN, 12.5,
                col["text_col"], font=SANS, leading=16.5)

txt_wrapped(c, ("Seven years of disciplined, capital light growth is a rare foundation. ProfitPulse "
                "would welcome the chance to help Attekus put its first outside capital to its best "
                "possible use."),
            0.4 * IN, 6.3 * IN, 12.5 * IN, 13, C_OFFWHITE, font=SANS, leading=17)

footer(c)

# ============================================================================
# SLIDE 3: THE RECOMMENDATION
# ============================================================================
c.showPage()
base_page(c)

txt(c, "THE RECOMMENDATION", 0.4 * IN, 0.32 * IN, 24, C_AMBER_B, font=SERIF)
hline(c, 0.4 * IN, 0.82 * IN, 12.5 * IN, C_TEAL, 1.5)

lx = 0.4 * IN
lw = 7.4 * IN

txt(c, "Capital Allocation Review", lx, 1.05 * IN, 27, C_AMBER_B, font=SERIF)
txt(c, "AUD 6,000 one off", lx, 1.66 * IN, 16, C_WHITE, font=SANS_B)
txt_wrapped(c, ("An independent ranking of the United Kingdom, education and government "
                "expansion paths by expected return, with a redeployment plan the board can act on."),
            lx, 2.10 * IN, lw, 13, C_OFFWHITE, font=SANS, leading=17)

bordered_rect(c, lx, 2.95 * IN, lw, 1.55 * IN, C_BLACK, C_TEAL, 1)
txt(c, "STEP ONE", lx + 0.22 * IN, 3.12 * IN, 11, C_TEAL, font=SANS_B)
txt(c, "See the solutions matched to your size and industry",
    lx + 0.22 * IN, 3.44 * IN, 13, C_WHITE, font=SANS)
txt(c, QUESTIONNAIRE_CLEAN, lx + 0.22 * IN, 3.84 * IN, 14, C_TEAL, font=SANS_B, link=QUESTIONNAIRE_URL)

bordered_rect(c, lx, 4.65 * IN, lw, 0.95 * IN, C_BLACK, C_AMBER_D, 1)
txt(c, "Purchase the suggested product now to get started",
    lx + 0.22 * IN, 4.9 * IN, 14, C_AMBER_D, font=SANS_B, link=STRIPE_URL)

txt(c, "Prefer a conversation first", lx, 5.8 * IN, 12, C_OFFWHITE, font=SANS)
txt(c, "Book a complimentary discovery call", lx, 6.14 * IN, 13, C_TEAL, font=SANS_B, link=BOOKING_URL)

rx = 8.15 * IN
rw = 4.75 * IN
bordered_rect(c, rx, 1.05 * IN, rw, 5.55 * IN, C_BLACK, C_OFFWHITE, 0.75)
fill_rect(c, rx, 1.05 * IN, rw, 0.06 * IN, C_GOLD)

txt_wrapped(c, "Nitesh Roopa, CA, Managing Partner, ProfitPulse",
            rx + 0.25 * IN, 1.3 * IN, rw - 0.5 * IN, 16, C_AMBER_B, font=SERIF, leading=19)

cred_lines = [
    "16 years of experience across 4 countries",
    "Over 52 deals executed and managed",
    "Queensland GRBT project value over AUD 10 billion",
    "CA qualification, SAICA South Africa",
]
cy = 2.35 * IN
for cred in cred_lines:
    cy = txt_wrapped(c, "•  " + cred, rx + 0.25 * IN, cy, rw - 0.5 * IN, 12.5, C_OFFWHITE, font=SANS, leading=16)
    cy += 0.06 * IN

hline(c, rx + 0.25 * IN, 4.15 * IN, rw - 0.5 * IN, C_TEAL, 1.5)

contact_lines = [
    ("Profit-Pulse.com.au", C_OFFWHITE, 13, SANS),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL, 13, SANS_B),
    ("+61 411 876 267", C_OFFWHITE, 13, SANS),
    ("linkedin.com/in/nitesh-roopa-77594163", C_OFFWHITE, 11, SANS),
    ("Brisbane, Australia", C_OFFWHITE, 11, SANS),
]
cy = 4.35 * IN
for text_, colour, size, font in contact_lines:
    txt(c, text_, rx + 0.25 * IN, cy, size, colour, font=font)
    cy += 0.27 * IN

footer(c)

c.save()
print(f"PDF saved: {out_path}")
