"""
ProfitPulse Brief PDF Builder
Target: Taskforce Australia | Date for: 18 Jun 2026
White-background house style. Seven approved brand colours only. Zero dashes. No tier names.
Page: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

W = 960
H = 540

# ── Seven approved brand colours only ────────────────────────────────────────
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")

STRIPE_W  = 6
HEADER_H  = 50
CONTENT_L = STRIPE_W + 13
DATE_STR  = "18 Jun 2026"


def rl_y(screen_y):
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def fill_stroke_rect(c, x, y_top, w, h, fc, sc, lw=1):
    c.setFillColor(fc)
    c.setStrokeColor(sc)
    c.setLineWidth(lw)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold",
        align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline_y = rl_y(y_top + size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline_y, text)


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica",
                leading=None):
    if leading is None:
        leading = size * 1.45
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
    c.line(x, rl_y(y_top), x + w, rl_y(y_top))


def add_link(c, x, y_top, w, h, url):
    c.linkURL(url, (x, rl_y(y_top + h), x + w, rl_y(y_top)), relative=0)


def standard_header(c, eyebrow, right_text="PROFITPULSE"):
    fill_rect(c, 0, 0, W, H, C_WHITE)
    fill_rect(c, 0, 0, STRIPE_W, H, C_AMBER_D)
    fill_rect(c, STRIPE_W, 0, W - STRIPE_W, HEADER_H, C_BLACK)
    txt(c, eyebrow, CONTENT_L, 14, 11, C_OFF_WHITE, font="Helvetica-Bold")
    txt(c, right_text, 680, 14, 11, C_TEAL, font="Helvetica-Bold",
        align="right", max_width=266)


def standard_footer(c):
    footer_y = 524
    txt(c, "Nitesh Roopa CA  |  Managing Partner  |  ProfitPulse  |  Profit-Pulse.com.au",
        CONTENT_L, footer_y, 8, C_BLACK, font="Helvetica")
    txt(c, DATE_STR, 714, footer_y, 8, C_BLACK, font="Helvetica",
        align="right", max_width=232)


def stat_card(c, x, y_top, w, h, number, label, source):
    """Black tile, teal top stripe, amber number, off-white label, teal source."""
    fill_rect(c, x, y_top, w, 3, C_TEAL)
    fill_rect(c, x, y_top + 3, w, h - 3, C_BLACK)
    txt(c, number, x + 8, y_top + 10, 22, C_AMBER_B, font="Helvetica-Bold")
    txt_wrapped(c, label, x + 8, y_top + 38, w - 16, 9, C_OFF_WHITE,
                font="Helvetica", leading=12)
    txt(c, source, x + 8, y_top + h - 14, 7, C_TEAL,
        font="Helvetica-Oblique")


# ── Output file ───────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_18Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Taskforce Australia | ProfitPulse Brief | 18 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
standard_header(c, "COMMERCIAL INTELLIGENCE BRIEF")

txt(c, "Taskforce Australia", CONTENT_L, 57, 34, C_BLACK, font="Times-Bold")
txt(c, "Property maintenance, safety compliance and PropTech platform  |  Burnley, Melbourne VIC",
    CONTENT_L, 97, 11, C_TEAL, font="Helvetica")
hline(c, CONTENT_L, 113, W - CONTENT_L - 14, C_TEAL, 1)

# Stat cards
card_top = 118
card_h   = 86
card_gap = 4
available = W - CONTENT_L - 14
card_w = (available - card_gap * 5) // 6

cards = [
    ("$12.8M",  "Revenue FY2025",          "Smart50 2025, rank 37"),
    ("31%",     "Revenue growth FY2025",   "Smart50 2025"),
    ("19",      "Employees",               "Smart50 2025"),
    ("5,000+",  "National tradespeople",   "taskforce.com.au"),
    ("140K+",   "Jobs via RentSafe",       "taskforce.com.au"),
    ("Rank 37", "Smart50 2025",            "Up from rank 46 in 2024"),
]
cx = CONTENT_L
for num, lbl, src in cards:
    stat_card(c, cx, card_top, card_w, card_h, num, lbl, src)
    cx += card_w + card_gap

# Revenue chart (left)
chart_l   = CONTENT_L
chart_top = 212
chart_w   = 290
chart_h   = 278

txt(c, "REVENUE GROWTH (VERIFIED)", chart_l, chart_top, 9, C_TEAL, font="Helvetica-Bold")

bar_area_h = 200
bar_base_y = chart_top + 14 + bar_area_h
bar_w_e    = 70

bar_24_h = int(bar_area_h * (9.7 / 13.5))
fill_rect(c, chart_l + 36, bar_base_y - bar_24_h, bar_w_e, bar_24_h, C_TEAL)
txt(c, "$9.7M", chart_l + 36, bar_base_y - bar_24_h - 14, 9, C_TEAL, font="Helvetica-Bold")
txt(c, "FY2024", chart_l + 42, bar_base_y + 6, 9, C_BLACK, font="Helvetica")

bar_25_h = int(bar_area_h * (12.8 / 13.5))
fill_rect(c, chart_l + 168, bar_base_y - bar_25_h, bar_w_e, bar_25_h, C_AMBER_D)
txt(c, "$12.8M", chart_l + 163, bar_base_y - bar_25_h - 14, 9, C_AMBER_D, font="Helvetica-Bold")
txt(c, "FY2025", chart_l + 174, bar_base_y + 6, 9, C_BLACK, font="Helvetica")

txt(c, "Source: SmartCompany Smart50 2024 and 2025 award citations",
    chart_l, chart_top + chart_h, 7, C_TEAL, font="Helvetica-Oblique")

# Key Commercial Signals (right panel)
sig_l   = chart_l + chart_w + 22
sig_w   = W - sig_l - 14
sig_top = chart_top

txt(c, "KEY COMMERCIAL SIGNALS", sig_l, sig_top, 9, C_TEAL, font="Helvetica-Bold")
hline(c, sig_l, sig_top + 14, sig_w, C_TEAL, 0.75)

signals = [
    ("Housing division described publicly as strongest and most profitable line; new housing provider clients onboarded in last 12 months.",
     "Smart50 2025, Nov 2025"),
    ("Partnership with Real+ announced March 2025; CEO Jason Bright co-presented webinar on rental compliance.",
     "taskforce.com.au, Mar 2025"),
    ("Smart50 rank improved from 46 to 37 in one year, confirming accelerating growth relative to peers.",
     "SmartCompany Smart50 2024 and 2025"),
    ("Telstra Best of Business VIC State Winner, Outstanding Growth. PropTech Award winner.",
     "eliteagent.com; proptechaustralia.com.au"),
    ("RentSafe: 140,000 plus jobs across 20 consumer brands, 300 real estate offices, 180 real estate companies.",
     "taskforce.com.au (public)"),
]

sy = sig_top + 22
for sig_text, sig_src in signals:
    fill_rect(c, sig_l, sy, 3, 11, C_TEAL)
    sy_end = txt_wrapped(c, sig_text, sig_l + 8, sy, sig_w - 8, 9,
                         C_BLACK, font="Helvetica", leading=13)
    txt(c, sig_src, sig_l + 8, sy_end + 2, 7, C_TEAL, font="Helvetica-Oblique")
    sy = sy_end + 18

standard_footer(c)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
standard_header(c, "THE OPPORTUNITY", "Taskforce Australia")

txt(c, "Three commercial observations from ProfitPulse",
    CONTENT_L, 58, 11, C_BLACK, font="Helvetica")

col_top = 76
col_h   = 435
col_gap = 3
col_w   = (W - CONTENT_L - 14 - col_gap * 2) // 3
c1_l = CONTENT_L
c2_l = c1_l + col_w + col_gap
c3_l = c2_l + col_w + col_gap

fills  = [C_TEAL,   C_BLACK,   C_GOLD]
nums   = ["01",     "02",      "03"]
num_c  = [C_WHITE,  C_AMBER_B, C_BLACK]
hdr_c  = [C_WHITE,  C_AMBER_B, C_BLACK]
body_c = [C_WHITE,  C_OFF_WHITE, C_BLACK]
src_c  = [C_WHITE,  C_OFF_WHITE, C_BLACK]

headers = [
    "Revenue velocity needs a financial map to match it",
    "The housing division is the edge. A financial map confirms it.",
    "High revenue per head signals a coming capacity inflection",
]
bodies = [
    ("Taskforce grew revenue from $9.7 million to $12.8 million in one year, a 32 percent "
     "increase on top of 26 percent the year before. That trajectory is verifiably compounding "
     "across two independent Smart50 citations. At this pace, every resourcing and client "
     "onboarding decision compounds financially. The architecture that should sit behind "
     "compounding growth includes a margin model by service line, a capacity analysis for the "
     "team, and a capital allocation plan that maps each growth scenario in dollar terms."),
    ("The housing division has been publicly identified as the strongest and most profitable "
     "segment of the business. Yet without a line-level profitability analysis, the capital "
     "allocation case for accelerating housing sits on instinct rather than numbers. "
     "Understanding the true margin by service line changes every resourcing, pricing, and "
     "sales focus decision. This is the financial work that turns a strong instinct into a "
     "deliberate and defensible strategy with a clear financial outcome."),
    ("With 19 people and $12.8 million in revenue, the revenue per head ratio is well above "
     "average for a services business. That is a mark of an exceptional technology platform. "
     "It is also a signal that a ceiling on further growth without structural change is "
     "approaching. A strategic growth diagnostic maps exactly where that ceiling sits, what "
     "the next team or technology investment needs to look like, and what three scenarios for "
     "the next 12 months look like in dollar terms so the decision to grow is made "
     "deliberately rather than reactively."),
]
src_notes = [
    "Smart50 2024 and 2025 (revenue and growth data)",
    "Smart50 2025 (housing division public statement)",
    "Smart50 2025 (revenue $12.8M, employees 19)",
]

for cl, fc, num, nc, hc, bc, sc, hdr, bdy, src in zip(
        [c1_l, c2_l, c3_l], fills, nums, num_c, hdr_c, body_c, src_c,
        headers, bodies, src_notes):
    fill_rect(c, cl, col_top, col_w, col_h, fc)
    txt(c, num, cl + 14, col_top + 12, 32, nc, font="Times-Bold")
    txt_wrapped(c, hdr, cl + 14, col_top + 54, col_w - 28, 11, hc,
                font="Helvetica-Bold", leading=15)
    hline(c, cl + 14, col_top + 102, col_w - 28, nc, 0.75)
    txt_wrapped(c, bdy, cl + 14, col_top + 110, col_w - 28, 9.5, bc,
                font="Helvetica", leading=14)
    txt(c, src, cl + 14, col_top + col_h - 18, 7, sc, font="Helvetica-Oblique")

txt(c,
    "These observations are offered in good faith. Taskforce has built something genuinely impressive. "
    "The question is simply whether the financial architecture now matches the commercial ambition.",
    CONTENT_L, col_top + col_h + 8, 9.5, C_BLACK, font="Helvetica-Oblique",
    align="center", max_width=W - CONTENT_L - 14)

standard_footer(c)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3: THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
standard_header(c, "THE RECOMMENDATION", "Taskforce Australia")

LEFT_W  = 560
RIGHT_L = CONTENT_L + LEFT_W + 16
RIGHT_W = W - RIGHT_L - 14

# ── Left column ──────────────────────────────────────────────────────────────
txt(c, "Strategic Growth Diagnostic",
    CONTENT_L, 57, 22, C_BLACK, font="Times-Bold")

txt(c, "$5,000 one off", CONTENT_L, 87, 17, C_TEAL, font="Helvetica-Bold")
txt(c, "ProfitPulse verified price. The questionnaire confirms the exact figure for your size and industry.",
    CONTENT_L, 110, 9, C_BLACK, font="Helvetica")

hline(c, CONTENT_L, 124, LEFT_W, C_TEAL, 0.75)

txt_wrapped(c,
    "A six-week engagement that maps revenue, capacity, and margin headroom, "
    "then produces a 12-month growth plan with funding and capital allocation "
    "steps spelled out in three scenarios. Exactly matched to a business at "
    "Taskforce's growth rate and stage.",
    CONTENT_L, 130, LEFT_W, 10, C_BLACK, font="Helvetica", leading=15)

# Step one block: teal outline with white fill
fill_stroke_rect(c, CONTENT_L, 185, LEFT_W, 82, C_WHITE, C_TEAL, 1)
fill_rect(c, CONTENT_L, 185, 4, 82, C_TEAL)  # left teal accent strip

txt(c, "STEP ONE", CONTENT_L + 14, 192, 8, C_TEAL, font="Helvetica-Bold")
txt(c, "Answer a few quick questions", CONTENT_L + 14, 207, 12, C_BLACK, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry.",
    CONTENT_L + 14, 227, 9.5, C_BLACK, font="Helvetica")

# Clean questionnaire URL (no UTM tags per brief link rules)
txt(c, "profit-pulse.com.au/full-suite-of-products",
    CONTENT_L + 14, 246, 10, C_TEAL, font="Helvetica-Bold")
add_link(c, CONTENT_L + 14, 243, 310, 16,
         "https://profit-pulse.com.au/full-suite-of-products")

# Direct purchase CTA: amber tile, Stripe link hidden behind text
fill_rect(c, CONTENT_L, 276, LEFT_W, 30, C_AMBER_D)
txt(c, "Purchase the suggested product now to get started",
    CONTENT_L + 10, 281, 11, C_WHITE, font="Helvetica-Bold")
add_link(c, CONTENT_L, 276, LEFT_W, 30,
         "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h")

# Prefer conversation
hline(c, CONTENT_L, 316, LEFT_W, C_BLACK, 0.5)
txt(c, "Prefer a conversation first?", CONTENT_L, 323, 9.5, C_BLACK, font="Helvetica")
txt(c, "Book a complimentary discovery call",
    CONTENT_L, 340, 10, C_TEAL, font="Helvetica-Bold")
add_link(c, CONTENT_L, 337, 230, 16,
         "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Supporting services note
hline(c, CONTENT_L, 366, LEFT_W, C_OFF_WHITE, 0.5)
txt(c, "Supporting services: G3 Product and Service Line Profitability  |  B1 Fractional CFO Partnership",
    CONTENT_L, 373, 8, C_BLACK, font="Helvetica")

# ── Right column: Nitesh credibility panel ───────────────────────────────────
fill_rect(c, RIGHT_L, 57, RIGHT_W, 440, C_BLACK)
fill_rect(c, RIGHT_L, 57, 4, 440, C_TEAL)   # left teal accent on panel

txt(c, "Nitesh Roopa", RIGHT_L + 14, 66, 17, C_AMBER_B, font="Times-Bold")
txt(c, "CA, Managing Partner", RIGHT_L + 14, 92, 10, C_WHITE, font="Helvetica")
txt(c, "ProfitPulse", RIGHT_L + 14, 110, 14, C_TEAL, font="Helvetica-Bold")
hline(c, RIGHT_L + 14, 130, RIGHT_W - 28, C_TEAL, 1)

creds = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3B, Cahora Bassa",
    "Total GRBT value in QLD: AUD 10B",
]
cy = 138
for line in creds:
    txt(c, line, RIGHT_L + 14, cy, 9, C_OFF_WHITE, font="Helvetica")
    cy += 15

hline(c, RIGHT_L + 14, cy + 4, RIGHT_W - 28, C_OFF_WHITE, 0.5)
cy += 12

contacts = [
    ("Profit-Pulse.com.au",                    C_OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au",             C_TEAL),
    ("+61 411 876 267",                        C_OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163",  C_OFF_WHITE),
]
for ct, cc in contacts:
    txt(c, ct, RIGHT_L + 14, cy, 9, cc, font="Helvetica")
    cy += 14

standard_footer(c)

c.save()
print(f"PDF saved: {out_path}")
