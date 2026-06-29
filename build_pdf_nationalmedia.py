"""
ProfitPulse Brief PDF Builder
Target: National Media | Date: 30 Jun 2026
Direct PDF generation, 3 slides, brand colours only, zero dashes. No tier names.
Section 6 house style: white background, amber left stripe, black header band,
black stat card tiles with teal accent, teal section labels.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.units import inch

pt = 1
W = 960
H = 540

# ── Brand colours ─────────────────────────────────────────────────────────────
C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_MID_GREY  = HexColor("#888888")
C_DRK_GREY  = HexColor("#444444")
C_LT_GREY   = HexColor("#CCCCCC")
C_NEAR_WHT  = HexColor("#F8F7F2")
C_PANEL_LT  = HexColor("#F0F0EB")


def rl_y(screen_y):
    return H - screen_y


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, stroke_colour, line_width=1):
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica",
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
                leading=None, align="left"):
    if leading is None:
        leading = size * 1.5
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
    return y


def hline(c, x, y_top, w, colour, thickness=1):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def vline(c, x, y_top, h, colour, thickness=1):
    c.setStrokeColor(colour)
    c.setLineWidth(thickness)
    c.line(x, rl_y(y_top), x, rl_y(y_top + h))


def stat_card(c, x, y_top, w, h, number_text, label_text, source_text):
    """Black tile, teal top accent bar, amber number, off-white label, grey source."""
    fill_rect(c, x, y_top, w, h, C_BLACK)
    fill_rect(c, x, y_top, w, 3, C_TEAL)
    # Number
    c.setFillColor(C_AMBER_B)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(x + 6, rl_y(y_top + 30), number_text)
    # Label (wrapped)
    txt_wrapped(c, label_text, x + 6, y_top + 32, w - 12, 9, C_OFF_WHITE,
                font="Helvetica", leading=12)
    # Source
    txt_wrapped(c, source_text, x + 6, y_top + 74, w - 12, 6.5, C_MID_GREY,
                font="Helvetica-Oblique", leading=9)


# ── Output ────────────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_NationalMedia_30Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("National Media | ProfitPulse Brief | 30 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Fractional CFO Partnership")


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)           # white background
fill_rect(c, 0, 0, 7, H, C_AMBER_D)         # left amber stripe

# Black header band
fill_rect(c, 7, 0, W - 7, 56, C_BLACK)

# Eyebrow in header
txt(c, "COMMERCIAL INTELLIGENCE BRIEF", 18, 12, 10, C_OFF_WHITE, font="Helvetica-Bold")

# PROFITPULSE right in header
txt(c, "PROFITPULSE", 700, 12, 10, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=248)

# Company name
c.setFillColor(C_BLACK)
c.setFont("Helvetica-Bold", 34)
c.drawString(18, rl_y(56 + 38), "National Media")

# Descriptor
txt(c, "B2B exhibitions and events company  |  Bundall, Gold Coast QLD",
    18, 102, 11, C_DRK_GREY, font="Helvetica")

# VERIFIED METRICS label
txt(c, "VERIFIED METRICS", 18, 118, 10, C_TEAL, font="Helvetica-Bold")

# Six stat cards across
card_w = 152
card_h = 102
card_y = 132
card_x_start = 18
gap = 4

card_data = [
    ("$18.4M", "Annual Revenue", "Smart50 2025 citation, Nov 2025"),
    ("42%",    "3 Year Avg Growth", "Smart50 2025, published Nov 2025"),
    ("48",     "Employees", "Smart50 2025 profile"),
    ("13",     "National Events", "Exhibition Industry News, Nov 2025"),
    ("#28",    "Smart50 2025", "SmartCompany Smart50 2025"),
    ("#62",    "AFR Fast 100 2025", "Australian Financial Review, Nov 2025"),
]
for i, (num, lbl, src) in enumerate(card_data):
    cx = card_x_start + i * (card_w + gap)
    stat_card(c, cx, card_y, card_w, card_h, num, lbl, src)

# KEY COMMERCIAL SIGNALS label
txt(c, "KEY COMMERCIAL SIGNALS", 18, 242, 10, C_TEAL, font="Helvetica-Bold")
hline(c, 18, 255, 924, C_TEAL, 0.5)

signals = [
    "Revenue doubled in the past year through organic growth and three acquisition programmes  (Exhibition Industry News, Nov 2025)",
    "Three acquisitions confirmed in 2024 to 2025: Workplace Health and Safety Shows, Foodservice Australia portfolio, and Cafe Culture and Cafe Biz Expo",
    "Dual award recognition: Smart50 rank 28 and AFR Fast 100 rank 62 confirmed in the same year  (SmartCompany and AFR, Nov 2025)",
    "13 national events across 10 brands spanning food service, hospitality, safety, fitness, architecture, and accommodation  (Exhibition Industry News)",
    "Founded 1993, over 30 years in B2B exhibitions, now among Australia's largest independent event organisers  (nationalmedia.com.au)",
]
sy = 262
for sig in signals:
    txt_wrapped(c, sig, 22, sy, 916, 9.5, C_BLACK, font="Helvetica", leading=13)
    sy += 39

# Footer
hline(c, 18, 518, 924, C_TEAL, 0.5)
txt(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse  |  Profit-Pulse.com.au",
    18, 523, 8, C_DRK_GREY, font="Helvetica")
txt(c, "30 Jun 2026", 700, 523, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=242)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
fill_rect(c, 7, 0, W - 7, 56, C_BLACK)

txt(c, "THE OPPORTUNITY", 18, 12, 10, C_AMBER_B, font="Helvetica-Bold")
txt(c, "PROFITPULSE", 700, 12, 10, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=248)
txt(c, "National Media  |  Three commercial observations from ProfitPulse",
    18, 38, 11, C_OFF_WHITE, font="Helvetica")

# Three columns
col_w = 300
col_h = 460
col_y = 62
gap2  = 8
col_fills  = [C_TEAL, C_BLACK, C_GOLD]
idx_cols   = [C_WHITE, C_AMBER_B, C_BLACK]
hdr_cols   = [C_WHITE, C_AMBER_B, C_BLACK]
body_cols  = [C_WHITE, C_OFF_WHITE, C_BLACK]

observations = [
    {
        "index": "01",
        "header": "Revenue doubled in one year. The financial architecture needs to keep pace.",
        "body": (
            "National Media grew from roughly $9 million to $18.4 million in a single year. "
            "That kind of velocity is remarkable. It also means the financial reporting cadence, "
            "management pack structure, and capital allocation framework were likely built for a "
            "smaller business. At $18.4 million across 10 brands, the P and L story becomes harder "
            "to read without a clear framework. ProfitPulse works with growing businesses to build "
            "that reporting layer before complexity outpaces visibility."
        ),
    },
    {
        "index": "02",
        "header": "Three acquisitions in 18 months. Integration economics need a disciplined approach.",
        "body": (
            "National Media executed three acquisition programmes: the Workplace Health and Safety "
            "Shows, the Foodservice Australia portfolio from Specialised Events, and Cafe Culture "
            "and Cafe Biz Expo. Each carries its own cost base, revenue model, and integration risk. "
            "Without a framework for modelling acquisition returns and tracking integration milestones, "
            "the risk is that acquisitions look successful by revenue but unclear by margin. A senior "
            "financial partner with transaction experience closes that gap."
        ),
    },
    {
        "index": "03",
        "header": "10 brands across 6 sectors. Which exhibitions deliver the real margin?",
        "body": (
            "National Media runs exhibitions in food service, hospitality, workplace safety, fitness, "
            "architecture, and accommodation. Shows vary in size, exhibitor mix, and operational cost. "
            "At this portfolio scale, the critical question is which shows generate the highest margin "
            "after full cost allocation and which consume disproportionate team capacity. "
            "The Fractional CFO Partnership builds the monthly reporting infrastructure to answer that "
            "question, enabling faster and more capital-efficient decisions across the portfolio."
        ),
    },
]

for i, obs in enumerate(observations):
    cx = 18 + i * (col_w + gap2)
    fill_rect(c, cx, col_y, col_w, col_h, col_fills[i])
    # Index
    c.setFillColor(idx_cols[i])
    c.setFont("Helvetica-Bold", 28)
    c.drawString(cx + 10, rl_y(col_y + 38), obs["index"])
    # Header
    hy = col_y + 50
    txt_wrapped(c, obs["header"], cx + 10, hy, col_w - 20, 11, hdr_cols[i],
                font="Helvetica-Bold", leading=15)
    # Body
    by = hy + 60
    txt_wrapped(c, obs["body"], cx + 10, by, col_w - 20, 9.5, body_cols[i],
                font="Helvetica", leading=14)

# Closing warm line
closing = (
    "These are observations offered in good faith. National Media has built something genuinely impressive. "
    "The question is simply whether the financial architecture matches the ambition."
)
txt_wrapped(c, closing, 18, 528, 924, 8, C_DRK_GREY,
            font="Helvetica-Oblique", leading=11)


# Footer
hline(c, 18, 518, 924, C_TEAL, 0.5)
txt(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  Profit-Pulse.com.au",
    18, 523, 8, C_DRK_GREY, font="Helvetica")
txt(c, "30 Jun 2026", 700, 523, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=242)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
fill_rect(c, 0, 0, W, H, C_WHITE)
fill_rect(c, 0, 0, 7, H, C_AMBER_D)
fill_rect(c, 7, 0, W - 7, 56, C_BLACK)

txt(c, "THE RECOMMENDATION", 18, 12, 10, C_AMBER_B, font="Helvetica-Bold")
txt(c, "PROFITPULSE", 700, 12, 10, C_TEAL, font="Helvetica-Bold",
    align="right", max_width=248)
txt(c, "National Media", 18, 38, 11, C_OFF_WHITE, font="Helvetica")

# ── LEFT COLUMN (recommendation) ──────────────────────────────────────────────
lx  = 18
lw  = 560
ly  = 66

# Service name
c.setFillColor(C_BLACK)
c.setFont("Helvetica-Bold", 19)
c.drawString(lx, rl_y(ly + 22), "Fractional CFO Partnership")

# Price line (no tier name)
txt(c, "$4,950 per month  |  ProfitPulse verified price", lx, ly + 28, 12, C_TEAL,
    font="Helvetica")

# Description
desc = (
    "A senior financial partner at the table on a monthly cadence. Includes monthly "
    "management pack, quarterly board grade review, ad hoc decision support, and a "
    "single annual deep dive. For National Media: acquisition economics modelling, "
    "portfolio capital allocation, integration milestone tracking, and a reporting "
    "cadence that turns 10 brands into a clear P and L story."
)
txt_wrapped(c, desc, lx, ly + 46, lw, 10, C_DRK_GREY, font="Helvetica", leading=14)

hline(c, lx, ly + 118, lw, C_TEAL, 0.75)

# Step 1
txt(c, "Step one: answer a few quick questions", lx, ly + 125, 11, C_BLACK, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry", lx, ly + 142, 10, C_DRK_GREY, font="Helvetica")

# Clean questionnaire address (NO tags, per Section 6.4)
txt(c, "profit-pulse.com.au/full-suite-of-products",
    lx, ly + 160, 11, C_TEAL, font="Helvetica-Bold")

# Add hyperlink annotation to questionnaire address
aq_url = "https://profit-pulse.com.au/full-suite-of-products"
aq_y_rl = rl_y(ly + 160 + 11)
try:
    c.linkURL(aq_url, (lx, aq_y_rl, lx + 300, aq_y_rl + 14), thickness=0)
except Exception:
    pass

hline(c, lx, ly + 180, lw, C_AMBER_D, 0.75)

# Direct purchase CTA (text; hyperlinked)
txt(c, "Purchase the suggested product now to get started",
    lx, ly + 188, 11, C_AMBER_D, font="Helvetica-Bold")

# Add hyperlink annotation to CTA
stripe_url = "https://buy.stripe.com/cNibJ28SYbqwepG3833ks0q"
cta_y_rl = rl_y(ly + 188 + 11)
try:
    c.linkURL(stripe_url, (lx, cta_y_rl, lx + 340, cta_y_rl + 14), thickness=0)
except Exception:
    pass

hline(c, lx, ly + 210, lw, C_LT_GREY, 0.5)

# Prefer a conversation line
txt(c, "Prefer a conversation first?", lx, ly + 218, 10, C_DRK_GREY, font="Helvetica")
booking_url = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
txt(c, "Book a complimentary discovery call", lx, ly + 232, 10, C_TEAL, font="Helvetica-Bold")
try:
    book_y_rl = rl_y(ly + 232 + 10)
    c.linkURL(booking_url, (lx, book_y_rl, lx + 220, book_y_rl + 12), thickness=0)
except Exception:
    pass

# Supporting services note
hline(c, lx, ly + 258, lw, HexColor("#dddddd"), 0.5)
txt(c, "Supporting services: G3 Product and Service Line Profitability  |  A4 Strategic Growth Diagnostic",
    lx, ly + 268, 8, C_MID_GREY, font="Helvetica-Oblique")

# ── RIGHT COLUMN (Nitesh credibility) ────────────────────────────────────────
rx = 596
rw = 346
fill_rect(c, rx, 62, rw, 456, C_PANEL_LT)

# Name
c.setFillColor(C_BLACK)
c.setFont("Helvetica-Bold", 16)
c.drawString(rx + 12, rl_y(62 + 22), "Nitesh Roopa")

txt(c, "CA, Managing Partner  |  ProfitPulse",
    rx + 12, 62 + 26, 10, C_TEAL, font="Helvetica")

hline(c, rx + 12, 62 + 42, rw - 24, C_TEAL, 0.75)

cred_lines = [
    "16 years of commercial finance experience across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3B Cahora Bassa Hydro, Mozambique",
    "Total GRBT project value in Queensland: over AUD 10 billion",
    "QIC 2023 to 2025: Finance and Commercial Lead,",
    "  AUD 10B Gympie Road Bypass Tunnel detailed business case",
    "Nedbank CIB 2015 to 2022: Energy Finance and",
    "  Principal and Equity Finance",
    "PwC South Africa 2010 to 2014: CA traineeship,",
    "  Audit Manager, Top 40 listed clients",
]

cy = 62 + 52
for cline in cred_lines:
    txt(c, cline, rx + 12, cy, 9, C_DRK_GREY, font="Helvetica")
    cy += 13

hline(c, rx + 12, cy + 4, rw - 24, C_TEAL, 0.5)
cy += 14

contact_items = [
    ("Profit-Pulse.com.au",              C_DRK_GREY, "Helvetica"),
    ("Nitesh@Profit-Pulse.com.au",       C_TEAL,     "Helvetica"),
    ("+61 411 876 267",                  C_DRK_GREY, "Helvetica"),
    ("linkedin.com/in/nitesh-roopa-77594163", C_MID_GREY, "Helvetica-Oblique"),
]
for ctext, ccol, cfont in contact_items:
    txt(c, ctext, rx + 12, cy, 9, ccol, font=cfont)
    cy += 13

# Footer
hline(c, 18, 518, 924, C_TEAL, 0.5)
txt(c, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  Profit-Pulse.com.au",
    18, 523, 8, C_DRK_GREY, font="Helvetica")
txt(c, "30 Jun 2026", 700, 523, 8, C_DRK_GREY, font="Helvetica",
    align="right", max_width=242)


# ── Save ──────────────────────────────────────────────────────────────────────
c.save()
print(f"PDF saved: {out_path}")
