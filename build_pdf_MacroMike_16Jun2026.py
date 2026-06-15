"""
ProfitPulse Brief PDF Builder
Target: Macro Mike | Date: 16 Jun 2026
Three-page PDF matching the PPTX brief. Brand colours only. Zero dashes.
Page size: 960pt x 540pt (widescreen, 13.333in x 7.5in at 72 dpi).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Page dimensions ──────────────────────────────────────────────────────────
W = 960   # pt
H = 540   # pt

# ── Brand colours (Rule 3) ───────────────────────────────────────────────────
C_BLK  = HexColor("#000000")
C_TEL  = HexColor("#01A296")
C_AMB  = HexColor("#F8C806")
C_AMD  = HexColor("#F6A102")
C_GLD  = HexColor("#E3A712")
C_WHT  = HexColor("#FFFFFF")
C_OFW  = HexColor("#E6E5DE")

DATE = "16 Jun 2026"
CO   = "Macro Mike"

QUESTIONNAIRE_URL = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_URL        = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_URL       = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

# ── Helpers ──────────────────────────────────────────────────────────────────

def rl(y_screen):
    """Convert top-origin y to ReportLab bottom-origin y."""
    return H - y_screen


def frect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl(y_top + h), w, h, fill=1, stroke=0)


def hline(c, x, y_top, w, colour, lw=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.line(x, rl(y_top), x + w, rl(y_top))


def vline(c, x, y_top, h, colour, lw=1.5):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.line(x, rl(y_top), x, rl(y_top + h))


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold",
        align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    bl = rl(y_top + size)
    if align == "right" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(text, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, bl, text)


def txt_wrap(c, text, x, y_top, max_w, size, colour,
             font="Helvetica", leading=None):
    """Word-wrap text; returns next y_top."""
    if leading is None:
        leading = size * 1.45
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
    y = y_top
    for ln in lines:
        c.drawString(x, rl(y + size), ln)
        y += leading
    return y


def link_txt(c, text, x, y_top, size, colour, url, font="Helvetica"):
    """Draw text and attach a URL annotation."""
    c.setFillColor(colour)
    c.setFont(font, size)
    bl = rl(y_top + size)
    c.drawString(x, bl, text)
    tw = c.stringWidth(text, font, size)
    c.linkURL(url, (x, rl(y_top + size + 2), x + tw, rl(y_top - 2)),
              relative=0)


def header_band(c, eyebrow):
    frect(c, 8, 0, W - 8, 52, C_BLK)
    txt(c, eyebrow, 22, 10, 12, C_OFW, font="Helvetica-Bold")
    txt(c, "PROFITPULSE", 720, 10, 12, C_OFW, font="Helvetica-Bold",
        align="right", max_w=222)


def footer_line(c):
    hline(c, 22, 510, W - 32, C_TEL, 1)
    txt(c,
        f"Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse   |   "
        f"Profit-Pulse.com.au   |   {DATE}",
        22, 520, 8, C_OFW, font="Helvetica")


# ── Output ───────────────────────────────────────────────────────────────────
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MacroMike_16Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle(f"Macro Mike | ProfitPulse Brief | {DATE}")
c.setAuthor("ProfitPulse")
c.setSubject("Working Capital Unlock")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════════
frect(c, 0, 0, W, H, C_WHT)          # white background
frect(c, 0, 0, 8, H, C_AMD)          # amber left stripe

header_band(c, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
txt(c, "Macro Mike", 22, 62, 40, C_BLK, font="Helvetica-Bold")

# Descriptor
txt(c, "Plant based health supplement manufacturer, Burleigh Heads, Gold Coast QLD",
    22, 108, 12, C_TEL, font="Helvetica")

# ── 6 Stat cards ─────────────────────────────────────────────────────────────
cards = [
    ("Over $10M",  "Annual Revenue",      "Podcast 'The Frankie Lee', Jan 2023"),
    ("50 Staff",   "Gold Coast Facility", "Business News Australia"),
    ("100+",       "Product SKUs",        "macromike.com.au/pages/about"),
    ("5,000+",     "Retail Stockists",    "macromike.com.au/pages/about"),
    ("500,000+",   "Customers",           "macromike.com.au/pages/about"),
    ("2016",       "Year Founded",        "Pemba Capital Partners, 2025"),
]

card_w = 153   # (W-22-22-5*5) / 6 ≈ 153
card_h = 100
card_y = 128
gap    = 5
cx     = 22

for num, lbl, src in cards:
    frect(c, cx, card_y, card_w, card_h, C_BLK)
    frect(c, cx, card_y, card_w, 5, C_TEL)              # teal top stripe
    txt(c, num, cx + 8, card_y + 10, 24, C_AMB, font="Helvetica-Bold")
    txt(c, lbl, cx + 8, card_y + 44, 10, C_OFW, font="Helvetica")
    txt(c, src, cx + 8, card_y + 82, 7, C_OFW, font="Helvetica-Oblique")
    cx += card_w + gap

# Section label
txt(c, "COMMERCIAL SIGNALS", 22, 242, 11, C_TEL, font="Helvetica-Bold")

# Signals
signals = [
    ("Kellogg's collaboration: approximately 1,000 new store distribution from November 2025; "
     "supermarket push planned for May 2026",
     "Mi3 and Stack3d, October 2025"),
    ("December 2024 stock movement tripled year on year following Golden Gaytime partnership "
     "launch in October 2024",
     "Business News Australia"),
    ("Previous Golden Gaytime collaboration generated $2.5 million in sales in its first year",
     "Business News Australia"),
    ("International expansion targeted for US, UK, Europe and Asia following new 2,000sqm "
     "production facility, opened January 2023",
     "Business News Australia, June 2023"),
    ("2024 Australian Young Entrepreneur Award, Food and Beverage category, winner",
     "Business News Australia"),
    ("Official sponsor of the Gold Coast Titans NRL club since November 2021",
     "Gold Coast Titans website, November 2021"),
]

sy = 260
for sig, src in signals:
    txt_wrap(c, sig, 22, sy, 916, 10, C_BLK, font="Helvetica")
    txt(c, src, 22, sy + 14, 7, C_TEL, font="Helvetica-Oblique")
    sy += 38

footer_line(c)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2: THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
frect(c, 0, 0, W, H, C_WHT)
frect(c, 0, 0, 8, H, C_AMD)

frect(c, 8, 0, W - 8, 52, C_BLK)
txt(c, "THE OPPORTUNITY", 22, 7, 12, C_OFW, font="Helvetica-Bold")
txt(c, f"Macro Mike   |   Three commercial observations from ProfitPulse",
    22, 32, 10, C_TEL, font="Helvetica")

# Three columns
col_w = (W - 8) / 3
col_h = 435
col_y = 52
colours = [(C_TEL, C_BLK, C_AMB), (C_BLK, C_OFW, C_AMB), (C_GLD, C_BLK, C_BLK)]
indices = ["01", "02", "03"]
headings = [
    "Inventory and Retail Terms Create a Hidden Cash Gap",
    "100 SKUs Demand a Profitability Map Before the Next Retail Push",
    "International Expansion Needs a Costed Financial Architecture",
]
bodies = [
    ("Macro Mike is growing rapidly with stock movement tripling year on year in December 2024 "
     "and is now pushing into approximately 1,000 new store doors through the Kellogg's "
     "distribution launch from November 2025, with a supermarket expansion planned for May 2026. "
     "Food manufacturers carry significant inventory at every production stage. When major retail "
     "partners pay on standard 30 to 60 day terms and manufacturing costs are upfront, a working "
     "capital gap forms that grows in direct proportion to revenue. At over $10 million in annual "
     "revenue, even a 10 percent timing gap represents more than $1 million in cash pressure. "
     "Working capital discipline at this stage determines whether growth fuels the business or "
     "consumes it."),
    ("With over 100 products across protein powders, baking mixes and supplements, distributed "
     "through direct to consumer online retail, Coles, major supplement chains and 5,000 plus "
     "independent retailers, every revenue line carries a different cost to serve profile. Brand "
     "collaborations with Golden Gaytime and Kellogg's create significant revenue events but also "
     "compress margin if pricing and cost allocation are not tightly managed before the launch. "
     "At this scale, businesses that win are those that know exactly which products and channels "
     "generate the most margin after full cost allocation. Without that clarity, a major retail "
     "push risks scaling the wrong lines alongside the right ones."),
    ("Macro Mike has publicly announced plans to enter the US, UK, Europe and Asia following "
     "the completion of its new 2,000sqm Gold Coast production facility. Entering multiple "
     "international markets from a single Australian site, with a product range exceeding 100 "
     "SKUs, requires clear capital allocation, channel sequencing and cash flow staging across "
     "each market entry. The questions of how much to deploy, in which geography, through which "
     "distribution model and in which order are exactly where structured financial planning "
     "prevents expensive missteps. A costed 12 month growth plan with scenario modelling is "
     "the scaffolding international ambition of this scale requires."),
]

for i, ((bg, tc, nc), idx, hd, bd) in enumerate(zip(colours, indices, headings, bodies)):
    cx = 8 + col_w * i
    frect(c, cx, col_y, col_w, col_h, bg)
    pad = 14
    txt(c, idx, cx + pad, col_y + 8, 32, nc, font="Helvetica-Bold")
    txt_wrap(c, hd, cx + pad, col_y + 50, col_w - pad * 2, 12, tc,
             font="Helvetica-Bold", leading=17)
    txt_wrap(c, bd, cx + pad, col_y + 102, col_w - pad * 2, 10, tc,
             font="Helvetica", leading=15)

# Warm closing
txt_wrap(c,
    "These are observations offered in good faith. Macro Mike has built something genuinely "
    "impressive. The question is simply whether the financial architecture matches the ambition ahead.",
    22, 494, W - 32, 9, C_OFW, font="Helvetica-Oblique")

footer_line(c)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3: THE RECOMMENDATION AND HOW TO START
# ════════════════════════════════════════════════════════════════════════════
c.showPage()
frect(c, 0, 0, W, H, C_WHT)
frect(c, 0, 0, 8, H, C_AMD)
header_band(c, "THE RECOMMENDATION")

# ── LEFT COLUMN ──────────────────────────────────────────────────────────────
lx = 22
lw = 570

txt(c, "Working Capital Unlock", lx, 60, 28, C_TEL, font="Helvetica-Bold")
txt(c, "$6,000 one off   |   ProfitPulse verified price",
    lx, 98, 12, C_AMD, font="Helvetica-Bold")

txt_wrap(c,
    "A four-week project mapping cash trapped in inventory, work in progress, supplier terms "
    "and bank facilities, with a prioritised action list to release that cash. Typical clients "
    "release 8 to 15 percent of revenue. At Macro Mike's current scale, that represents between "
    "$800,000 and $1.5 million in released working capital.",
    lx, 118, lw, 11, C_BLK, font="Helvetica", leading=16)

# Step 1 block
frect(c, lx, 190, lw, 82, C_TEL)
txt(c, "Step one: answer a few quick questions", lx + 10, 198, 13, C_BLK, font="Helvetica-Bold")
txt(c, "See the solutions matched to your size and industry:", lx + 10, 220, 11, C_BLK, font="Helvetica")
# Clean questionnaire address (no UTM on brief, per Section 6.4)
link_txt(c, "profit-pulse.com.au/full-suite-of-products",
         lx + 10, 238, 12, C_BLK, QUESTIONNAIRE_URL, font="Helvetica-Bold")

# Purchase CTA (Stripe link hidden behind words)
frect(c, lx, 282, lw, 36, C_AMD)
link_txt(c, "Purchase the suggested product now to get started",
         lx + 10, 288, 13, C_BLK, STRIPE_URL, font="Helvetica-Bold")

# Conversation option
txt(c, "Prefer a conversation first?", lx, 334, 11, C_BLK, font="Helvetica-Bold")
link_txt(c, "Book a complimentary discovery call",
         lx, 354, 11, C_TEL, BOOKING_URL, font="Helvetica")

# Supporting services note
hline(c, lx, 384, lw, C_OFW, 0.5)
txt(c, "Supporting services: A4 Strategic Growth Diagnostic   |   G3 Product and Service Line Profitability",
    lx, 392, 9, C_TEL, font="Helvetica")

# ── RIGHT COLUMN: Credibility ─────────────────────────────────────────────
rx = 614
rw = 330

frect(c, rx, 56, rw, 434, C_BLK)
txt(c, "Nitesh Roopa", rx + 16, 68, 22, C_AMB, font="Helvetica-Bold")
txt(c, "CA, Managing Partner, ProfitPulse", rx + 16, 100, 12, C_WHT, font="Helvetica")
hline(c, rx + 16, 120, rw - 32, C_TEL, 1.5)

cred = [
    "16 years of experience across 4 countries",
    "Over 52 deals executed and managed",
    "Largest single deal: USD 1.3B Cahora Bassa Hydro",
    "Total GRBT value in Queensland: over AUD 10B",
    "CA qualification: SAICA, South Africa",
    "QIC 2023 to 2025: Finance and Commercial Lead,",
    "AUD 10B Gympie Road Bypass Tunnel",
    "Nedbank CIB 2015 to 2022: Energy Finance",
    "and Principal and Equity Finance",
]
cy = 130
for ln in cred:
    txt(c, ln, rx + 16, cy, 10, C_OFW, font="Helvetica")
    cy += 14

hline(c, rx + 16, cy + 4, rw - 32, C_TEL, 1)
cy += 14

contacts = [
    ("Profit-Pulse.com.au",               C_OFW),
    ("Nitesh@Profit-Pulse.com.au",        C_TEL),
    ("+61 411 876 267",                   C_OFW),
    ("linkedin.com/in/nitesh-roopa-77594163", C_OFW),
]
for ctxt, ccol in contacts:
    txt(c, ctxt, rx + 16, cy, 10, ccol, font="Helvetica")
    cy += 14

footer_line(c)

# ── Save ──────────────────────────────────────────────────────────────────────
c.save()
print(f"PDF saved: {out_path}")
