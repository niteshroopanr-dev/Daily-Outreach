"""
ProfitPulse Brief PDF Builder
Target: TradieSpec | Date: 17 Jun 2026
White-background house style, three slides, brand colours only, zero dashes.
Page: 960pt x 540pt (widescreen, 13.333in x 7.5in at 72dpi)
"""

from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.units import inch

pt = 1
W = 960
H = 540

# Brand colours
C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WH   = HexColor("#E6E5DE")
C_MID_GR   = HexColor("#666666")
C_DRK_GR   = HexColor("#333333")
C_MUTED    = HexColor("#999999")
C_TEAL_MID = HexColor("#016E65")
C_LIGHT_BG = HexColor("#F7F7F5")
C_TEAL_BG  = HexColor("#F2FBFA")

DATE    = "17 Jun 2026"
COMPANY = "TradieSpec"
STRIPE_URL  = "https://buy.stripe.com/5kQ9AU7OU9io95m8sn3ks0l"
QUEST_CLEAN = "profit-pulse.com.au/full-suite-of-products"
QUEST_FULL  = "https://profit-pulse.com.au/full-suite-of-products"
BOOKING_URL = ("https://bookings.cloud.microsoft/book/"
               "ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")


# ── Coordinate helpers ────────────────────────────────────────────────────────
def ry(screen_y, h=H):
    return h - screen_y


def fill(c, x, y_top, w, h_rect, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, ry(y_top + h_rect), w, h_rect, fill=1, stroke=0)


def stroke(c, x, y_top, w, h_rect, col, lw=1):
    c.setStrokeColor(col)
    c.setLineWidth(lw)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, ry(y_top + h_rect), w, h_rect, fill=0, stroke=1)


def fill_stroke(c, x, y_top, w, h_rect, fc, sc, lw=1):
    c.setFillColor(fc)
    c.setStrokeColor(sc)
    c.setLineWidth(lw)
    c.rect(x, ry(y_top + h_rect), w, h_rect, fill=1, stroke=1)


def hline(c, x, y_top, w, col, thickness=1.5):
    c.setStrokeColor(col)
    c.setLineWidth(thickness)
    c.line(x, ry(y_top), x + w, ry(y_top))


def vline(c, x, y_top, h_line, col, thickness=1.5):
    c.setStrokeColor(col)
    c.setLineWidth(thickness)
    c.line(x, ry(y_top), x, ry(y_top + h_line))


def t(c, text, x, y_top, size, colour, font="Helvetica-Bold",
      align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = ry(y_top + size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline, text)


def tw(c, text, x, y_top, max_w, size, colour, font="Helvetica", leading=None):
    """Word-wrap text. Returns next y_top."""
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
        c.drawString(x, ry(y + size), line)
        y += leading
    return y


def para_wrap(c, text, x, y_top, max_w, size, colour, font="Helvetica",
              leading=None, para_gap=12):
    """Wrap multi-paragraph text (paragraphs separated by \n\n)."""
    paras = text.split("\n\n")
    y = y_top
    for para in paras:
        y = tw(c, para, x, y, max_w, size, colour, font=font, leading=leading)
        y += para_gap
    return y


def slide_header(c, eyebrow, company):
    fill(c, 7, 0, W - 7, 50, C_BLACK)
    t(c, eyebrow, 22, 14, 11, C_OFF_WH, font="Helvetica-Bold")
    t(c, company, 680, 14, 11, C_OFF_WH, font="Helvetica-Bold",
      align="right", max_width=270)


def amber_bar(c):
    fill(c, 0, 0, 7, H, C_AMBER_D)


def slide_footer(c, date=DATE):
    hline(c, 7, 514, W - 14, C_MUTED, thickness=0.8)
    t(c,
      "Prepared by Nitesh Roopa CA, Managing Partner and Founder, "
      "ProfitPulse, Profit-Pulse.com.au",
      14, 518, 7.5, C_MID_GR, font="Helvetica")
    t(c, date, 750, 518, 7.5, C_MID_GR, font="Helvetica",
      align="right", max_width=195)


def stat_card(c, x, y_top, w, h_card, number, label1, label2, source):
    fill(c, x, y_top, w, h_card, C_BLACK)
    fill(c, x, y_top, w, 4, C_TEAL)
    t(c, number, x + 8, y_top + 12, 24, C_WHITE, font="Helvetica-Bold")
    t(c, label1, x + 8, y_top + 44, 9, C_OFF_WH, font="Helvetica")
    if label2:
        t(c, label2, x + 8, y_top + 56, 8, C_OFF_WH, font="Helvetica")
    src_y = y_top + 70 if label2 else y_top + 58
    t(c, source, x + 8, src_y, 6.5, C_TEAL_MID, font="Helvetica-Oblique")


# ═════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TradieSpec_17Jun2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("TradieSpec | ProfitPulse Brief | 17 Jun 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Capital Allocation Review")


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ═════════════════════════════════════════════════════════════════════════════
fill(c, 0, 0, W, H, C_WHITE)
amber_bar(c)
slide_header(c, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

# Company name
t(c, "TradieSpec", 14, 56, 42, C_BLACK, font="Helvetica-Bold")
# Descriptor
t(c, "Rent to own and flexible trade vehicle hire  |  Peakhurst, Sydney NSW",
  14, 106, 11, C_TEAL, font="Helvetica")

# ── Stat cards (6 cards in one row) ──────────────────────────────────────────
CARD_Y = 128
CARD_H = 92
CARD_W = 152
GAP    = 5
SX     = 14

cards = [
    ("$15.3M",  "Revenue",             "",                      "Smart50 2025, SmartCompany"),
    ("57%",     "Three year average",   "revenue growth",        "Smart50 2025, SmartCompany"),
    ("#18",     "Smart50 2025",        "rank of 50",            "SmartCompany, Nov 2025"),
    ("22",      "Team members",        "",                      "Smart50 2025, SmartCompany"),
    ("600+",    "Trade vehicles",      "in fleet",              "Company public statements"),
    ("2,500+",  "Trade businesses",    "served since 2018",     "Company public statements"),
]
for i, (num, l1, l2, src) in enumerate(cards):
    stat_card(c, SX + i * (CARD_W + GAP), CARD_Y, CARD_W, CARD_H, num, l1, l2, src)

# ── Key Commercial Signals ────────────────────────────────────────────────────
SY = 234
t(c, "KEY COMMERCIAL SIGNALS", 14, SY, 9, C_TEAL, font="Helvetica-Bold")

signals = [
    ("$13 million raised to fund national fleet expansion, with goal to double the fleet within 12 months",
     "Business Daily Media, 2023"),
    ("National operations span Sydney, Melbourne, Brisbane, Perth, Central Coast and Newcastle",
     "tradiespec.com.au"),
    ("CEO Tim Cullen publicly stated ambition to make TradieSpec Australia's largest trade vehicle provider",
     "Business Daily Media, 2023"),
    ("57% three year average growth places TradieSpec among Australia's top 20 fastest growing SMEs",
     "SmartCompany Smart50 2025"),
    ("Also ranked 9th in SmartCompany Smart50 in 2022, confirming a sustained high growth trajectory",
     "SmartCompany Smart50 2022"),
    ("Founded November 2018 and now serves over 2,500 trade businesses across four Australian states",
     "Company public statements"),
]

sy = SY + 14
for sig_text, sig_src in signals:
    fill(c, 14, sy + 3, 5, 5, C_TEAL)
    tw(c, sig_text, 24, sy, 560, 9, C_BLACK, font="Helvetica")
    t(c, "Src: " + sig_src, 592, sy + 2, 7, C_MUTED, font="Helvetica-Oblique")
    sy += 28

slide_footer(c)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THREE COMMERCIAL OBSERVATIONS
# ═════════════════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
amber_bar(c)
slide_header(c, "THE OPPORTUNITY", COMPANY)

t(c, "Three commercial observations from ProfitPulse",
  14, 56, 10, C_MID_GR, font="Helvetica-Oblique")

# Three columns
COL_Y  = 74
COL_H  = 418
COL_W  = 305
COL_G  = 9
COL_XS = [7, 7 + COL_W + COL_G, 7 + 2 * (COL_W + COL_G)]
FILLS  = [C_TEAL, C_BLACK, C_GOLD]
NUM_C  = [C_WHITE, C_WHITE, C_BLACK]
HDR_C  = [C_WHITE, C_AMBER_B, C_BLACK]
BDY_C  = [C_WHITE, C_OFF_WH, C_BLACK]

obs = [
    {
        "idx": "01",
        "hdr": "Capital allocated at scale without a return map",
        "bdy": (
            "TradieSpec has raised $13 million and deployed it across a fleet of over 600 vehicles "
            "in four capital cities. Revenue sits at $15.3 million with 22 people "
            "running operations across six locations.\n\n"
            "The central financial question is not how fast to grow but where each dollar "
            "generates the strongest return. Which markets lead on utilisation? Which "
            "vehicle categories carry the highest margin?\n\n"
            "Without a formal capital allocation map, the risk is investing in the next "
            "city rather than the best city. A Capital Allocation Review produces exactly "
            "that clarity before the next commitment."
        ),
    },
    {
        "idx": "02",
        "hdr": "Working capital dynamics compound at growth pace",
        "bdy": (
            "The rent to own model creates a layered cash flow structure that grows "
            "harder to manage as scale increases. Fleet acquisition is a large upfront "
            "outflow. Rental income spreads across over 2,500 accounts at varying "
            "intervals. Ownership transitions shift the cash profile at conversion.\n\n"
            "With $13 million in fleet financing likely carrying covenants and the "
            "business growing at 57 percent, the gap between reported profit and "
            "available cash can widen quickly without active 13 week cash flow "
            "visibility."
        ),
    },
    {
        "idx": "03",
        "hdr": "Senior financial leadership absent at a decisive stage",
        "bdy": (
            "At $15.3 million in revenue, $13 million in fleet financing, and an "
            "ambition to become Australia's largest trade vehicle provider, TradieSpec "
            "is making decisions that carry material financial consequence.\n\n"
            "Market entry sequencing, covenant management, fleet doubling, and a "
            "potential next capital raise all require senior financial thinking.\n\n"
            "With 22 employees, a full time CFO is premature. The Fractional CFO "
            "Partnership places a senior financial partner at the table on a monthly "
            "cadence, at exactly the stage where it matters most."
        ),
    },
]

for ci, ob in enumerate(obs):
    cx = COL_XS[ci]
    fill(c, cx, COL_Y, COL_W, COL_H, FILLS[ci])
    t(c, ob["idx"], cx + 10, COL_Y + 10, 26, NUM_C[ci], font="Helvetica-Bold")
    y_after = tw(c, ob["hdr"], cx + 10, COL_Y + 48, COL_W - 20,
                 12, HDR_C[ci], font="Helvetica-Bold", leading=17)
    hline(c, cx + 10, y_after + 8, COL_W - 20,
          C_AMBER_D if ci == 1 else (C_WHITE if ci == 0 else C_BLACK), 1)
    para_wrap(c, ob["bdy"], cx + 10, y_after + 18,
              COL_W - 20, 9.5, BDY_C[ci], font="Helvetica", leading=14)

# Warm closing line
t(c,
  "These are observations offered in good faith. TradieSpec has built something genuinely "
  "impressive. The question is simply whether the financial architecture matches the ambition.",
  14, 500, 8.5, C_MID_GR, font="Helvetica-Oblique")

slide_footer(c)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ═════════════════════════════════════════════════════════════════════════════
c.showPage()
fill(c, 0, 0, W, H, C_WHITE)
amber_bar(c)
slide_header(c, "THE RECOMMENDATION", COMPANY)

# ── LEFT column: recommendation and CTAs ─────────────────────────────────────
LX = 14
LW = 570

# Service name
t(c, "Capital Allocation Review", LX, 56, 18, C_BLACK, font="Helvetica-Bold")
# Price
t(c, "$7,500 one off", LX, 84, 13, C_TEAL, font="Helvetica-Bold")

# What it does
y = tw(c,
       "An independent review of where capital is deployed across the vehicle fleet, "
       "markets, and customer segments, against the return each generates. Produces a "
       "prioritised redeployment plan with expected ROI on each move, giving the team "
       "a clear investment thesis before committing the next round of fleet capital.",
       LX, 104, LW, 10, C_DRK_GR, font="Helvetica", leading=15)

hline(c, LX, y + 6, LW, C_TEAL, 1.5)

# Step one block
fill(c, LX, y + 12, LW, 80, C_TEAL_BG)
fill(c, LX, y + 12, 4, 80, C_TEAL)
t(c, "Step one: answer a few quick questions",
  LX + 12, y + 20, 11, C_BLACK, font="Helvetica-Bold")
t(c, "See the solutions matched to your size and industry.",
  LX + 12, y + 38, 10, C_DRK_GR, font="Helvetica")

# Questionnaire URL as clean visible text (no UTM on brief)
c.setFillColor(C_TEAL)
c.setFont("Helvetica", 10)
c.drawString(LX + 12, ry(y + 60 + 10), QUEST_CLEAN)
# Add clickable link annotation
link_y = ry(y + 60 + 12)
link_w = c.stringWidth(QUEST_CLEAN, "Helvetica", 10)
c.linkURL(QUEST_FULL,
           (LX + 12, link_y, LX + 12 + link_w, link_y + 12),
           relative=0)

step_bottom = y + 98

hline(c, LX, step_bottom + 4, LW, C_AMBER_D, 1.5)

# Direct CTA (Stripe link behind text, URL not shown)
cta_text = "Purchase the suggested product now to get started"
c.setFillColor(C_AMBER_D)
c.setFont("Helvetica-Bold", 11)
cta_y = step_bottom + 18
c.drawString(LX, ry(cta_y + 11), cta_text)
cta_w = c.stringWidth(cta_text, "Helvetica-Bold", 11)
cta_bottom_y = ry(cta_y + 13)
c.linkURL(STRIPE_URL,
           (LX, cta_bottom_y, LX + cta_w, cta_bottom_y + 14),
           relative=0)

t(c, "Capital Allocation Review  |  $7,500 one off  |  ProfitPulse verified price",
  LX, cta_y + 18, 8, C_MUTED, font="Helvetica-Oblique")

# Discovery call
t(c, "Prefer a conversation first?",
  LX, cta_y + 34, 10, C_DRK_GR, font="Helvetica")
bk_text = "Book a complimentary discovery call"
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 10)
bk_y = cta_y + 50
c.drawString(LX, ry(bk_y + 10), bk_text)
bk_w = c.stringWidth(bk_text, "Helvetica-Bold", 10)
c.linkURL(BOOKING_URL,
           (LX, ry(bk_y + 12), LX + bk_w, ry(bk_y + 12) + 12),
           relative=0)

t(c,
  "Supporting services: B1 Fractional CFO Partnership  |  F1 Capital Raise Feasibility",
  LX, cta_y + 68, 7.5, C_MUTED, font="Helvetica-Oblique")

# ── RIGHT column: About Nitesh ────────────────────────────────────────────────
RX = 598
RW = 350
fill(c, RX, 56, RW, 436, C_LIGHT_BG)
fill(c, RX, 56, 4, 436, C_TEAL)

t(c, "Nitesh Roopa", RX + 12, 64, 16, C_BLACK, font="Helvetica-Bold")
t(c, "CA, Managing Partner  |  ProfitPulse", RX + 12, 90, 10, C_TEAL, font="Helvetica")
hline(c, RX + 12, 106, RW - 24, C_TEAL, 1.5)

creds = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal: USD 1.3 billion,",
    "  Cahora Bassa, Mozambique, Hydro",
    "Total GRBT value in Queensland: AUD 10 billion",
    "QIC 2023 to 2025: Finance and Commercial Lead,",
    "  AUD 10B Gympie Road Bypass Tunnel",
    "Nedbank CIB 2015 to 2022: Energy Finance,",
    "  Principal and Equity Finance",
]
cy = 116
for line in creds:
    t(c, line, RX + 12, cy, 9, C_DRK_GR, font="Helvetica")
    cy += 14

hline(c, RX + 12, cy + 4, RW - 24, C_AMBER_D, 1)
cy += 14

contact = [
    ("Profit-Pulse.com.au",                    C_BLACK),
    ("Nitesh@Profit-Pulse.com.au",             C_TEAL),
    ("+61 411 876 267",                        C_BLACK),
    ("linkedin.com/in/nitesh-roopa-77594163",  C_MID_GR),
]
for ctxt, ccol in contact:
    t(c, ctxt, RX + 12, cy, 9, ccol, font="Helvetica")
    cy += 14

slide_footer(c)

c.save()
print(f"PDF saved: {out_path}")
