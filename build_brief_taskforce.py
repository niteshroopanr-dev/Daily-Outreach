"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 25 Jun 2026
Three-slide prospect-facing deck. Brand colours only. Zero dashes. White body background.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY  = RGBColor(0x88, 0x88, 0x88)
DRK_GREY  = RGBColor(0x44, 0x44, 0x44)
LT_GREY   = RGBColor(0x66, 0x66, 0x66)
NEAR_BLK  = RGBColor(0x11, 0x11, 0x11)
TEAL_LT   = RGBColor(0xF0, 0xFB, 0xFA)
AMBER_BG  = RGBColor(0x1A, 0x15, 0x00)

W = Inches(13.333)
H = Inches(7.5)


def set_bg(slide, colour):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = colour


def rect(slide, l, t, w, h, fill, line=None, lw=None):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        if lw:
            s.line.width = lw
    else:
        s.line.fill.background()
    return s


def txt(slide, text, l, t, w, h, size=12, bold=False, colour=BLACK,
        align=PP_ALIGN.LEFT, wrap=True, italic=False, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return box


def txt_link(slide, text, l, t, w, h, url, size=12, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.hyperlink.address = url
    return box


def multiline(slide, lines, l, t, w, h, default_size=12,
              default_colour=BLACK, default_bold=False, align=PP_ALIGN.LEFT,
              font="Calibri", spacing_after=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size,
                   "colour": default_colour, "bold": default_bold, "italic": False}
        else:
            cfg = line
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if spacing_after:
            from pptx.util import Pt as _Pt
            p.space_after = _Pt(spacing_after)
        run = p.add_run()
        run.text = cfg.get("text", "")
        run.font.name = cfg.get("font", font)
        run.font.size = Pt(cfg.get("size", default_size))
        run.font.bold = cfg.get("bold", default_bold)
        run.font.italic = cfg.get("italic", False)
        run.font.color.rgb = cfg.get("colour", default_colour)
    return box


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# =============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# =============================================================================
s1 = prs.slides.add_slide(blank)
set_bg(s1, WHITE)

# Left amber accent stripe
rect(s1, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)

# Header band
rect(s1, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
txt(s1, "COMMERCIAL INTELLIGENCE BRIEF",
    Inches(0.25), Inches(0.2), Inches(7.5), Inches(0.45),
    size=11, bold=True, colour=OFF_WHITE)
txt(s1, "TASKFORCE AUSTRALIA",
    Inches(7.8), Inches(0.2), Inches(5.3), Inches(0.45),
    size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
txt(s1, "25 Jun 2026",
    Inches(0.25), Inches(0.6), Inches(5), Inches(0.3),
    size=9, colour=MID_GREY)

# Company name
txt(s1, "Taskforce Australia",
    Inches(0.25), Inches(1.1), Inches(9), Inches(0.7),
    size=40, bold=True, colour=BLACK, font="Georgia")

# Descriptor
txt(s1, "Technology enabled property maintenance and compliance platform, Burnley, Melbourne VIC",
    Inches(0.25), Inches(1.85), Inches(12.5), Inches(0.4),
    size=12, colour=TEAL)

# Teal rule
rect(s1, Inches(0.25), Inches(2.3), Inches(12.85), Pt(2), TEAL)

# ── STAT CARDS (6 across) ─────────────────────────────────────────────────────
CARDS = [
    ("$12.8M",  "Revenue",       "FY2025",       "Smart50 Nov 2025"),
    ("31%",     "Year on year",  "growth",        "Smart50 Nov 2025"),
    ("19",      "Full time",     "team members",  "Smart50 Nov 2025"),
    ("5,500",   "Tradespeople",  "in network",    "Smart50 Nov 2025"),
    ("#37",     "Smart50 2025",  "national rank", "Smart50 Nov 2025"),
    ("500",     "Real estate",   "agency clients","taskforce.com.au"),
]
CARD_W = Inches(2.05)
CARD_H = Inches(1.38)
CARD_T = Inches(2.38)
CARD_GAP = Inches(0.05)
cx = Inches(0.25)
for num, lbl1, lbl2, src in CARDS:
    rect(s1, cx, CARD_T, CARD_W, CARD_H, BLACK)
    rect(s1, cx, CARD_T, CARD_W, Inches(0.07), TEAL)
    txt(s1, num,
        cx + Inches(0.1), CARD_T + Inches(0.1), CARD_W - Inches(0.2), Inches(0.58),
        size=26, bold=True, colour=AMBER_B, font="Georgia")
    txt(s1, lbl1 + " " + lbl2,
        cx + Inches(0.1), CARD_T + Inches(0.72), CARD_W - Inches(0.2), Inches(0.35),
        size=10, colour=OFF_WHITE)
    txt(s1, src,
        cx + Inches(0.1), CARD_T + Inches(1.1), CARD_W - Inches(0.2), Inches(0.22),
        size=7, colour=TEAL, italic=True)
    cx += CARD_W + CARD_GAP

# ── KEY COMMERCIAL SIGNALS (left column) ─────────────────────────────────────
txt(s1, "KEY COMMERCIAL SIGNALS",
    Inches(0.25), Inches(3.9), Inches(7.3), Inches(0.32),
    size=11, bold=True, colour=TEAL)

SIGNALS = [
    ("Three consecutive Smart50 placements (2023, 2024, 2025), with national ranking improving each year.",
     "SmartCompany Smart50 series 2023 to 2025."),
    ("Revenue grew from $7.43M to $12.8M in two years, a cumulative 72 percent increase.",
     "SmartCompany Smart50 2023 and Smart50 2025."),
    ("Housing division publicly named the strongest and most profitable area; social housing expansion is underway.",
     "SmartCompany Smart50 2025."),
    ("Founders have publicly committed to 50 to 60 percent CAGR for the next three years via organically funded growth.",
     "SmartCompany Smart50 2025."),
    ("Four point growth agenda includes AI platform investment and a staff equity plan, creating active capital allocation decisions.",
     "SmartCompany Smart50 2025."),
    ("RentSafe platform serves 500 real estate agencies with automated property safety compliance for rental properties.",
     "taskforce.com.au and PropTechPRO company profile."),
]
sy = Inches(4.28)
for sig_text, sig_src in SIGNALS:
    multiline(s1,
        [{"text": sig_text, "size": 9, "colour": BLACK, "bold": False},
         {"text": "Source: " + sig_src, "size": 7, "colour": MID_GREY, "italic": True, "bold": False}],
        Inches(0.25), sy, Inches(7.1), Inches(0.4),
        default_size=9, default_colour=BLACK)
    sy += Inches(0.385)

# ── REVENUE GROWTH CHART (right) ─────────────────────────────────────────────
chart_x = Inches(7.7)
chart_w = Inches(5.4)

txt(s1, "REVENUE GROWTH  (AUD)",
    chart_x, Inches(3.9), chart_w, Inches(0.32),
    size=11, bold=True, colour=TEAL)

CHART_BASE = Inches(6.5)
CHART_MAX_H = Inches(2.1)
BAR_W = Inches(1.25)
BAR_GAP = Inches(0.4)
BARS = [
    ("FY2023", 7.43, GOLD),
    ("FY2024", 9.7,  AMBER_D),
    ("FY2025", 12.8, TEAL),
]
MAX_VAL = 12.8
bx = chart_x + Inches(0.3)
for label, val, colour in BARS:
    bh = Inches(val / MAX_VAL * 2.1)
    rect(s1, bx, CHART_BASE - bh, BAR_W, bh, colour)
    txt(s1, f"${val}M",
        bx, CHART_BASE - bh - Inches(0.3), BAR_W, Inches(0.28),
        size=9, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
    txt(s1, label,
        bx, CHART_BASE + Inches(0.03), BAR_W, Inches(0.22),
        size=9, colour=DRK_GREY, align=PP_ALIGN.CENTER)
    bx += BAR_W + BAR_GAP

txt(s1, "Source: SmartCompany Smart50 award citations 2023, 2024, and 2025",
    chart_x, Inches(6.78), chart_w, Inches(0.22),
    size=7, colour=TEAL, italic=True)

# Footer rule and text
rect(s1, Inches(0.1), Inches(7.2), W - Inches(0.1), Pt(1), TEAL)
txt(s1, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
    Inches(0.25), Inches(7.22), Inches(10), Inches(0.25),
    size=8, colour=LT_GREY)
txt(s1, "25 Jun 2026",
    Inches(10.5), Inches(7.22), Inches(2.6), Inches(0.25),
    size=8, colour=LT_GREY, align=PP_ALIGN.RIGHT)


# =============================================================================
# SLIDE 2: THE OPPORTUNITY
# =============================================================================
s2 = prs.slides.add_slide(blank)
set_bg(s2, WHITE)

rect(s2, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
rect(s2, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
txt(s2, "THE OPPORTUNITY",
    Inches(0.25), Inches(0.2), Inches(7.5), Inches(0.45),
    size=11, bold=True, colour=OFF_WHITE)
txt(s2, "TASKFORCE AUSTRALIA",
    Inches(7.8), Inches(0.2), Inches(5.3), Inches(0.45),
    size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

txt(s2, "Three commercial observations from ProfitPulse",
    Inches(0.25), Inches(1.06), Inches(12.5), Inches(0.38),
    size=14, colour=DRK_GREY)

OBS = [
    {
        "idx": "01",
        "header": "Exceptional operating leverage sets the foundation",
        "body": (
            "Taskforce generates $12.8M in annual revenue with 19 full time staff, "
            "producing approximately $674K in revenue per employee. This figure is verifiable "
            "across three consecutive Smart50 award citations. A business operating at this ratio "
            "has a clearly productive core model. The next question is not whether the model works "
            "but what the financial architecture should look like to scale it intentionally and "
            "without overstretching the team or the balance sheet."
        ),
    },
    {
        "idx": "02",
        "header": "Four growth vectors competing for the same capital pool",
        "body": (
            "The four point growth agenda covers social housing expansion, a staff equity plan, "
            "AI platform development, and organic growth in the core compliance business. "
            "All four represent genuine opportunities backed by strong underlying market tailwinds "
            "in rental compliance and community housing. Without a structured capital allocation "
            "model and a scenario tested 12 month plan, the risk is that resources are spread "
            "thinly across all four vectors rather than sequenced against the highest return lever "
            "at each stage of the growth journey."
        ),
    },
    {
        "idx": "03",
        "header": "A 50% CAGR ambition needs a financial plan to match the ambition",
        "body": (
            "Jason Bright has publicly committed to a 50 to 60 percent CAGR over the next three "
            "years, implying revenue approaching $30M by FY2028 from a current base of $12.8M. "
            "That trajectory requires a clear view of margin headroom, capacity constraints, and "
            "funding requirements before capital is deployed across the four growth vectors. "
            "A costed 12 month growth plan with three scenario models is the instrument that "
            "converts that ambition into a fundable and executable path forward. This is the "
            "specific capability a Strategic Growth Diagnostic delivers."
        ),
    },
]

COL_FILLS  = [TEAL, BLACK, GOLD]
IDX_COLS   = [WHITE, AMBER_B, BLACK]
HDR_COLS   = [WHITE, AMBER_B, BLACK]
BODY_COLS  = [WHITE, OFF_WHITE, BLACK]

COL_TOP = Inches(1.55)
COL_H   = Inches(5.3)
COL_W   = Inches(4.1)
COL_GAP = Inches(0.22)

cx = Inches(0.25)
for i, obs in enumerate(OBS):
    fill  = COL_FILLS[i]
    idx_c = IDX_COLS[i]
    hdr_c = HDR_COLS[i]
    bod_c = BODY_COLS[i]
    rect(s2, cx, COL_TOP, COL_W, COL_H, fill)
    txt(s2, obs["idx"],
        cx + Inches(0.18), COL_TOP + Inches(0.2), COL_W - Inches(0.36), Inches(0.58),
        size=28, bold=True, colour=idx_c, font="Georgia")
    txt(s2, obs["header"],
        cx + Inches(0.18), COL_TOP + Inches(0.88), COL_W - Inches(0.36), Inches(0.6),
        size=12, bold=True, colour=hdr_c, wrap=True)
    txt(s2, obs["body"],
        cx + Inches(0.18), COL_TOP + Inches(1.58), COL_W - Inches(0.36), Inches(3.48),
        size=10, colour=bod_c, wrap=True)
    cx += COL_W + COL_GAP

txt(s2, (
    "These observations are offered in good faith. Taskforce Australia has built something "
    "genuinely impressive. The question is simply whether the financial architecture matches the ambition."
    ),
    Inches(0.25), Inches(6.97), Inches(12.8), Inches(0.38),
    size=9, italic=True, colour=LT_GREY)

rect(s2, Inches(0.1), Inches(7.2), W - Inches(0.1), Pt(1), TEAL)
txt(s2, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
    Inches(0.25), Inches(7.22), Inches(10), Inches(0.25),
    size=8, colour=LT_GREY)
txt(s2, "25 Jun 2026",
    Inches(10.5), Inches(7.22), Inches(2.6), Inches(0.25),
    size=8, colour=LT_GREY, align=PP_ALIGN.RIGHT)


# =============================================================================
# SLIDE 3: THE RECOMMENDATION
# =============================================================================
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)

rect(s3, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
rect(s3, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
txt(s3, "THE RECOMMENDATION",
    Inches(0.25), Inches(0.2), Inches(7.5), Inches(0.45),
    size=11, bold=True, colour=OFF_WHITE)
txt(s3, "TASKFORCE AUSTRALIA",
    Inches(7.8), Inches(0.2), Inches(5.3), Inches(0.45),
    size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# ── LEFT COLUMN ───────────────────────────────────────────────────────────────
lx = Inches(0.25)
lw = Inches(7.3)

txt(s3, "Strategic Growth Diagnostic",
    lx, Inches(1.12), lw, Inches(0.62),
    size=26, bold=True, colour=BLACK, font="Georgia")
txt(s3, "$5,000 one off",
    lx, Inches(1.77), lw, Inches(0.38),
    size=16, bold=True, colour=TEAL)

txt(s3, (
    "A six week engagement that maps your revenue, capacity, and margin headroom, then "
    "produces a 12 month growth plan with funding and capital allocation steps spelled out. "
    "Three scenario modelling. Designed for a business with a clear growth target "
    "that needs the financial architecture to support it."
    ),
    lx, Inches(2.22), lw, Inches(0.75),
    size=11, colour=RGBColor(0x33, 0x33, 0x33), wrap=True)

rect(s3, lx, Inches(3.07), lw, Pt(1.5), TEAL)

# Step 1 box
rect(s3, lx, Inches(3.15), lw, Inches(1.5), TEAL_LT, line=TEAL, lw=Pt(1))
txt(s3, "Step one: answer a few quick questions",
    lx + Inches(0.15), Inches(3.25), lw - Inches(0.3), Inches(0.4),
    size=13, bold=True, colour=TEAL)
txt(s3, "See the solutions matched to your size and industry.",
    lx + Inches(0.15), Inches(3.68), lw - Inches(0.3), Inches(0.28),
    size=11, colour=BLACK)
txt_link(s3,
    "profit-pulse.com.au/full-suite-of-products",
    lx + Inches(0.15), Inches(3.98), lw - Inches(0.3), Inches(0.3),
    url="https://profit-pulse.com.au/full-suite-of-products",
    size=11, bold=True, colour=TEAL)

# Purchase CTA (amber button)
rect(s3, lx, Inches(4.77), lw, Inches(0.52), AMBER_D)
txt_link(s3,
    "Purchase the suggested product now to get started",
    lx + Inches(0.15), Inches(4.84), lw - Inches(0.3), Inches(0.38),
    url="https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
    size=12, bold=True, colour=BLACK)

# Booking line
txt(s3, "Prefer a conversation first?",
    lx, Inches(5.43), Inches(3.0), Inches(0.3),
    size=11, colour=DRK_GREY)
txt_link(s3,
    "Book a complimentary discovery call",
    lx + Inches(3.1), Inches(5.43), Inches(4.0), Inches(0.3),
    url="https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
    size=11, bold=False, colour=TEAL)

# ── RIGHT COLUMN (credibility panel) ─────────────────────────────────────────
rx = Inches(7.85)
rw = Inches(5.25)

rect(s3, rx, Inches(1.1), rw, Inches(5.65), BLACK)

txt(s3, "Nitesh Roopa",
    rx + Inches(0.2), Inches(1.25), rw - Inches(0.4), Inches(0.52),
    size=22, bold=True, colour=AMBER_B, font="Georgia")
txt(s3, "CA, Managing Partner",
    rx + Inches(0.2), Inches(1.79), rw - Inches(0.4), Inches(0.3),
    size=13, colour=WHITE)
txt(s3, "ProfitPulse",
    rx + Inches(0.2), Inches(2.12), rw - Inches(0.4), Inches(0.42),
    size=20, bold=True, colour=TEAL)

rect(s3, rx + Inches(0.2), Inches(2.62), rw - Inches(0.4), Pt(1.5), TEAL)

CREDS = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed across career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa Hydro, Mozambique",
    "Total GRBT Queensland project value: over AUD 10 billion",
]
cy = Inches(2.77)
for cline in CREDS:
    txt(s3, cline,
        rx + Inches(0.2), cy, rw - Inches(0.4), Inches(0.36),
        size=10, colour=MID_GREY)
    cy += Inches(0.37)

rect(s3, rx + Inches(0.2), cy + Inches(0.05), rw - Inches(0.4), Pt(1), RGBColor(0x33, 0x33, 0x33))
cy += Inches(0.22)

CONTACTS = [
    ("Profit-Pulse.com.au",                OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au",         TEAL),
    ("+61 411 876 267",                    OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", MID_GREY),
]
for ctext, ccolour in CONTACTS:
    txt(s3, ctext,
        rx + Inches(0.2), cy, rw - Inches(0.4), Inches(0.3),
        size=10, colour=ccolour)
    cy += Inches(0.29)

rect(s3, Inches(0.1), Inches(7.2), W - Inches(0.1), Pt(1), TEAL)
txt(s3, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
    Inches(0.25), Inches(7.22), Inches(10), Inches(0.25),
    size=8, colour=LT_GREY)
txt(s3, "25 Jun 2026",
    Inches(10.5), Inches(7.22), Inches(2.6), Inches(0.25),
    size=8, colour=LT_GREY, align=PP_ALIGN.RIGHT)


# =============================================================================
# SAVE
# =============================================================================
out = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_25Jun2026.pptx"
prs.save(out)
print(f"PPTX saved: {out}")
