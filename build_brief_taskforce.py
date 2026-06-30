"""
ProfitPulse Brief Builder V3
Target: Taskforce Australia | Date: 01 Jul 2026
White background V3 house style. Brand colours only. Zero dashes.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# Brand colours only (Rule 3)
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


def set_bg(slide, colour):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = colour


def r(slide, left, top, width, height, fill, border=None, bw=None):
    s = slide.shapes.add_shape(1, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border
        if bw:
            s.line.width = bw
    else:
        s.line.fill.background()
    return s


def tb(slide, text, left, top, width, height,
       fn="Arial", fs=12, bold=False, col=BLACK,
       align=PP_ALIGN.LEFT, wrap=True, italic=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = fn
    run.font.size = Pt(fs)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = col
    return box


def mlb(slide, lines_cfg, left, top, width, height,
        fn="Arial", fs=12, col=BLACK, bold=False,
        align=PP_ALIGN.LEFT):
    """Multi-line textbox. Each item in lines_cfg is a dict or string."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf  = box.text_frame
    tf.word_wrap = True
    first = True
    for item in lines_cfg:
        if isinstance(item, str):
            cfg = dict(text=item, size=fs, colour=col, bold=bold, italic=False)
        else:
            cfg = dict(text=item.get("t",""), size=item.get("fs",fs),
                       colour=item.get("col",col), bold=item.get("bold",bold),
                       italic=item.get("italic",False))
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = fn
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def hyperlink_tb(slide, text, url, left, top, width, height,
                 fn="Arial", fs=12, bold=False, col=TEAL,
                 align=PP_ALIGN.LEFT):
    """Text box with a single hyperlinked run."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf  = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = fn
    run.font.size = Pt(fs)
    run.font.bold = bold
    run.font.color.rgb = col
    rId = slide.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    rPr = run._r.get_or_add_rPr()
    hl  = etree.SubElement(rPr, qn("a:hlinkClick"))
    hl.set(qn("r:id"), rId)
    return box


def footer(slide, date_str="01 Jul 2026"):
    r(slide, Inches(0.18), Inches(7.02), Inches(13.0), Pt(1), BLACK)
    tb(slide,
       "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
       Inches(0.25), Inches(7.1), Inches(10), Inches(0.3),
       fs=9, col=BLACK)
    tb(slide, date_str,
       Inches(11.0), Inches(7.1), Inches(2.15), Inches(0.3),
       fs=9, col=BLACK, align=PP_ALIGN.RIGHT)


def header_band(slide, left_label, right_label="PROFITPULSE"):
    r(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(0.95), BLACK)
    tb(slide, left_label,
       Inches(0.25), Inches(0.22), Inches(9), Inches(0.5),
       fn="Arial", fs=11, bold=True, col=OFF_WHITE)
    tb(slide, right_label,
       Inches(10.0), Inches(0.22), Inches(3.1), Inches(0.5),
       fn="Arial", fs=11, bold=True, col=OFF_WHITE, align=PP_ALIGN.RIGHT)


# ════════════════════════════════════════════════════════════════
# SLIDE 1  COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, WHITE)

# Amber left accent stripe
r(s1, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)

header_band(s1, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
tb(s1, "Taskforce Australia",
   Inches(0.25), Inches(1.02), Inches(9.5), Inches(0.72),
   fn="Cambria", fs=44, bold=True, col=BLACK)

# One-line descriptor
tb(s1,
   "Technology enabled property maintenance and compliance platform"
   "     Burnley, Melbourne VIC",
   Inches(0.25), Inches(1.77), Inches(9.5), Inches(0.32),
   fn="Arial", fs=12, italic=True, col=TEAL)

# Teal horizontal rule
r(s1, Inches(0.18), Inches(2.15), Inches(13.0), Pt(2), TEAL)

# ── Stat cards ────────────────────────────────────────────────
CARDS = [
    ("$12.8M",  ["Revenue", "FY2025"],            "Smart50 2025 citation, SmartCompany"),
    ("31%",     ["Growth Rate", "Smart50 2025"],   "SmartCompany Smart50 2025"),
    ("19",      ["Team", "Members"],               "Smart50 2025 citation, SmartCompany"),
    ("5,500",   ["Tradie", "Network"],             "Smart50 2025 citation, SmartCompany"),
    ("Rank 37", ["Smart50", "2025"],               "SmartCompany Smart50 2025"),
]
cw  = Inches(2.5)
ch  = Inches(1.35)
gap = Inches(0.083)
cx  = Inches(0.18)
cy  = Inches(2.22)

for big, lns, src in CARDS:
    r(s1, cx, cy, cw, ch, BLACK)
    r(s1, cx, cy, cw, Pt(3), TEAL)
    tb(s1, big, cx + Inches(0.1), cy + Inches(0.07),
       cw - Inches(0.2), Inches(0.5),
       fn="Cambria", fs=26, bold=True, col=AMBER_B)
    mlb(s1,
        [{"t": lns[0], "fs": 10, "col": OFF_WHITE},
         {"t": lns[1], "fs": 10, "col": OFF_WHITE}],
        cx + Inches(0.1), cy + Inches(0.58),
        cw - Inches(0.2), Inches(0.46))
    tb(s1, src, cx + Inches(0.1), cy + Inches(1.05),
       cw - Inches(0.2), Inches(0.27),
       fs=7, italic=True, col=OFF_WHITE)
    cx += cw + gap

# ── Key Commercial Signals ────────────────────────────────────
tb(s1, "KEY COMMERCIAL SIGNALS",
   Inches(0.25), Inches(3.7), Inches(7.5), Inches(0.3),
   fn="Arial", fs=11, bold=True, col=TEAL)

SIGNALS = [
    "Smart50 consecutive listings: 2023, rank 46 in 2024, rank 37 in 2025  |  Source: SmartCompany Smart50 2023/2024/2025",
    "Revenue grew from $9.7M to $12.8M year on year (31 per cent)  |  Source: SmartCompany Smart50 2024 and 2025",
    "RentRepair subscription launched 2023, from $59 per month  |  Source: Taskforce Australia website",
    "HousingSolutions: preferred supplier to community housing organisations  |  Source: Smart50 2025 and Taskforce website",
    "Real+ partnership announced March 2025  |  Source: Taskforce Australia website (March 2025)",
    "Telstra Best of Business Awards Victorian State Winner  |  Source: Elite Agent publication",
]
sy = Inches(4.06)
for sig in SIGNALS:
    tb(s1, sig, Inches(0.25), sy, Inches(7.55), Inches(0.36),
       fn="Arial", fs=9, col=BLACK, wrap=True)
    sy += Inches(0.37)

# ── Revenue Bar Chart ─────────────────────────────────────────
chart_x = Inches(8.15)
tb(s1, "Revenue Growth", chart_x, Inches(3.7), Inches(4.9), Inches(0.3),
   fn="Arial", fs=11, bold=True, col=TEAL)

baseline = Inches(6.55)
max_h    = Inches(2.5)
bw       = Inches(1.5)

# FY2024 bar (amber)
bh24 = Inches((9.7 / 12.8) * 2.5)
bx24 = chart_x + Inches(0.2)
r(s1, bx24, baseline - bh24, bw, bh24, AMBER_D)
tb(s1, "$9.7M", bx24, baseline - bh24 - Inches(0.3), bw, Inches(0.27),
   fn="Arial", fs=10, bold=True, col=BLACK, align=PP_ALIGN.CENTER)
tb(s1, "FY2024", bx24, baseline + Inches(0.06), bw, Inches(0.25),
   fn="Arial", fs=9, col=BLACK, align=PP_ALIGN.CENTER)

# FY2025 bar (teal)
bh25 = Inches(2.5)
bx25 = bx24 + bw + Inches(0.4)
r(s1, bx25, baseline - bh25, bw, bh25, TEAL)
tb(s1, "$12.8M", bx25, baseline - bh25 - Inches(0.3), bw, Inches(0.27),
   fn="Arial", fs=10, bold=True, col=BLACK, align=PP_ALIGN.CENTER)
tb(s1, "FY2025", bx25, baseline + Inches(0.06), bw, Inches(0.25),
   fn="Arial", fs=9, col=BLACK, align=PP_ALIGN.CENTER)

# Baseline rule
r(s1, chart_x + Inches(0.1), baseline, Inches(4.6), Pt(2), BLACK)

# Chart source
tb(s1,
   "Source: SmartCompany Smart50 2024 and 2025 award citations",
   chart_x, baseline + Inches(0.37), Inches(4.9), Inches(0.28),
   fn="Arial", fs=7.5, italic=True, col=BLACK)

footer(s1)


# ════════════════════════════════════════════════════════════════
# SLIDE 2  THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, WHITE)

r(s2, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)
r(s2, Inches(0.1), Inches(0), W - Inches(0.1), Inches(0.95), BLACK)
tb(s2, "THE OPPORTUNITY",
   Inches(0.25), Inches(0.22), Inches(7), Inches(0.5),
   fn="Arial", fs=11, bold=True, col=OFF_WHITE)
tb(s2, "PROFITPULSE",
   Inches(10.0), Inches(0.22), Inches(3.1), Inches(0.5),
   fn="Arial", fs=11, bold=True, col=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Subtitle below header
tb(s2,
   "Taskforce Australia   |   Three commercial observations from ProfitPulse",
   Inches(0.25), Inches(0.98), Inches(12.5), Inches(0.32),
   fn="Arial", fs=11, col=BLACK)

# ── Three observation columns ─────────────────────────────────
COL_TOP = Inches(1.35)
COL_BOT = Inches(6.72)
COL_H   = COL_BOT - COL_TOP
COL_W   = (W - Inches(0.1)) / 3  # equal thirds

OBS = [
    {
        "idx":   "01",
        "head":  "Three products, one financial architecture",
        "body":  (
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
        "fill":  TEAL,
        "txt":   WHITE,
    },
    {
        "idx":   "02",
        "head":  "The software investment needs an ROI number first",
        "body":  (
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
        "fill":  BLACK,
        "txt":   OFF_WHITE,
    },
    {
        "idx":   "03",
        "head":  "Subscription revenue carries a higher value multiple",
        "body":  (
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
        "fill":  GOLD,
        "txt":   BLACK,
    },
]

for i, obs in enumerate(OBS):
    cx_col = Inches(0.1) + COL_W * i
    fill   = obs["fill"]
    txt    = obs["txt"]

    r(s2, cx_col, COL_TOP, COL_W, COL_H, fill)

    # Index number
    tb(s2, obs["idx"],
       cx_col + Inches(0.18), COL_TOP + Inches(0.18),
       COL_W - Inches(0.36), Inches(0.55),
       fn="Cambria", fs=36, bold=True, col=txt)

    # Thin rule below index
    accent = WHITE if fill != WHITE else TEAL
    r(s2, cx_col + Inches(0.18), COL_TOP + Inches(0.82), COL_W - Inches(0.36), Pt(1.5),
      accent if fill in (TEAL, BLACK) else BLACK)

    # Observation header
    tb(s2, obs["head"],
       cx_col + Inches(0.18), COL_TOP + Inches(0.9),
       COL_W - Inches(0.36), Inches(0.65),
       fn="Arial", fs=13, bold=True, col=txt, wrap=True)

    # Body text
    tb(s2, obs["body"],
       cx_col + Inches(0.18), COL_TOP + Inches(1.6),
       COL_W - Inches(0.36), COL_H - Inches(1.7),
       fn="Arial", fs=10, col=txt, wrap=True)

# Warm closing line
tb(s2,
   "These are observations offered in good faith. Taskforce Australia has "
   "built something genuinely impressive. The question is simply whether "
   "the financial architecture now matches the ambition.",
   Inches(0.25), COL_BOT + Inches(0.1), Inches(12.5), Inches(0.5),
   fn="Arial", fs=11, italic=True, col=BLACK)

footer(s2)


# ════════════════════════════════════════════════════════════════
# SLIDE 3  THE RECOMMENDATION
# ════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)

r(s3, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)
r(s3, Inches(0.1), Inches(0), W - Inches(0.1), Inches(0.95), BLACK)
tb(s3, "THE RECOMMENDATION",
   Inches(0.25), Inches(0.22), Inches(9), Inches(0.5),
   fn="Arial", fs=11, bold=True, col=OFF_WHITE)
tb(s3, "PROFITPULSE",
   Inches(10.0), Inches(0.22), Inches(3.1), Inches(0.5),
   fn="Arial", fs=11, bold=True, col=OFF_WHITE, align=PP_ALIGN.RIGHT)

# ── LEFT COLUMN: Recommendation ──────────────────────────────
LX = Inches(0.28)
LW = Inches(8.0)
ly = Inches(1.1)

# Service name
tb(s3, "Strategic Growth Diagnostic",
   LX, ly, LW, Inches(0.72),
   fn="Cambria", fs=32, bold=True, col=BLACK)
ly += Inches(0.75)

# Price (no tier name per Rule 6)
tb(s3, "$5,000 one off   ·   ProfitPulse verified price",
   LX, ly, LW, Inches(0.32),
   fn="Arial", fs=13, bold=True, col=TEAL)
ly += Inches(0.38)

# Description
tb(s3,
   "A six week engagement mapping revenue, capacity, and margin headroom "
   "across all three revenue streams, producing a twelve month growth plan "
   "with capital allocation steps and three scenario modelling. Designed for "
   "Taskforce's current moment: three concurrent products, a software "
   "investment decision pending, and a team of nineteen managing $12.8M.",
   LX, ly, LW, Inches(0.9),
   fn="Arial", fs=11, col=BLACK, wrap=True)
ly += Inches(0.96)

# Thin teal rule
r(s3, LX, ly, LW, Pt(1.5), TEAL)
ly += Inches(0.12)

# Step 1 block
tb(s3, "Step one: answer a few quick questions",
   LX, ly, LW, Inches(0.35),
   fn="Arial", fs=13, bold=True, col=BLACK)
ly += Inches(0.38)

tb(s3, "See the solutions matched to your size and industry.",
   LX, ly, LW, Inches(0.28),
   fn="Arial", fs=11, col=BLACK)
ly += Inches(0.32)

# Clean URL (hyperlinked to clean address, no UTM per Section 6.4)
hyperlink_tb(s3,
             "profit-pulse.com.au/full-suite-of-products",
             "https://profit-pulse.com.au/full-suite-of-products",
             LX, ly, LW, Inches(0.3),
             fn="Arial", fs=12, bold=True, col=TEAL)
ly += Inches(0.38)

# Thin rule
r(s3, LX, ly, LW, Pt(1), BLACK)
ly += Inches(0.14)

# Direct CTA (Stripe link hidden behind text, no visible address)
hyperlink_tb(s3,
             "Purchase the suggested product now to get started",
             "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
             LX, ly, LW, Inches(0.32),
             fn="Arial", fs=12, bold=True, col=AMBER_D)
ly += Inches(0.4)

# Thin rule
r(s3, LX, ly, LW, Pt(1), BLACK)
ly += Inches(0.14)

# Booking link
tb(s3, "Prefer a conversation first?",
   LX, ly, LW, Inches(0.28),
   fn="Arial", fs=11, col=BLACK)
ly += Inches(0.3)
hyperlink_tb(s3,
             "Book a complimentary discovery call",
             "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/"
             "?ismsaljsauthenabled=true",
             LX, ly, LW, Inches(0.3),
             fn="Arial", fs=11, bold=True, col=TEAL)

# ── RIGHT COLUMN: Credibility ─────────────────────────────────
RX = Inches(8.55)
RW = Inches(4.6)
r(s3, RX, Inches(1.1), RW, Inches(5.75), BLACK)

ry = Inches(1.22)
tb(s3, "Nitesh Roopa",
   RX + Inches(0.2), ry, RW - Inches(0.4), Inches(0.55),
   fn="Cambria", fs=22, bold=True, col=AMBER_B)
ry += Inches(0.55)

tb(s3, "CA, Managing Partner",
   RX + Inches(0.2), ry, RW - Inches(0.4), Inches(0.28),
   fn="Arial", fs=12, col=WHITE)
ry += Inches(0.28)

tb(s3, "ProfitPulse",
   RX + Inches(0.2), ry, RW - Inches(0.4), Inches(0.32),
   fn="Arial", fs=14, bold=True, col=TEAL)
ry += Inches(0.38)

r(s3, RX + Inches(0.2), ry, RW - Inches(0.4), Pt(1.5), TEAL)
ry += Inches(0.16)

CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal: USD 1.3 billion, Cahora Bassa Hydro, Mozambique",
    "Total GRBT project value: over AUD 10 billion",
]
for crd in CREDS:
    tb(s3, crd,
       RX + Inches(0.2), ry, RW - Inches(0.4), Inches(0.3),
       fn="Arial", fs=10, col=OFF_WHITE)
    ry += Inches(0.3)

ry += Inches(0.1)
r(s3, RX + Inches(0.2), ry, RW - Inches(0.4), Pt(1.5), TEAL)
ry += Inches(0.16)

CONTACT = [
    ("Profit-Pulse.com.au",                   OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au",            TEAL),
    ("+61 411 876 267",                       OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", OFF_WHITE),
]
for ctext, ccol in CONTACT:
    tb(s3, ctext, RX + Inches(0.2), ry, RW - Inches(0.4), Inches(0.26),
       fn="Arial", fs=11, col=ccol)
    ry += Inches(0.27)

footer(s3)


# ════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════
OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Jul2026.pptx"
prs.save(OUT)
print(f"PPTX saved: {OUT}")
