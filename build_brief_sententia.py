"""
ProfitPulse Brief Builder
Target: Sententia Consulting | Date for: 19 Jun 2026
Section 6 house style: white body, amber left bar, black header band.
Three slides: Commercial Intelligence Brief, The Opportunity, The Recommendation.
Zero dashes in all visible text. No tier names in any prospect-facing content.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import lxml.etree as etree

# ── Brand colours ──────────────────────────────────────────────────────────────
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
MID_GREY  = RGBColor(0x88, 0x88, 0x88)
DRK_GREY  = RGBColor(0x44, 0x44, 0x44)
LT_GREY   = RGBColor(0xCC, 0xCC, 0xCC)
TEAL_DRK  = RGBColor(0x04, 0x1A, 0x18)
GOLD_DRK  = RGBColor(0x4A, 0x33, 0x00)

W = Inches(13.333)
H = Inches(7.5)

HLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# ── Helpers ────────────────────────────────────────────────────────────────────

def set_background(slide, colour):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, fill=None, line=None, lw=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        if lw:
            shape.line.width = lw
    else:
        shape.line.fill.background()
    return shape


def txt(slide, text, left, top, width, height,
        fname="Arial", fsize=12, bold=False, italic=False,
        colour=BLACK, align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = fname
    run.font.size = Pt(fsize)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return tb, tf


def multiline(slide, lines, left, top, width, height,
              default_size=12, default_colour=OFF_WHITE,
              default_bold=False, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size,
                   "colour": default_colour, "bold": default_bold, "italic": False}
        else:
            cfg = {"text": line.get("text", ""),
                   "size": line.get("size", default_size),
                   "colour": line.get("colour", default_colour),
                   "bold": line.get("bold", default_bold),
                   "italic": line.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = "Arial"
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return tb, tf


def hyperlink_run(slide, run, url):
    """Attach a clickable hyperlink to an existing run."""
    r_id = slide.part.relate_to(url, HLINK_TYPE, is_external=True)
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        rPr = etree.Element(qn("a:rPr"))
        run._r.insert(0, rPr)
    hl = etree.SubElement(rPr, qn("a:hlinkClick"))
    hl.set(qn("r:id"), r_id)


def stat_card(slide, x, y, w, h, number, label_line1, label_line2, source):
    """Draw one stat card: black tile, teal top accent, amber number, labels."""
    STRIPE_H = Pt(4)
    rect(slide, x, y, w, STRIPE_H, fill=TEAL)
    rect(slide, x, y + STRIPE_H, w, h - STRIPE_H, fill=BLACK)
    PAD = Inches(0.12)
    txt(slide, number,
        x + PAD, y + Inches(0.05), w - PAD * 2, Inches(0.45),
        fsize=26, bold=True, colour=AMBER_B)
    txt(slide, label_line1,
        x + PAD, y + Inches(0.52), w - PAD * 2, Inches(0.22),
        fsize=9, colour=OFF_WHITE)
    txt(slide, label_line2,
        x + PAD, y + Inches(0.72), w - PAD * 2, Inches(0.22),
        fsize=9, colour=OFF_WHITE)
    txt(slide, source,
        x + PAD, y + Inches(0.94), w - PAD * 2, Inches(0.2),
        fsize=7, colour=MID_GREY, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ─────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank)
set_background(s1, WHITE)

# Amber left accent bar
rect(s1, Inches(0), Inches(0), Inches(0.1), H, fill=AMBER_D)

# Black header band
HEADER_H = Inches(0.88)
rect(s1, Inches(0.1), Inches(0), W - Inches(0.1), HEADER_H, fill=BLACK)

# Eyebrow: left
txt(s1, "COMMERCIAL INTELLIGENCE BRIEF",
    Inches(0.28), Inches(0.22), Inches(7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE)

# PROFITPULSE: right
txt(s1, "PROFITPULSE",
    Inches(9.5), Inches(0.22), Inches(3.7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Company name on white background
txt(s1, "Sententia Consulting",
    Inches(0.28), Inches(0.93), Inches(10), Inches(0.75),
    fname="Georgia", fsize=40, bold=True, colour=BLACK)

# Descriptor
txt(s1, "Government and public sector management consultancy  |  Canberra ACT",
    Inches(0.28), Inches(1.68), Inches(10), Inches(0.35),
    fsize=12, colour=TEAL)

# ── Stat cards ────────────────────────────────────────────────────────────────
CARD_Y  = Inches(2.1)
CARD_H  = Inches(1.22)
CARD_W  = Inches(2.13)
CARD_GAP = Inches(0.02)
CARD_X0  = Inches(0.28)

cards = [
    ("Over $10M",  "Annual revenue",    "surpassed",           "AFR Fast 100 2025 #45"),
    ("60.20%",     "Three year CAGR",   "compound annual",     "AFR Fast 100 2025"),
    ("#45",        "AFR Fast 100",      "rank 2025",           "AFR Fast 100 2025"),
    ("2020",       "Year founded",      "two person start",    "Company history"),
    ("50+",        "Government and",    "sector clients",      "sententiaconsulting.com.au"),
    ("#28",        "AFR Fast Starters", "rank 2022",           "AFR Fast Starters 2022"),
]
for i, (num, l1, l2, src) in enumerate(cards):
    cx = CARD_X0 + i * (CARD_W + CARD_GAP)
    stat_card(s1, cx, CARD_Y, CARD_W, CARD_H, num, l1, l2, src)

# ── Key Commercial Signals ────────────────────────────────────────────────────
SIG_Y = Inches(3.42)
txt(s1, "KEY COMMERCIAL SIGNALS",
    Inches(0.28), SIG_Y, Inches(10), Inches(0.3),
    fsize=10, bold=True, colour=TEAL)

signals = [
    ("Revenue tripled in under five years via 240% jump in government contract value",
     "Source: consultancy.com.au"),
    ("AFR Fast 100 2025 rank 45 with 60.20% CAGR, following AFR Fast Starters 2022 rank 28",
     "Source: AFR award citations"),
    ("National expansion: Melbourne office open, dedicated Managing Director appointed",
     "Source: consultancy.com.au"),
    ("Major recruitment drive including professionals from Protiviti and other firms",
     "Source: consultancy.com.au"),
    ("Active 2025 Graduate Program with applications open, sustained team scaling",
     "Source: sententiaconsulting.com.au"),
    ("50+ clients across federal, state, and territory government and sector",
     "Source: sententiaconsulting.com.au"),
]

sig_y = SIG_Y + Inches(0.35)
for sig_text, sig_src in signals:
    # Bullet dot
    rect(s1, Inches(0.28), sig_y + Inches(0.08), Inches(0.06), Inches(0.06), fill=TEAL)
    # Signal text
    txt(s1, sig_text,
        Inches(0.42), sig_y, Inches(10.5), Inches(0.28),
        fsize=11, colour=BLACK)
    # Source
    txt(s1, sig_src,
        Inches(0.42), sig_y + Inches(0.28), Inches(10.5), Inches(0.18),
        fsize=8, colour=MID_GREY, italic=True)
    sig_y += Inches(0.5)

# Footer line
rect(s1, Inches(0.28), Inches(6.95), W - Inches(0.38), Pt(1), fill=LT_GREY)

# Footer text
txt(s1,
    "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
    "Profit-Pulse.com.au",
    Inches(0.28), Inches(7.0), Inches(9), Inches(0.38),
    fsize=8, colour=DRK_GREY)
txt(s1, "19 Jun 2026",
    Inches(10.5), Inches(7.0), Inches(2.6), Inches(0.38),
    fsize=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE OPPORTUNITY
# ─────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank)
set_background(s2, WHITE)

rect(s2, Inches(0), Inches(0), Inches(0.1), H, fill=AMBER_D)

HEADER_H2 = Inches(0.88)
rect(s2, Inches(0.1), Inches(0), W - Inches(0.1), HEADER_H2, fill=BLACK)
txt(s2, "THE OPPORTUNITY",
    Inches(0.28), Inches(0.22), Inches(7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE)
txt(s2, "SENTENTIA CONSULTING",
    Inches(9.5), Inches(0.22), Inches(3.7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)

# Subtitle
txt(s2, "Sententia Consulting  |  Three commercial observations from ProfitPulse",
    Inches(0.28), Inches(0.92), Inches(12.5), Inches(0.38),
    fsize=12, colour=BLACK)

# ── Three observation columns ─────────────────────────────────────────────────
COL_Y  = Inches(1.4)
COL_H  = Inches(5.35)
COL_W  = (W - Inches(0.1) - Inches(0.26)) / 3
COL_X0 = Inches(0.28)

col_fills = [TEAL, BLACK, GOLD]
col_text_colour = [BLACK, WHITE, BLACK]
col_num_colour  = [WHITE, AMBER_B, BLACK]
col_hdr_colour  = [WHITE, AMBER_B, BLACK]

observations = [
    {
        "num":  "01",
        "hdr":  "Financial architecture has not kept pace with the growth",
        "body": (
            "Sententia has tripled its revenue on the back of government contract wins. "
            "That is the momentum. The question a senior finance professional asks next "
            "is whether the business knows which contracts produce the best margin after "
            "full cost allocation. At $10 million in government consulting revenue, "
            "the difference between a high margin engagement and a breakeven one is "
            "rarely visible without a deliberate diagnostic. "
            "The Strategic Growth Diagnostic maps exactly that."
        ),
    },
    {
        "num":  "02",
        "hdr":  "Geographic expansion introduces new structural cost questions",
        "body": (
            "A Melbourne office with its own Managing Director creates real overhead "
            "and delivery complexity. Without a financial model for the Melbourne "
            "contribution alongside Canberra, expansion can dilute overall margin "
            "in the short term before it scales. "
            "Understanding how the national footprint changes the cost base, "
            "and planning the revenue ramp required to absorb it, is the work "
            "that belongs in the next financial planning cycle. "
            "That planning cycle is what the Strategic Growth Diagnostic produces."
        ),
    },
    {
        "num":  "03",
        "hdr":  "Workforce utilisation is the lever that scales with headcount growth",
        "body": (
            "As Sententia grows its team through senior hires and a graduate cohort, "
            "revenue per person and billable utilisation become the efficiency "
            "measures that define profitability. "
            "In a professional services firm at this scale, a five percentage point "
            "improvement in utilisation across 30 people can move margin materially. "
            "Without visibility into how team time translates to revenue, "
            "headcount growth does not guarantee margin growth. "
            "The Workforce Capacity and Utilisation Review is the instrument for this."
        ),
    },
]

PAD = Inches(0.18)
for i, obs in enumerate(observations):
    cx = COL_X0 + i * COL_W
    rect(s2, cx, COL_Y, COL_W, COL_H, fill=col_fills[i])

    # Index number
    txt(s2, obs["num"],
        cx + PAD, COL_Y + Inches(0.18), COL_W - PAD * 2, Inches(0.55),
        fname="Georgia", fsize=36, bold=True, colour=col_num_colour[i])

    # Header
    multiline(s2,
        [{"text": obs["hdr"], "size": 12, "colour": col_hdr_colour[i], "bold": True}],
        cx + PAD, COL_Y + Inches(0.75), COL_W - PAD * 2, Inches(0.75),
        default_colour=col_hdr_colour[i])

    # Body
    multiline(s2,
        [{"text": obs["body"], "size": 10.5, "colour": col_text_colour[i], "bold": False}],
        cx + PAD, COL_Y + Inches(1.5), COL_W - PAD * 2, Inches(3.6),
        default_colour=col_text_colour[i])

# Warm closing line
txt(s2,
    "These observations are offered in good faith. Sententia Consulting has built "
    "something genuinely impressive. The question is simply whether the financial "
    "architecture matches the ambition.",
    Inches(0.28), Inches(6.82), Inches(12.7), Inches(0.4),
    fsize=10, italic=True, colour=DRK_GREY)

# Footer
rect(s2, Inches(0.28), Inches(7.22), W - Inches(0.38), Pt(1), fill=LT_GREY)
txt(s2, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
    Inches(0.28), Inches(7.27), Inches(9), Inches(0.2),
    fsize=8, colour=DRK_GREY)
txt(s2, "19 Jun 2026",
    Inches(10.5), Inches(7.27), Inches(2.6), Inches(0.2),
    fsize=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE RECOMMENDATION
# ─────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank)
set_background(s3, WHITE)

rect(s3, Inches(0), Inches(0), Inches(0.1), H, fill=AMBER_D)
HEADER_H3 = Inches(0.88)
rect(s3, Inches(0.1), Inches(0), W - Inches(0.1), HEADER_H3, fill=BLACK)
txt(s3, "THE RECOMMENDATION",
    Inches(0.28), Inches(0.22), Inches(7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE)
txt(s3, "SENTENTIA CONSULTING",
    Inches(9.5), Inches(0.22), Inches(3.7), Inches(0.4),
    fsize=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)

# ── LEFT COLUMN (0.28" to 8.5") ──────────────────────────────────────────────
LX = Inches(0.28)
LW = Inches(8.2)

# Service name
txt(s3, "Strategic Growth Diagnostic",
    LX, Inches(1.0), LW, Inches(0.65),
    fname="Georgia", fsize=28, bold=True, colour=BLACK)

# Price (no tier name)
txt(s3, "$5,000 one off",
    LX, Inches(1.65), LW, Inches(0.35),
    fsize=15, bold=True, colour=TEAL)

# What it does
txt(s3,
    "A six week engagement mapping revenue, capacity, and margin headroom, "
    "then producing a 12 month growth plan with funding and capital allocation "
    "steps across three scenarios. For Sententia, this clarifies which contracts "
    "earn the best margin and builds the financial model for the next growth phase.",
    LX, Inches(2.05), LW, Inches(0.75),
    fsize=11, colour=BLACK, wrap=True)

# Divider
rect(s3, LX, Inches(2.85), LW, Pt(1), fill=LT_GREY)

# Step one block
txt(s3, "STEP ONE",
    LX, Inches(2.95), LW, Inches(0.3),
    fsize=10, bold=True, colour=TEAL)

txt(s3, "Answer a few quick questions",
    LX, Inches(3.28), LW, Inches(0.3),
    fsize=13, bold=True, colour=BLACK)

txt(s3, "See the solutions matched to your size and industry.",
    LX, Inches(3.6), LW, Inches(0.28),
    fsize=11, colour=BLACK)

# Questionnaire clean URL (hyperlinked to clean URL, no UTM on brief)
CLEAN_URL = "https://profit-pulse.com.au/full-suite-of-products"
tb_url, tf_url = txt(s3, "",
    LX, Inches(3.9), LW, Inches(0.32),
    fsize=12, bold=True, colour=TEAL)
p_url = tf_url.paragraphs[0]
run_url = p_url.add_run()
run_url.text = "profit-pulse.com.au/full-suite-of-products"
run_url.font.name = "Arial"
run_url.font.size = Pt(12)
run_url.font.bold = True
run_url.font.color.rgb = TEAL
hyperlink_run(s3, run_url, CLEAN_URL)

# Purchase CTA
STRIPE_URL = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"
rect(s3, LX, Inches(4.3), LW, Inches(0.5), fill=TEAL)

tb_cta, tf_cta = txt(s3, "",
    LX + Inches(0.18), Inches(4.37), LW - Inches(0.36), Inches(0.36),
    fsize=12, bold=True, colour=WHITE)
p_cta = tf_cta.paragraphs[0]
p_cta.alignment = PP_ALIGN.CENTER
run_cta = p_cta.add_run()
run_cta.text = "Purchase the suggested product now to get started"
run_cta.font.name = "Arial"
run_cta.font.size = Pt(12)
run_cta.font.bold = True
run_cta.font.color.rgb = WHITE
hyperlink_run(s3, run_cta, STRIPE_URL)

# Prefer conversation
rect(s3, LX, Inches(4.9), LW, Pt(1), fill=LT_GREY)

txt(s3, "Prefer a conversation first?",
    LX, Inches(5.0), LW, Inches(0.28),
    fsize=11, colour=BLACK)

BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
tb_bk, tf_bk = txt(s3, "",
    LX, Inches(5.3), LW, Inches(0.28),
    fsize=11, colour=TEAL)
p_bk = tf_bk.paragraphs[0]
run_bk = p_bk.add_run()
run_bk.text = "Book a complimentary discovery call"
run_bk.font.name = "Arial"
run_bk.font.size = Pt(11)
run_bk.font.bold = False
run_bk.font.color.rgb = TEAL
hyperlink_run(s3, run_bk, BOOKING_URL)

# ── RIGHT COLUMN credibility panel ────────────────────────────────────────────
RX = Inches(8.75)
RW = Inches(4.38)
rect(s3, RX, Inches(1.0), RW, Inches(5.88), fill=BLACK)

txt(s3, "NITESH ROOPA",
    RX + Inches(0.2), Inches(1.12), RW - Inches(0.3), Inches(0.45),
    fsize=16, bold=True, colour=AMBER_B)

txt(s3, "CA, Managing Partner",
    RX + Inches(0.2), Inches(1.57), RW - Inches(0.3), Inches(0.3),
    fsize=11, colour=WHITE)

txt(s3, "ProfitPulse",
    RX + Inches(0.2), Inches(1.87), RW - Inches(0.3), Inches(0.4),
    fsize=16, bold=True, colour=TEAL)

rect(s3, RX + Inches(0.2), Inches(2.3), RW - Inches(0.4), Pt(1.5), fill=TEAL)

cred_lines = [
    {"text": "16 years across 4 countries",     "size": 10, "colour": MID_GREY, "bold": False},
    {"text": "52 deals executed and managed",   "size": 10, "colour": MID_GREY, "bold": False},
    {"text": "Largest deal: USD 1.3B Cahora Bassa", "size": 10, "colour": MID_GREY, "bold": False},
    {"text": "QLD GRBT project value AUD 10B",  "size": 10, "colour": MID_GREY, "bold": False},
    {"text": "CA qualification: SAICA South Africa", "size": 10, "colour": MID_GREY, "bold": False},
]
multiline(s3, cred_lines,
    RX + Inches(0.2), Inches(2.38), RW - Inches(0.3), Inches(1.35),
    default_colour=MID_GREY)

rect(s3, RX + Inches(0.2), Inches(3.78), RW - Inches(0.4), Pt(1), fill=RGBColor(0x33, 0x33, 0x33))

contact_lines = [
    {"text": "Profit-Pulse.com.au",          "size": 11, "colour": OFF_WHITE, "bold": False},
    {"text": "Nitesh@Profit-Pulse.com.au",   "size": 11, "colour": TEAL,      "bold": False},
    {"text": "+61 411 876 267",              "size": 11, "colour": OFF_WHITE, "bold": False},
    {"text": "",                              "size": 6,  "colour": WHITE,     "bold": False},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 9, "colour": MID_GREY, "bold": False},
    {"text": "Brisbane, Australia",           "size": 9,  "colour": MID_GREY, "bold": False},
]
multiline(s3, contact_lines,
    RX + Inches(0.2), Inches(3.86), RW - Inches(0.3), Inches(2.2),
    default_colour=OFF_WHITE)

# Footer
rect(s3, Inches(0.28), Inches(6.95), W - Inches(0.38), Pt(1), fill=LT_GREY)
txt(s3, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
    Inches(0.28), Inches(7.0), Inches(9), Inches(0.38),
    fsize=8, colour=DRK_GREY)
txt(s3, "19 Jun 2026",
    Inches(10.5), Inches(7.0), Inches(2.6), Inches(0.38),
    fsize=8, colour=DRK_GREY, align=PP_ALIGN.RIGHT)

# ── Save ───────────────────────────────────────────────────────────────────────
OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_SententiaConsulting_19Jun2026.pptx"
prs.save(OUT)
print(f"PPTX saved: {OUT}")
