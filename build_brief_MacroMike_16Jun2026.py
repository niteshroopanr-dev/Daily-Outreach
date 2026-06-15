"""
ProfitPulse Brief Builder
Target: Macro Mike | Date: 16 Jun 2026
Three-slide prospect-facing deck. Section 6 house style. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree
from pptx.oxml.ns import qn

# ── Brand colours (Rule 3, 7 approved colours only) ───────────────────────────
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

DATE = "16 Jun 2026"
CO   = "Macro Mike"

QUESTIONNAIRE_URL = "https://profit-pulse.com.au/full-suite-of-products"
STRIPE_URL        = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"
BOOKING_URL       = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# ── Helpers ───────────────────────────────────────────────────────────────────

def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


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


def tb(slide, text, l, t, w, h,
       fname="Calibri", fsize=12, bold=False,
       colour=BLACK, align=PP_ALIGN.LEFT, wrap=True, italic=False):
    bx = slide.shapes.add_textbox(l, t, w, h)
    tf = bx.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = fname
    r.font.size = Pt(fsize)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = colour
    return bx


def ml(slide, lines, l, t, w, h,
       fname="Calibri", dsize=12, dcol=BLACK, dbold=False,
       align=PP_ALIGN.LEFT, sp=None):
    bx = slide.shapes.add_textbox(l, t, w, h)
    tf = bx.text_frame
    tf.word_wrap = True
    first = True
    for ln in lines:
        if isinstance(ln, str):
            cfg = {"text": ln, "size": dsize, "colour": dcol,
                   "bold": dbold, "italic": False}
        else:
            cfg = ln
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if sp:
            p.space_after = Pt(sp)
        r = p.add_run()
        r.text = cfg.get("text", "")
        r.font.name = fname
        r.font.size = Pt(cfg.get("size", dsize))
        r.font.bold = cfg.get("bold", dbold)
        r.font.italic = cfg.get("italic", False)
        r.font.color.rgb = cfg.get("colour", dcol)
    return bx


def add_hyperlink_run(slide, txbox, para_idx, text, url,
                      fsize=11, colour=TEAL, bold=False, italic=False):
    """Append a hyperlinked run to a paragraph inside a textbox."""
    tf = txbox.text_frame
    p = tf.paragraphs[para_idx]
    r = p.add_run()
    r.text = text
    r.font.name = "Calibri"
    r.font.size = Pt(fsize)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = colour
    r.font.underline = True
    rPr = r._r.get_or_add_rPr()
    hl = etree.SubElement(rPr, qn("a:hlinkClick"))
    rId = slide.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hl.set(qn("r:id"), rId)
    return r


def header_band(slide, eyebrow):
    """Draw the standard black header band with eyebrow and PROFITPULSE on right."""
    rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(0.9), BLACK)
    tb(slide, eyebrow,
       Inches(0.22), Inches(0.12), Inches(8), Inches(0.5),
       fname="Calibri", fsize=12, bold=True, colour=OFF_WHITE)
    tb(slide, "PROFITPULSE",
       Inches(10), Inches(0.12), Inches(3.1), Inches(0.5),
       fname="Calibri", fsize=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)


def footer(slide):
    """Draw the thin footer line and text."""
    rect(slide, Inches(0.22), Inches(7.1), Inches(13.0), Inches(0.02), TEAL)
    tb(slide, f"Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse   |   Profit-Pulse.com.au   |   {DATE}",
       Inches(0.22), Inches(7.15), Inches(12.8), Inches(0.3),
       fname="Calibri", fsize=8, colour=OFF_WHITE)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, WHITE)

# Left amber accent stripe
rect(s1, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)

header_band(s1, "COMMERCIAL INTELLIGENCE BRIEF")

# Company name
tb(s1, "Macro Mike",
   Inches(0.22), Inches(1.0), Inches(9), Inches(0.75),
   fname="Georgia", fsize=44, bold=True, colour=BLACK)

# Descriptor
tb(s1, "Plant based health supplement manufacturer, Burleigh Heads, Gold Coast QLD",
   Inches(0.22), Inches(1.78), Inches(12.8), Inches(0.3),
   fname="Calibri", fsize=12, colour=TEAL)

# ── 6 Stat cards ────────────────────────────────────────────────────────────
cards = [
    ("Over $10M",  "Annual Revenue",         "Podcast 'The Frankie Lee', Jan 2023"),
    ("50 Staff",   "Gold Coast Facility",    "Business News Australia"),
    ("100+",       "Product SKUs",           "macromike.com.au/pages/about"),
    ("5,000+",     "Retail Stockists",       "macromike.com.au/pages/about"),
    ("500,000+",   "Customers",              "macromike.com.au/pages/about"),
    ("2016",       "Year Founded",           "Pemba Capital Partners, 2025"),
]

card_w   = Inches(2.10)
card_h   = Inches(1.45)
card_gap = Inches(0.05)
cx       = Inches(0.22)
cy       = Inches(2.18)

for num_txt, lbl_txt, src_txt in cards:
    rect(s1, cx, cy, card_w, card_h, BLACK)
    rect(s1, cx, cy, card_w, Inches(0.07), TEAL)          # top teal accent
    tb(s1, num_txt, cx + Inches(0.1), cy + Inches(0.1),
       card_w - Inches(0.2), Inches(0.52),
       fname="Georgia", fsize=28, bold=True, colour=AMBER_B)
    tb(s1, lbl_txt, cx + Inches(0.1), cy + Inches(0.65),
       card_w - Inches(0.2), Inches(0.45),
       fname="Calibri", fsize=11, colour=OFF_WHITE)
    tb(s1, src_txt, cx + Inches(0.1), cy + Inches(1.15),
       card_w - Inches(0.2), Inches(0.28),
       fname="Calibri", fsize=7, colour=OFF_WHITE, italic=True)
    cx += card_w + card_gap

# ── Section label ───────────────────────────────────────────────────────────
tb(s1, "COMMERCIAL SIGNALS",
   Inches(0.22), Inches(3.76), Inches(6), Inches(0.28),
   fname="Calibri", fsize=11, bold=True, colour=TEAL)

# ── 6 Verified signals ──────────────────────────────────────────────────────
signals = [
    ("Kellogg's collaboration: approximately 1,000 new store distribution launch from November 2025; supermarket push planned for May 2026",
     "Mi3 and Stack3d, October 2025"),
    ("December 2024 stock movement tripled year on year following Golden Gaytime partnership launch in October 2024",
     "Business News Australia"),
    ("Previous Golden Gaytime collaboration generated $2.5 million in sales in its first year",
     "Business News Australia"),
    ("International expansion targeted for US, UK, Europe and Asia following new 2,000sqm production facility, opened January 2023",
     "Business News Australia, June 2023"),
    ("2024 Australian Young Entrepreneur Award, Food and Beverage category, winner",
     "Business News Australia"),
    ("Official sponsor of the Gold Coast Titans NRL club since November 2021",
     "Gold Coast Titans website, November 2021"),
]

sy = Inches(4.1)
for sig_txt, sig_src in signals:
    ml(s1,
       [{"text": sig_txt, "size": 10, "colour": BLACK, "bold": False},
        {"text": sig_src,  "size":  7, "colour": TEAL,  "bold": False, "italic": True}],
       Inches(0.22), sy, Inches(12.8), Inches(0.42))
    sy += Inches(0.44)

footer(s1)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY, THREE COMMERCIAL OBSERVATIONS
# ══════════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, WHITE)

rect(s2, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)

# Header band
rect(s2, Inches(0.1), Inches(0), W - Inches(0.1), Inches(0.9), BLACK)
tb(s2, "THE OPPORTUNITY",
   Inches(0.22), Inches(0.08), Inches(6), Inches(0.38),
   fname="Calibri", fsize=12, bold=True, colour=OFF_WHITE)
tb(s2, f"Macro Mike   |   Three commercial observations from ProfitPulse",
   Inches(0.22), Inches(0.50), Inches(12.8), Inches(0.32),
   fname="Calibri", fsize=10, colour=TEAL)

# Three full-height columns (teal, black, gold)
col_w  = (W - Inches(0.1)) / 3   # ~4.411"
col_h  = Inches(5.85)
col_y  = Inches(0.9)

cols_cfg = [
    (TEAL,    BLACK,     AMBER_B,  "01",
     "Inventory and Retail Terms Create a Hidden Cash Gap",
     ("Macro Mike is growing rapidly with stock movement tripling year on year in December 2024 "
      "and is now pushing into approximately 1,000 new store doors through the Kellogg's distribution "
      "launch from November 2025, with a supermarket expansion planned for May 2026. Food "
      "manufacturers carry significant inventory at every production stage. When major retail "
      "partners pay on standard 30 to 60 day terms and manufacturing costs are upfront, a "
      "working capital gap forms that grows in direct proportion to revenue. At over $10 million "
      "in annual revenue, even a 10 percent timing gap represents more than $1 million in cash "
      "pressure. Working capital discipline at this stage determines whether growth fuels the "
      "business or consumes it. ProfitPulse specialises in exactly this challenge for Australian "
      "manufacturers at the growth inflection point.")),
    (BLACK,   OFF_WHITE, AMBER_B,  "02",
     "100 SKUs Demand a Profitability Map Before the Next Retail Push",
     ("With over 100 products across protein powders, baking mixes and supplements, distributed "
      "through direct to consumer online retail, Coles, major supplement chains and 5,000 plus "
      "independent retailers, every revenue line carries a different cost to serve profile. Brand "
      "collaborations with Golden Gaytime and Kellogg's create significant revenue events but also "
      "compress margin if pricing and cost allocation are not tightly managed before the launch. "
      "At this scale, businesses that win are those that know exactly which products and channels "
      "generate the most margin after full cost allocation. Without that clarity, a major retail "
      "push risks scaling the wrong lines as much as the right ones. A product and service line "
      "profitability review ranks every SKU by contribution, which is the insight that sharpens "
      "the next growth move.")),
    (GOLD,    BLACK,     BLACK,    "03",
     "International Expansion Needs a Costed Financial Architecture",
     ("Macro Mike has publicly announced plans to enter the US, UK, Europe and Asia following "
      "the completion of its new 2,000sqm Gold Coast production facility. Entering multiple "
      "international markets from a single Australian site, with a product range exceeding "
      "100 SKUs, requires clear capital allocation, channel sequencing and cash flow staging "
      "across each market entry. The questions of how much to deploy, in which geography, through "
      "which distribution model and in which order are exactly where structured financial planning "
      "prevents expensive missteps. A costed 12 month growth plan with scenario modelling is the "
      "financial scaffolding that international ambition of this scale requires, and the checkpoint "
      "at which capital needs and timing are defined before they become urgent.")),
]

for i, (bg, txt_col, num_col, idx, heading, body) in enumerate(cols_cfg):
    cx = Inches(0.1) + col_w * i
    rect(s2, cx, col_y, col_w, col_h, bg)
    pad = Inches(0.2)
    tb(s2, idx,
       cx + pad, col_y + Inches(0.18), col_w - pad * 2, Inches(0.5),
       fname="Georgia", fsize=36, bold=True, colour=num_col)
    tb(s2, heading,
       cx + pad, col_y + Inches(0.72), col_w - pad * 2, Inches(0.75),
       fname="Calibri", fsize=12, bold=True, colour=txt_col, wrap=True)
    tb(s2, body,
       cx + pad, col_y + Inches(1.52), col_w - pad * 2, Inches(4.0),
       fname="Calibri", fsize=10, colour=txt_col, wrap=True)

# Closing warm line
tb(s2,
   "These are observations offered in good faith. Macro Mike has built something genuinely "
   "impressive. The question is simply whether the financial architecture matches the ambition ahead.",
   Inches(0.22), Inches(6.82), Inches(13.0), Inches(0.35),
   fname="Calibri", fsize=9, colour=OFF_WHITE, italic=True)

footer(s2)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)

rect(s3, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
header_band(s3, "THE RECOMMENDATION")

# ── LEFT COLUMN: Recommendation and CTA ────────────────────────────────────
LEFT_X = Inches(0.22)
LEFT_W = Inches(7.9)

# Service name
tb(s3, "Working Capital Unlock",
   LEFT_X, Inches(1.05), LEFT_W, Inches(0.65),
   fname="Georgia", fsize=30, bold=True, colour=TEAL)

# Price (no tier name per Rule 6)
tb(s3, "$6,000 one off  |  ProfitPulse verified price",
   LEFT_X, Inches(1.72), LEFT_W, Inches(0.35),
   fname="Calibri", fsize=13, colour=AMBER_D, bold=True)

# Service description
tb(s3,
   ("A four-week project mapping cash trapped in inventory, work in progress, "
    "supplier terms and bank facilities, with a prioritised action list to release that cash. "
    "Typical clients release 8 to 15 percent of revenue. At Macro Mike's current scale, "
    "that represents between $800,000 and $1.5 million in released working capital."),
   LEFT_X, Inches(2.12), LEFT_W, Inches(0.9),
   fname="Calibri", fsize=11, colour=BLACK, wrap=True)

# Step 1 block
rect(s3, LEFT_X, Inches(3.1), LEFT_W, Inches(1.18), TEAL)
tb(s3, "Step one: answer a few quick questions",
   LEFT_X + Inches(0.15), Inches(3.18), LEFT_W - Inches(0.3), Inches(0.35),
   fname="Calibri", fsize=13, bold=True, colour=BLACK)
tb(s3, "See the solutions matched to your size and industry:",
   LEFT_X + Inches(0.15), Inches(3.55), LEFT_W - Inches(0.3), Inches(0.28),
   fname="Calibri", fsize=11, colour=BLACK)

# Questionnaire link (clean text, no UTM tags per Section 6.4, hyperlinked to clean URL)
link_bx = slide_bx = slide_tb = None
link_bx = s3.shapes.add_textbox(LEFT_X + Inches(0.15), Inches(3.85),
                                  LEFT_W - Inches(0.3), Inches(0.35))
tf_lnk = link_bx.text_frame
tf_lnk.word_wrap = False
p_lnk = tf_lnk.paragraphs[0]
r_lnk = p_lnk.add_run()
r_lnk.text = "profit-pulse.com.au/full-suite-of-products"
r_lnk.font.name  = "Calibri"
r_lnk.font.size  = Pt(12)
r_lnk.font.bold  = True
r_lnk.font.color.rgb = BLACK
r_lnk.font.underline  = True
rPr_lnk = r_lnk._r.get_or_add_rPr()
hl_lnk = etree.SubElement(rPr_lnk, qn("a:hlinkClick"))
rId_lnk = s3.part.relate_to(QUESTIONNAIRE_URL,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    is_external=True)
hl_lnk.set(qn("r:id"), rId_lnk)

# Purchase CTA (Stripe link hidden behind words, per Section 6.3)
rect(s3, LEFT_X, Inches(4.38), LEFT_W, Inches(0.55), AMBER_D)
cta_bx = s3.shapes.add_textbox(LEFT_X + Inches(0.15), Inches(4.44),
                                 LEFT_W - Inches(0.3), Inches(0.42))
tf_cta = cta_bx.text_frame
p_cta  = tf_cta.paragraphs[0]
r_cta  = p_cta.add_run()
r_cta.text = "Purchase the suggested product now to get started"
r_cta.font.name  = "Calibri"
r_cta.font.size  = Pt(13)
r_cta.font.bold  = True
r_cta.font.color.rgb = BLACK
r_cta.font.underline  = True
rPr_cta = r_cta._r.get_or_add_rPr()
hl_cta  = etree.SubElement(rPr_cta, qn("a:hlinkClick"))
rId_cta = s3.part.relate_to(STRIPE_URL,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    is_external=True)
hl_cta.set(qn("r:id"), rId_cta)

# Prefer a conversation
tb(s3, "Prefer a conversation first?",
   LEFT_X, Inches(5.08), LEFT_W, Inches(0.3),
   fname="Calibri", fsize=11, colour=BLACK, bold=True)

conv_bx = s3.shapes.add_textbox(LEFT_X, Inches(5.40), LEFT_W, Inches(0.35))
tf_conv = conv_bx.text_frame
p_conv  = tf_conv.paragraphs[0]
r_conv  = p_conv.add_run()
r_conv.text = "Book a complimentary discovery call"
r_conv.font.name  = "Calibri"
r_conv.font.size  = Pt(11)
r_conv.font.color.rgb = TEAL
r_conv.font.underline  = True
rPr_conv = r_conv._r.get_or_add_rPr()
hl_conv  = etree.SubElement(rPr_conv, qn("a:hlinkClick"))
rId_conv = s3.part.relate_to(BOOKING_URL,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    is_external=True)
hl_conv.set(qn("r:id"), rId_conv)

# ── RIGHT COLUMN: Credibility panel ────────────────────────────────────────
RIGHT_X = Inches(8.35)
RIGHT_W = Inches(4.78)

rect(s3, RIGHT_X, Inches(1.0), RIGHT_W, Inches(5.98), BLACK)

tb(s3, "Nitesh Roopa",
   RIGHT_X + Inches(0.2), Inches(1.12), RIGHT_W - Inches(0.4), Inches(0.55),
   fname="Georgia", fsize=24, bold=True, colour=AMBER_B)

tb(s3, "CA, Managing Partner, ProfitPulse",
   RIGHT_X + Inches(0.2), Inches(1.70), RIGHT_W - Inches(0.4), Inches(0.35),
   fname="Calibri", fsize=12, colour=WHITE)

rect(s3, RIGHT_X + Inches(0.2), Inches(2.10), RIGHT_W - Inches(0.4), Inches(0.02), TEAL)

cred_lines = [
    {"text": "16 years of experience across 4 countries",    "size": 11, "colour": OFF_WHITE},
    {"text": "Over 52 deals executed and managed",           "size": 11, "colour": OFF_WHITE},
    {"text": "Largest single deal: USD 1.3 billion, Cahora Bassa Hydro, Mozambique",
                                                             "size": 10, "colour": OFF_WHITE},
    {"text": "Total GRBT project value in Queensland: over AUD 10 billion",
                                                             "size": 10, "colour": OFF_WHITE},
]
ml(s3, cred_lines,
   RIGHT_X + Inches(0.2), Inches(2.18), RIGHT_W - Inches(0.4), Inches(1.5),
   fname="Calibri", dsize=11, dcol=OFF_WHITE, sp=3)

rect(s3, RIGHT_X + Inches(0.2), Inches(3.85), RIGHT_W - Inches(0.4), Inches(0.02), TEAL)

contact_lines = [
    {"text": "Profit-Pulse.com.au",           "size": 11, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au",    "size": 11, "colour": TEAL},
    {"text": "+61 411 876 267",               "size": 11, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10, "colour": OFF_WHITE},
]
ml(s3, contact_lines,
   RIGHT_X + Inches(0.2), Inches(3.92), RIGHT_W - Inches(0.4), Inches(1.3),
   fname="Calibri", dsize=11, dcol=OFF_WHITE, sp=3)

footer(s3)


# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MacroMike_16Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
