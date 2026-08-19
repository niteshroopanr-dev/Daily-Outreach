"""
ProfitPulse Brief Builder
Target: My Wealth Solutions | Date: 20 Aug 2026
Three slide prospect facing deck. House style per Section 6 of the nightly
outreach engine. Brand colours only. Zero dashes anywhere in slide text.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Brand colours, official ProfitPulse palette only
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

COMPANY = "My Wealth Solutions"
DATE_STR = "20 Aug 2026"
FOOTER_PREPARED = "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au"

QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_G1_COMMAND = "https://buy.stripe.com/bJe00kc5agKQepG0ZV3ks1x"
BOOKING_LINK = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

W = Inches(13.333)
H = Inches(7.5)

# Fixed chrome geometry, per Section 6.0A
STRIPE_W = Inches(0.1)
HEADER_H = Inches(1.0)
FOOTER_Y = Inches(7.05)
CONTENT_TOP = Inches(1.15)
CONTENT_BOTTOM = Inches(6.9)
MARGIN_L = Inches(0.35)
MARGIN_R = Inches(0.35)
CONTENT_W_IN = W.inches - MARGIN_L.inches - MARGIN_R.inches
CONTENT_W = Inches(CONTENT_W_IN)


def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, fill_colour, line_colour=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.shadow.inherit = False
    if fill_colour is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    return shp


def textbox(slide, text, left, top, width, height, size=12, bold=False, italic=False,
            colour=BLACK, align=PP_ALIGN.LEFT, font="Calibri", anchor=None,
            line_spacing=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = colour
    return tb


def multiline(slide, lines, left, top, width, height, font="Calibri",
              default_size=11, default_colour=BLACK, align=PP_ALIGN.LEFT,
              space_after=4, line_spacing=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for ln in lines:
        cfg = {"text": ln, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False} if isinstance(ln, str) else {
            "text": ln.get("text", ""), "size": ln.get("size", default_size),
            "colour": ln.get("colour", default_colour), "bold": ln.get("bold", False),
            "italic": ln.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = cfg["text"]
        r.font.name = font
        r.font.size = Pt(cfg["size"])
        r.font.bold = cfg["bold"]
        r.font.italic = cfg["italic"]
        r.font.color.rgb = cfg["colour"]
    return tb


def hyperlink_text(slide, text, url, left, top, width, height, size=12, bold=True,
                    colour=TEAL, font="Calibri", align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = colour
    r.font.underline = True
    r.hyperlink.address = url
    return tb


def chrome(slide, eyebrow, right_label, right_colour=TEAL):
    """Fixed chrome: left accent stripe, black header band, footer line."""
    set_bg(slide, WHITE)
    rect(slide, 0, 0, STRIPE_W, H, AMBER_B)
    rect(slide, 0, 0, W, HEADER_H, BLACK)
    textbox(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
            size=12, bold=True, colour=OFF_WHITE, font="Calibri")
    textbox(slide, right_label, Inches(9.0), Inches(0.32), Inches(4.0), Inches(0.4),
            size=12, bold=True, colour=right_colour, align=PP_ALIGN.RIGHT, font="Calibri")
    # footer
    rect(slide, MARGIN_L, FOOTER_Y, CONTENT_W, Pt(0.75), RGBColor(0xCC, 0xCC, 0xCC))
    textbox(slide, FOOTER_PREPARED, MARGIN_L, Inches(7.12), Inches(9.5), Inches(0.3),
            size=8, colour=RGBColor(0x66, 0x66, 0x66), font="Calibri")
    textbox(slide, DATE_STR, Inches(10.5), Inches(7.12), Inches(2.483), Inches(0.3),
            size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.RIGHT, font="Calibri")


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", COMPANY)

textbox(s1, COMPANY, MARGIN_L, Inches(1.22), Inches(9.5), Inches(0.7),
        size=40, bold=True, colour=BLACK, font="Georgia")
textbox(s1, "Financial planning and wealth advisory group, Newstead, Brisbane QLD",
        MARGIN_L, Inches(1.92), Inches(11.0), Inches(0.35),
        size=13, colour=RGBColor(0x33, 0x33, 0x33), font="Calibri")

# Stat cards: 6 cards, single row
CARD_TOP = Inches(2.4)
CARD_H = Inches(1.6)
N_CARDS = 6
GAP = Inches(0.15)
CARD_W = Inches((CONTENT_W_IN - GAP.inches * (N_CARDS - 1)) / N_CARDS)

cards = [
    ("$9.9M", "FY2025 revenue,", "Smart50 profile", "SmartCompany Smart50 2025"),
    ("47%",   "Revenue growth", "rate, Smart50", "SmartCompany Smart50 2025"),
    ("#25",   "Smart50 2025", "rank of 50", "SmartCompany Smart50 2025"),
    ("55",    "People across", "six offices", "Company website, 2026"),
    ("2011",  "Founded in", "Newstead, QLD", "Company website, our story"),
    ("3x",    "GPS Wealth Practice", "of the Year winner", "GPS Wealth conference recaps"),
]

for i, (num, l1, l2, src) in enumerate(cards):
    x = MARGIN_L + i * (CARD_W + GAP)
    rect(s1, x, CARD_TOP, CARD_W, CARD_H, BLACK)
    rect(s1, x, CARD_TOP, CARD_W, Pt(3), TEAL)
    textbox(s1, num, x + Inches(0.12), CARD_TOP + Inches(0.14), CARD_W - Inches(0.24), Inches(0.5),
            size=28, bold=True, colour=AMBER_B, font="Georgia")
    multiline(s1, [l1, l2], x + Inches(0.12), CARD_TOP + Inches(0.68),
              CARD_W - Inches(0.24), Inches(0.55), font="Calibri",
              default_size=10, default_colour=OFF_WHITE, space_after=0, line_spacing=1.0)
    textbox(s1, src, x + Inches(0.12), CARD_TOP + Inches(1.32), CARD_W - Inches(0.24), Inches(0.24),
            size=7, colour=TEAL, italic=True, font="Calibri")

# Key Commercial Signals
SIG_TOP = Inches(4.28)
textbox(s1, "KEY COMMERCIAL SIGNALS", MARGIN_L, SIG_TOP, Inches(6), Inches(0.3),
        size=12, bold=True, colour=TEAL, font="Calibri")

signals = [
    "Ranked 25th nationally, 2025 Smart50 Awards. Source: SmartCompany Smart50 2025 profile.",
    "Grew to six offices: Newstead, Sydney CBD, Neutral Bay, Gold Coast, Melbourne, Townsville. Source: company website.",
    "Named GPS Wealth Practice of the Year three years running. Source: company blog, GPS Wealth conference recaps.",
    "Recognised in the AFR Fast 100 and the 2025 ifa Excellence Awards. Source: company website, ifa Excellence Awards 2025.",
    "Founded 2011 by Guy Freeman and Ben Budge, both trained together in financial planning. Source: company Our Story page.",
    "Rated a top 1 percent Australian financial planner, Brisbane Outer Suburbs 2026. Source: Quality Business Awards Australia.",
]
y = SIG_TOP + Inches(0.38)
for sig in signals:
    textbox(s1, "•  " + sig, MARGIN_L, y, Inches(12.4), Inches(0.42),
            size=11, colour=RGBColor(0x22, 0x22, 0x22), font="Calibri")
    y += Inches(0.46)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY", COMPANY)

textbox(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
        MARGIN_L, Inches(1.2), Inches(12.4), Inches(0.4),
        size=15, bold=True, colour=BLACK, font="Georgia")

COL_TOP = Inches(1.7)
COL_H = Inches(4.55)
COL_GAP = Inches(0.15)
COL_W = Inches((CONTENT_W_IN - COL_GAP.inches * 2) / 3)

observations = [
    ("01", TEAL, BLACK,
     "Six offices, one performance picture, or not yet",
     "The practice has grown from one Newstead office in 2011 to six locations "
     "across Queensland, New South Wales, and Victoria, lifting revenue 47 "
     "percent to 9.9 million dollars in the year that earned Smart50 rank 25. "
     "Each office carries its own referral splits, servicing cost, and "
     "licensee fees. Without a location by location margin view, a quietly "
     "subsidised office can look identical to a strong one."),
    ("02", BLACK, WHITE,
     "Fifty five people is a company, not a practice",
     "Headcount has grown to 55 alongside revenue, a point where informal, "
     "founder led oversight usually starts to strain. Three consecutive GPS "
     "Wealth Practice of the Year wins confirm client outcomes are strong. "
     "What those awards do not confirm is whether cost base and workforce "
     "capacity are still returning a proportional margin as the team has "
     "grown."),
    ("03", GOLD, BLACK,
     "The award record is proof, the numbers are the next step",
     "An AFR Fast 100 mention, an ifa Excellence Awards nod, and three GPS "
     "Wealth trophies make this one of the more credentialed growth stories "
     "in Australian financial advice. That record carries real weight with a "
     "bank, a licensee, or a future partner. Businesses that convert growth "
     "stories into value are the ones holding clean, location level numbers "
     "behind the headline."),
]

for i, (idx, fill, textcol, header, para) in enumerate(observations):
    x = MARGIN_L + i * (COL_W + COL_GAP)
    rect(s2, x, COL_TOP, COL_W, COL_H, fill)
    textbox(s2, idx, x + Inches(0.22), COL_TOP + Inches(0.2), COL_W - Inches(0.44), Inches(0.6),
            size=32, bold=True, colour=(OFF_WHITE if fill == BLACK else BLACK), font="Georgia")
    textbox(s2, header, x + Inches(0.22), COL_TOP + Inches(0.95), COL_W - Inches(0.44), Inches(0.9),
            size=14, bold=True, colour=textcol, font="Calibri", line_spacing=1.05)
    textbox(s2, para, x + Inches(0.22), COL_TOP + Inches(1.85), COL_W - Inches(0.44), Inches(2.6),
            size=10.5, colour=textcol, font="Calibri", line_spacing=1.12)

CLOSE_TOP = COL_TOP + COL_H + Inches(0.15)
textbox(s2,
        "These observations are offered in good faith. My Wealth Solutions has "
        "built something genuinely impressive in fifteen years. The question is "
        "simply whether the financial architecture across six offices is keeping "
        "pace with the ambition.",
        MARGIN_L, CLOSE_TOP, Inches(12.6), Inches(0.6),
        size=11, italic=True, colour=RGBColor(0x33, 0x33, 0x33), font="Calibri")

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION", "PROFITPULSE")

LEFT_X = MARGIN_L
LEFT_W = Inches(7.4)
RIGHT_X = Inches(8.05)
RIGHT_W = Inches(4.93)

textbox(s3, "Operational Intelligence Review", LEFT_X, Inches(1.2), LEFT_W, Inches(0.55),
        size=24, bold=True, colour=BLACK, font="Georgia")
textbox(s3, "$6,500 one off", LEFT_X, Inches(1.78), LEFT_W, Inches(0.4),
        size=18, bold=True, colour=AMBER_D, font="Calibri")
textbox(s3,
        "A six week review across four lenses: customer concentration and "
        "profitability, product and service line margin, workforce capacity, "
        "and operational bottlenecks, mapped office by office and adviser by "
        "adviser.",
        LEFT_X, Inches(2.28), LEFT_W, Inches(0.85),
        size=11.5, colour=RGBColor(0x22, 0x22, 0x22), font="Calibri", line_spacing=1.15)

# Step one block
rect(s3, LEFT_X, Inches(3.28), LEFT_W, Inches(1.35), RGBColor(0xF4, 0xF4, 0xF2),
     line_colour=TEAL, line_w=Pt(1))
textbox(s3, "Step one, answer a few quick questions", LEFT_X + Inches(0.2), Inches(3.42),
        LEFT_W - Inches(0.4), Inches(0.35), size=13, bold=True, colour=TEAL, font="Calibri")
textbox(s3, "See the solutions matched to your size and industry.",
        LEFT_X + Inches(0.2), Inches(3.76), LEFT_W - Inches(0.4), Inches(0.32),
        size=11, colour=RGBColor(0x22, 0x22, 0x22), font="Calibri")
hyperlink_text(s3, QUESTIONNAIRE_CLEAN, QUESTIONNAIRE_URL, LEFT_X + Inches(0.2), Inches(4.1),
               LEFT_W - Inches(0.4), Inches(0.35), size=13, bold=True, colour=TEAL)

# Direct CTA button
btn = rect(s3, LEFT_X, Inches(4.85), Inches(4.6), Inches(0.55), None,
           line_colour=AMBER_D, line_w=Pt(1.5))
btn_tf = btn.text_frame
btn_tf.word_wrap = True
btn_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
bp = btn_tf.paragraphs[0]
bp.alignment = PP_ALIGN.CENTER
br = bp.add_run()
br.text = "Purchase the suggested product now to get started"
br.font.size = Pt(11)
br.font.bold = True
br.font.color.rgb = AMBER_D
br.font.name = "Calibri"
br.hyperlink.address = STRIPE_G1_COMMAND

textbox(s3, "Prefer a conversation first?", LEFT_X, Inches(5.65), LEFT_W, Inches(0.32),
        size=11, colour=RGBColor(0x22, 0x22, 0x22), font="Calibri")
hyperlink_text(s3, "Book a complimentary discovery call", BOOKING_LINK, LEFT_X, Inches(5.96),
               LEFT_W, Inches(0.35), size=12, bold=True, colour=TEAL)

# Right column: credibility panel
rect(s3, RIGHT_X, Inches(1.15), RIGHT_W, Inches(5.75), BLACK)
textbox(s3, "NITESH ROOPA", RIGHT_X + Inches(0.25), Inches(1.35), RIGHT_W - Inches(0.5), Inches(0.4),
        size=17, bold=True, colour=AMBER_B, font="Georgia")
textbox(s3, "CA, Managing Partner, ProfitPulse", RIGHT_X + Inches(0.25), Inches(1.78),
        RIGHT_W - Inches(0.5), Inches(0.32), size=12, colour=WHITE, font="Calibri")
rect(s3, RIGHT_X + Inches(0.25), Inches(2.18), RIGHT_W - Inches(0.5), Pt(1.5), TEAL)

cred = [
    "16 years of experience across 4 countries",
    "Over 52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique",
    "Over AUD 10 billion Gympie Road Bypass Tunnel, total project value",
]
multiline(s3, cred, RIGHT_X + Inches(0.25), Inches(2.35), RIGHT_W - Inches(0.5), Inches(1.9),
          font="Calibri", default_size=10.5, default_colour=OFF_WHITE, space_after=8, line_spacing=1.1)

rect(s3, RIGHT_X + Inches(0.25), Inches(4.35), RIGHT_W - Inches(0.5), Pt(1), RGBColor(0x33, 0x33, 0x33))

contact = [
    {"text": "Profit-Pulse.com.au", "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "colour": TEAL},
    {"text": "+61 411 876 267", "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "colour": TEAL},
]
multiline(s3, contact, RIGHT_X + Inches(0.25), Inches(4.55), RIGHT_W - Inches(0.5), Inches(1.6),
          font="Calibri", default_size=11, space_after=8)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_MyWealthSolutions_20Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
