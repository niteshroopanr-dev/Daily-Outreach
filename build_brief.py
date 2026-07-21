"""
ProfitPulse Brief Builder
Target: Paire | Date: 22 Jul 2026
Three slide prospect facing deck. House style v3.3. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Brand colours
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Cambria"
SANS  = "Calibri"

W = Inches(13.333)
H = Inches(7.5)

COMPANY = "Paire"
DATE_STR = "22 Jul 2026"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


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


def text(slide, s, left, top, width, height, size, colour, font=SANS,
         bold=False, italic=False, align=PP_ALIGN.LEFT, anchor=None, wrap=True,
         line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
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
    r.text = s
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = colour
    return box


def multiline(slide, lines, left, top, width, height, font=SANS,
              align=PP_ALIGN.LEFT, line_spacing=None, space_after=None):
    """lines: list of dicts with text,size,colour,bold,italic,font"""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        if space_after is not None:
            p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = ln.get("text", "")
        r.font.name = ln.get("font", font)
        r.font.size = Pt(ln.get("size", 12))
        r.font.bold = ln.get("bold", False)
        r.font.italic = ln.get("italic", False)
        r.font.color.rgb = ln.get("colour", BLACK)
    return box


def add_hyperlink(box, url):
    """Attach a hyperlink to every run in the given textbox."""
    tf = box.text_frame
    for p in tf.paragraphs:
        for r in p.runs:
            r.hyperlink.address = url


def chrome(slide, eyebrow, right_label, page_no_label=None):
    """Fixed chrome: white body bg, amber left stripe, black header band, footer line."""
    set_bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(7.5), Inches(0.4),
         12, OFF_WHITE, font=SANS, bold=True, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, right_label, Inches(9.0), Inches(0.32), Inches(4.1), Inches(0.4),
         12, AMBER_B, font=SANS, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # footer
    rect(slide, Inches(0.35), Inches(7.03), Inches(12.6), Pt(0.75), OFF_WHITE)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.35), Inches(7.12), Inches(9.0), Inches(0.3),
         8, BLACK, font=SANS, align=PP_ALIGN.LEFT)
    text(slide, DATE_STR, Inches(10.0), Inches(7.12), Inches(2.95), Inches(0.3),
         8, BLACK, font=SANS, align=PP_ALIGN.RIGHT)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    rect(slide, left, top, width, height, BLACK)
    rect(slide, left, top, width, Pt(4), TEAL)
    text(slide, number, left + Inches(0.14), top + Inches(0.14), width - Inches(0.28), Inches(0.5),
         28, AMBER_B, font=SERIF, bold=True, align=PP_ALIGN.LEFT)
    multiline(slide, [{"text": l, "size": 10.5, "colour": OFF_WHITE} for l in label_lines],
              left + Inches(0.14), top + Inches(0.68), width - Inches(0.28), Inches(0.55),
              line_spacing=1.0)
    text(slide, source, left + Inches(0.14), top + height - Inches(0.28), width - Inches(0.28), Inches(0.22),
         7.5, OFF_WHITE, font=SANS, italic=True, align=PP_ALIGN.LEFT)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ═══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

text(s1, COMPANY, Inches(0.35), Inches(1.18), Inches(9.5), Inches(0.75),
     40, BLACK, font=SERIF, bold=True, align=PP_ALIGN.LEFT)
text(s1, "Direct to consumer apparel brand, South Melbourne VIC",
     Inches(0.35), Inches(1.92), Inches(11.0), Inches(0.32),
     12, BLACK, font=SANS, align=PP_ALIGN.LEFT)

# Stat card row: 6 cards, top 2.32, height 1.6
card_top = Inches(2.32)
card_h = Inches(1.6)
card_w = Inches(1.98)
gap = Inches(0.15)
x0 = Inches(0.35)
cards = [
    ("$9.4M", ["FY25 revenue,", "Smart50 rank"], "SmartCompany, Nov 2025"),
    ("64%",   ["FY25 revenue", "growth"],         "SmartCompany, Nov 2025"),
    ("88%",   ["FY24 revenue", "growth"],          "SmartCompany, 2024"),
    ("#15",   ["2025 Smart50", "national rank"],   "SmartCompany, Nov 2025"),
    ("2021",  ["Year founded,", "per Smart50"],    "SmartCompany, 2023"),
    ("16",    ["Full time staff,", "FY25"],         "SmartCompany, Nov 2025"),
]
for i, (num, lbl, src) in enumerate(cards):
    left = x0 + i * (card_w + gap)
    stat_card(s1, left, card_top, card_w, card_h, num, lbl, src)

# Key Commercial Signals
sig_top = Inches(4.28)
text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), sig_top, Inches(6), Inches(0.3),
     12, TEAL, font=SANS, bold=True, align=PP_ALIGN.LEFT)

signals = [
    "Ranked 15th nationally on the 2025 Smart50 fastest growing companies list. Source: SmartCompany, Nov 2025",
    "Won the Smart50 2024 Rising Star Award and the Smart50 2024 Retail Award. Source: SmartCompany, 2024",
    "Opened first bricks and mortar flagship store at QV Melbourne, January 2025. Source: Inside Retail, Jan 2025",
    "Now stocked through wholesale partners in Singapore, 12 stores, and Malaysia, 3 stores. Source: Business News Australia",
    "Featured on Shark Tank Australia, national television exposure for the brand. Source: SmartCompany",
    "Revenue grew from about $1 million in 2021 to $9.4 million in FY25. Source: SmartCompany Smart50 profiles",
]
multiline(
    s1,
    [{"text": "•  " + t, "size": 11.5, "colour": BLACK} for t in signals],
    Inches(0.35), Inches(4.62), Inches(12.6), Inches(2.25),
    line_spacing=1.28, space_after=6,
)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ═══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY", COMPANY.upper())

text(s2, "Three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.15), Inches(12.6), Inches(0.35),
     14, BLACK, font=SANS, italic=True, align=PP_ALIGN.LEFT)

col_top = Inches(1.62)
col_h = Inches(4.9)
col_w = Inches(4.31)
col_x0 = Inches(0.35)
col_gap = Inches(0.02)

observations = [
    {
        "fill": TEAL, "num_colour": BLACK, "text_colour": BLACK,
        "index": "01",
        "header": "Three growth fronts, one balance sheet",
        "body": (
            "In under two years Paire has added two Melbourne retail stores to its "
            "online business and opened its first international wholesale accounts "
            "in Singapore and Malaysia. Each channel carries its own cost structure: "
            "retail fit out and staffing, wholesale margin and freight, online "
            "customer acquisition. Without a channel by channel view, it is easy to "
            "keep funding the loudest growth story rather than the most profitable one."
        ),
    },
    {
        "fill": BLACK, "num_colour": TEAL, "text_colour": OFF_WHITE,
        "index": "02",
        "header": "Inventory is the quiet cash risk",
        "body": (
            "Apparel businesses scaling into new stores and new export markets at "
            "the same time typically see cash tied up in stock rise faster than "
            "revenue, since every new store and every new wholesale partner needs "
            "its own opening stock position. A Working Capital Unlock exercise "
            "would show exactly how much cash sits in stock today and where it can "
            "be released without starving the next store opening."
        ),
    },
    {
        "fill": GOLD, "num_colour": BLACK, "text_colour": BLACK,
        "index": "03",
        "header": "Nine times revenue growth deserves a costed plan",
        "body": (
            "Revenue has grown roughly ninefold since 2021, reaching $9.4 million "
            "in FY25 on 64 percent growth, following 88 percent growth the year "
            "before. That is the profile of a business ready for a formal twelve "
            "month growth plan, ranking retail expansion, wholesale scaling, and "
            "new product lines by expected return for the team and for any future "
            "funding partner."
        ),
    },
]

for i, obs in enumerate(observations):
    left = col_x0 + i * (col_w + col_gap)
    rect(s2, left, col_top, col_w, col_h, obs["fill"])
    text(s2, obs["index"], left + Inches(0.22), col_top + Inches(0.18), col_w - Inches(0.44), Inches(0.7),
         34, obs["num_colour"], font=SERIF, bold=True, align=PP_ALIGN.LEFT)
    text(s2, obs["header"], left + Inches(0.22), col_top + Inches(0.92), col_w - Inches(0.44), Inches(0.8),
         15, obs["text_colour"], font=SERIF, bold=True, align=PP_ALIGN.LEFT, line_spacing=1.05)
    text(s2, obs["body"], left + Inches(0.22), col_top + Inches(1.82), col_w - Inches(0.44), Inches(2.95),
         11, obs["text_colour"], font=SANS, align=PP_ALIGN.LEFT, line_spacing=1.18)

text(s2, ("These observations are offered in good faith. Paire has built something "
          "genuinely impressive in four years. The question is simply whether the "
          "operating and capital plan is keeping pace with three simultaneous growth fronts."),
     Inches(0.35), Inches(6.62), Inches(12.6), Inches(0.4),
     10.5, BLACK, font=SANS, italic=True, align=PP_ALIGN.LEFT)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ═══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION", COMPANY.upper())

left_x = Inches(0.35)
left_w = Inches(7.35)
right_x = Inches(7.95)
right_w = Inches(5.0)

text(s3, "Strategic Growth Diagnostic", left_x, Inches(1.15), left_w, Inches(0.5),
     24, BLACK, font=SERIF, bold=True, align=PP_ALIGN.LEFT)
text(s3, "$5,000 one off, ProfitPulse verified price", left_x, Inches(1.66), left_w, Inches(0.32),
     13, TEAL, font=SANS, bold=True, align=PP_ALIGN.LEFT)

text(s3, ("Maps revenue, capacity, and margin headroom across retail, wholesale, "
          "and online, then produces a twelve month growth plan with funding and "
          "capital allocation steps for each channel."),
     left_x, Inches(2.06), left_w, Inches(0.75),
     11.5, BLACK, font=SANS, align=PP_ALIGN.LEFT, line_spacing=1.15)

# Step one block
rect(s3, left_x, Inches(2.95), left_w, Inches(1.35), OFF_WHITE, line_colour=TEAL, line_w=Pt(1))
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.08), left_w - Inches(0.4), Inches(0.32),
     12.5, BLACK, font=SANS, bold=True, align=PP_ALIGN.LEFT)
text(s3, "See the solutions matched to your size and industry.",
     left_x + Inches(0.2), Inches(3.42), left_w - Inches(0.4), Inches(0.3),
     11, BLACK, font=SANS, align=PP_ALIGN.LEFT)
qbox = text(s3, "profit-pulse.com.au/services/find-your-fit",
     left_x + Inches(0.2), Inches(3.75), left_w - Inches(0.4), Inches(0.32),
     12.5, TEAL, font=SANS, bold=True, align=PP_ALIGN.LEFT)
add_hyperlink(qbox, "https://profit-pulse.com.au/services/find-your-fit/")

# Direct CTA button
btn = rect(s3, left_x, Inches(4.5), Inches(5.1), Inches(0.55), AMBER_D)
btxt = text(s3, "Purchase the suggested product now to get started",
     left_x + Inches(0.18), Inches(4.5), Inches(4.75), Inches(0.55),
     12, BLACK, font=SANS, bold=True, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
add_hyperlink(btxt, "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h")

# Prefer a conversation
text(s3, "Prefer a conversation first?", left_x, Inches(5.28), left_w, Inches(0.3),
     11.5, BLACK, font=SANS, align=PP_ALIGN.LEFT)
cbox = text(s3, "Book a complimentary discovery call", left_x, Inches(5.6), left_w, Inches(0.3),
     11.5, TEAL, font=SANS, bold=True, align=PP_ALIGN.LEFT)
add_hyperlink(cbox, "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

text(s3, "Wedge: rapid multi channel expansion without a costed capital plan across retail, wholesale, and online.",
     left_x, Inches(6.15), left_w, Inches(0.55),
     9.5, BLACK, font=SANS, italic=True, align=PP_ALIGN.LEFT, line_spacing=1.1)

# Right column: About panel
rect(s3, right_x, Inches(1.15), right_w, Inches(5.55), BLACK)
text(s3, "Nitesh Roopa", right_x + Inches(0.22), Inches(1.35), right_w - Inches(0.44), Inches(0.4),
     18, AMBER_B, font=SERIF, bold=True, align=PP_ALIGN.LEFT)
text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.22), Inches(1.75), right_w - Inches(0.44), Inches(0.3),
     11.5, OFF_WHITE, font=SANS, align=PP_ALIGN.LEFT)
rect(s3, right_x + Inches(0.22), Inches(2.12), right_w - Inches(0.44), Pt(1), TEAL)

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal USD 1.3 billion, Cahora Bassa Hydro",
    "Queensland GRBT project value over AUD 10 billion",
]
multiline(s3, [{"text": "•  " + c, "size": 10.5, "colour": OFF_WHITE} for c in cred_lines],
          right_x + Inches(0.22), Inches(2.3), right_w - Inches(0.44), Inches(1.4),
          line_spacing=1.25, space_after=5)

rect(s3, right_x + Inches(0.22), Inches(3.85), right_w - Inches(0.44), Pt(1), TEAL)

contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
multiline(s3, [{"text": c, "size": 10.5, "colour": TEAL if "@" in c or "linkedin" in c else OFF_WHITE} for c in contact_lines],
          right_x + Inches(0.22), Inches(4.05), right_w - Inches(0.44), Inches(1.4),
          line_spacing=1.3, space_after=4)

text(s3, "Brisbane, Australia", right_x + Inches(0.22), Inches(6.3), right_w - Inches(0.44), Inches(0.3),
     10, OFF_WHITE, font=SANS, align=PP_ALIGN.LEFT)

prs.save("/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_22Jul2026.pptx")
print("pptx saved")
