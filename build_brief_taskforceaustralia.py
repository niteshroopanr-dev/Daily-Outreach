"""
ProfitPulse Brief Builder, Version 3.3 house style
Target: Taskforce Australia | Date: 28 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours, the seven only
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Cambria"
SANS  = "Arial"

W = Inches(13.333)
H = Inches(7.5)
DATE_STR = "28 Aug 2026"
COMPANY = "Taskforce Australia"

OUT = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_28Aug2026.pptx"


def bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, colour, line=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.shadow.inherit = False
    # Strip the theme style reference so no default shadow or outline colour
    # from the theme leaks in. Brand colours only, per Rule 3.
    style_el = shp._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style_el is not None:
        shp._element.remove(style_el)
    shp.fill.solid()
    shp.fill.fore_color.rgb = colour
    if line:
        shp.line.color.rgb = line
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    return shp


def text(slide, s, left, top, width, height, size, colour,
         bold=False, italic=False, align=PP_ALIGN.LEFT, font=SANS,
         anchor=None, shrink=True, wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = colour
    return box


def multiline(slide, lines, left, top, width, height, font=SANS,
              default_size=12, default_colour=OFF_WHITE, align=PP_ALIGN.LEFT,
              space_after=4, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    first = True
    for line in lines:
        cfg = line if isinstance(line, dict) else {"text": line}
        cfg.setdefault("size", default_size)
        cfg.setdefault("colour", default_colour)
        cfg.setdefault("bold", False)
        cfg.setdefault("italic", False)
        cfg.setdefault("font", font)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = cfg["text"]
        r.font.name = cfg["font"]
        r.font.size = Pt(cfg["size"])
        r.font.bold = cfg["bold"]
        r.font.italic = cfg["italic"]
        r.font.color.rgb = cfg["colour"]
    return box


def chrome(slide, eyebrow, header_right="PROFITPULSE"):
    """Fixed chrome per Section 6.0A: left stripe, header band, footer line."""
    bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
         12, OFF_WHITE, bold=True, font=SANS)
    text(slide, header_right, Inches(8.5), Inches(0.32), Inches(4.5), Inches(0.4),
         12, AMBER_B, bold=True, align=PP_ALIGN.RIGHT, font=SANS)
    # Footer line
    rect(slide, Inches(0.35), Inches(7.05), Inches(12.6), Pt(0.75), BLACK)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, "
                 "Profit-Pulse.com.au",
         Inches(0.35), Inches(7.12), Inches(9.5), Inches(0.3),
         8, BLACK, font=SANS)
    text(slide, DATE_STR, Inches(10.0), Inches(7.12), Inches(3.0), Inches(0.3),
         8, BLACK, align=PP_ALIGN.RIGHT, font=SANS)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    rect(slide, left, top, width, height, BLACK)
    rect(slide, left, top, width, Pt(4), TEAL)
    text(slide, number, left + Inches(0.1), top + Inches(0.14),
         width - Inches(0.2), Inches(0.5), 26, AMBER_B, bold=True, font=SERIF,
         wrap=False)
    multiline(slide, label_lines, left + Inches(0.12), top + Inches(0.68),
              width - Inches(0.24), Inches(0.55), font=SANS,
              default_size=10.5, default_colour=OFF_WHITE, space_after=0,
              line_spacing=1.0)
    text(slide, source, left + Inches(0.12), top + height - Inches(0.3),
         width - Inches(0.24), Inches(0.25), 7.5, OFF_WHITE,
         italic=True, font=SANS)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ============================================================================
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ============================================================================
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

text(s1, COMPANY, Inches(0.35), Inches(1.12), Inches(10.5), Inches(0.7),
     40, TEAL, bold=True, font=SERIF)
text(s1, "Property maintenance and trade services platform, Melbourne VIC",
     Inches(0.35), Inches(1.78), Inches(11.5), Inches(0.4),
     14, BLACK, font=SANS)

# Stat cards, six in a row
cards = [
    ("$12.8M", ["FY2025 revenue"], "SmartCompany Smart50 2025"),
    ("31%",    ["Revenue growth,", "FY24 to FY25"], "SmartCompany Smart50 2025"),
    ("19",     ["Team headcount,", "Burnley HQ"], "SmartCompany Smart50 2025"),
    ("#37",    ["National Smart50", "rank, 2025"], "SmartCompany Smart50 2025"),
    ("2014",   ["Founded,", "Melbourne"], "Company public records"),
    ("2024",   ["Telstra Best of", "Business, VIC winner"], "Telstra BOB Awards 2024"),
]
n = len(cards)
row_top = Inches(2.35)
row_h = Inches(1.6)
area_left_in = 0.35
area_right_in = 12.98
gap_in = 0.15
card_w_in = ((area_right_in - area_left_in) - gap_in * (n - 1)) / n
card_w = Inches(card_w_in)
x_in = area_left_in
for number, label_lines, source in cards:
    stat_card(s1, Inches(x_in), row_top, card_w, row_h, number, label_lines, source)
    x_in += card_w_in + gap_in

# Key Commercial Signals
sig_top = Inches(4.28)
text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), sig_top, Inches(6), Inches(0.3),
     11, TEAL, bold=True, font=SANS)

signals = [
    "Revenue rose from $9.7M to $12.8M, FY24 to FY25, per SmartCompany Smart50 citations.",
    "Agency client base doubled in 12 months, per 2024 Telstra Best of Business Awards VIC.",
    "Named Most Innovative Proptech at the 2024 Proptech Awards held in Sydney.",
    "Third straight Smart50 year, 2023 to 2025, rank improved from 46th to 37th nationally.",
    "Three active product lines in market: TradieConnect, RentSafe since 2021, and RentRepair.",
    "Single head office confirmed at Burnley VIC 3121, per the company's published contact page.",
]
sig_lines = [{"text": "•  " + s, "size": 12, "colour": BLACK} for s in signals]
multiline(s1, sig_lines, Inches(0.35), Inches(4.62), Inches(12.4), Inches(2.2),
          font=SANS, space_after=8, line_spacing=1.05)

# ============================================================================
# SLIDE 2: THE OPPORTUNITY, THREE COMMERCIAL OBSERVATIONS
# ============================================================================
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY")

text(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.1), Inches(12.4), Inches(0.4),
     15, BLACK, bold=True, font=SERIF)

col_top = Inches(1.7)
col_h = Inches(4.55)
col_gap = Inches(0.15)
col_w = Inches((12.98 - 0.35 - 0.15 * 2) / 3)

observations = [
    {
        "fill": TEAL, "text_colour": WHITE, "num_colour": BLACK,
        "index": "01",
        "header": "Three lines, one blended number",
        "body": ("TradieConnect, RentSafe since 2021, and RentRepair each serve a "
                 "different customer: manufacturers, property managers, and "
                 "landlords. Smart50 citations report only one blended revenue "
                 "and growth figure across all three. As volume compounds, the "
                 "open question is which line is carrying that growth and which "
                 "is diluting margin."),
    },
    {
        "fill": BLACK, "text_colour": OFF_WHITE, "num_colour": AMBER_B,
        "index": "02",
        "header": "Three straight years, one inflection point",
        "body": ("Taskforce Australia has placed in the Smart50 list three years "
                 "running, 2023 to 2025, climbing from rank 46 to rank 37 as "
                 "revenue rose from $9.7M to $12.8M. Telstra named this Outstanding "
                 "Growth in 2024. A three year compounding pattern like this is "
                 "usually the point where the plan that carried early growth needs "
                 "a deliberate reset for the next stage."),
    },
    {
        "fill": GOLD, "text_colour": BLACK, "num_colour": WHITE,
        "index": "03",
        "header": "Doubled clients, one dashboard",
        "body": ("The 2024 Telstra Best of Business citation records the agency "
                 "client base doubling within twelve months, the same year the "
                 "2024 Proptech Awards named the platform Most Innovative Proptech. "
                 "Client growth at that pace across a large tradesperson network "
                 "adds real operating complexity behind a clean headline number."),
    },
]

x = Inches(0.35)
for obs in observations:
    rect(s2, x, col_top, col_w, col_h, obs["fill"])
    text(s2, obs["index"], x + Inches(0.2), col_top + Inches(0.18),
         col_w - Inches(0.4), Inches(0.7), 34, obs["num_colour"], bold=True, font=SERIF)
    text(s2, obs["header"], x + Inches(0.2), col_top + Inches(0.95),
         col_w - Inches(0.4), Inches(0.75), 14, obs["text_colour"], bold=True, font=SERIF)
    multiline(s2, [{"text": obs["body"], "size": 11, "colour": obs["text_colour"]}],
              x + Inches(0.2), col_top + Inches(1.75), col_w - Inches(0.4), Inches(2.6),
              font=SANS, space_after=0, line_spacing=1.12)
    x = x + col_w + col_gap

text(s2, "These are observations offered in good faith. The company has built "
         "something impressive. The question is simply whether the financial "
         "architecture matches the ambition.",
     Inches(0.35), Inches(6.42), Inches(12.4), Inches(0.55),
     11, BLACK, italic=True, font=SANS)

# ============================================================================
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ============================================================================
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION")

left_x = Inches(0.35)
left_w = Inches(7.5)
right_x = Inches(8.15)
right_w = Inches(4.83)

text(s3, "Strategic Growth Diagnostic", left_x, Inches(1.15), left_w, Inches(0.55),
     22, TEAL, bold=True, font=SERIF)
text(s3, "$5,000 one off", left_x, Inches(1.68), left_w, Inches(0.4),
     16, BLACK, bold=True, font=SANS)
text(s3, "A six week engagement mapping revenue, capacity, and margin headroom "
         "across your active product lines, producing a twelve month growth plan "
         "with funding and capital allocation steps spelled out.",
     left_x, Inches(2.1), left_w, Inches(0.95), 12, BLACK, font=SANS)

# Step one block
rect(s3, left_x, Inches(3.15), left_w, Inches(1.55), OFF_WHITE,
     line=TEAL, line_w=Pt(1))
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.3),
     left_w - Inches(0.4), Inches(0.35), 13, TEAL, bold=True, font=SANS)
text(s3, "See the solutions matched to your size and industry.",
     left_x + Inches(0.2), Inches(3.68), left_w - Inches(0.4), Inches(0.35),
     11, BLACK, font=SANS)
text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.2), Inches(4.05),
     left_w - Inches(0.4), Inches(0.35), 13, TEAL, bold=True, font=SANS)
text(s3, "Purchase the suggested product now to get started",
     left_x + Inches(0.2), Inches(4.4), left_w - Inches(0.4), Inches(0.35),
     11, AMBER_D, bold=True, font=SANS)

text(s3, "Prefer a conversation first?", left_x, Inches(4.95), left_w, Inches(0.35),
     12, BLACK, bold=True, font=SANS)
text(s3, "Book a complimentary discovery call through the ProfitPulse booking page.",
     left_x, Inches(5.28), left_w, Inches(0.4), 11, BLACK, font=SANS)

rect(s3, left_x, Inches(5.85), left_w, Pt(0.75), BLACK)
text(s3, "Also considered: Product and Service Line Profitability, and "
         "KPI Dashboard Build and Run, to track all three lines in one view.",
     left_x, Inches(6.0), left_w, Inches(0.6), 10.5, BLACK,
     italic=True, font=SANS)

# Right column, credibility panel
rect(s3, right_x, Inches(1.15), right_w, Inches(5.55), BLACK)
text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.35),
     right_w - Inches(0.5), Inches(0.5), 19, AMBER_B, bold=True, font=SERIF)
text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.82),
     right_w - Inches(0.5), Inches(0.35), 12, OFF_WHITE, font=SANS)
rect(s3, right_x + Inches(0.25), Inches(2.3), right_w - Inches(0.5), Pt(1.25), TEAL)

cred = [
    "16 years of experience across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa, Mozambique",
    "Total GRBT project value in Queensland over AUD 10 billion",
]
multiline(s3, [{"text": c, "size": 12, "colour": OFF_WHITE} for c in cred],
          right_x + Inches(0.25), Inches(2.5), right_w - Inches(0.5), Inches(2.2),
          space_after=14, line_spacing=1.1)

rect(s3, right_x + Inches(0.25), Inches(4.75), right_w - Inches(0.5), Pt(1.25),
     TEAL)

contact = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
multiline(s3, [{"text": c, "size": 11.5, "colour": TEAL if i == 1 else OFF_WHITE}
               for i, c in enumerate(contact)],
          right_x + Inches(0.25), Inches(4.95), right_w - Inches(0.5), Inches(1.6),
          space_after=10)

text(s3, "Fractional CFO services: cashflow, investor ready documentation, "
         "and growth strategy for Australian SMEs.",
     right_x + Inches(0.25), Inches(6.15), right_w - Inches(0.5), Inches(0.5),
     9.5, OFF_WHITE, italic=True, font=SANS)

prs.save(OUT)
print("Deck saved:", OUT)
