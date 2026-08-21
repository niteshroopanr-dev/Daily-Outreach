"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 22 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
House style per Section 6 / 6.0A: white body, black header band, amber left stripe.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours, exactly seven
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STAMP = "22 Aug 2026"
COMPANY = "Taskforce Australia"

HEAD = "Georgia"
BODY = "Calibri"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


def bg(slide, colour):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = colour


def _strip_style_ref(shp):
    """Remove the theme p:style element (lnRef/fillRef/effectRef/fontRef) so no
    theme shadow or colour bleeds through, and force an explicit empty effectLst."""
    el = shp._element
    style = el.find("{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style is not None:
        el.remove(style)


def rect(slide, x, y, w, h, fill_colour, line_colour=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    _strip_style_ref(shp)
    return shp


def text(slide, s, x, y, w, h, size=12, bold=False, italic=False, colour=BLACK,
         align=PP_ALIGN.LEFT, font=BODY, anchor=None, line_spacing=None, shrink=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor:
        tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = s.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = colour
    return box


def multiline(slide, runs, x, y, w, h, align=PP_ALIGN.LEFT, line_spacing=None, space_after=None):
    """runs: list of dicts text/size/colour/bold/font/italic, one per paragraph."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, cfg in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        if space_after is not None:
            p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = cfg["text"]
        r.font.name = cfg.get("font", BODY)
        r.font.size = Pt(cfg.get("size", 12))
        r.font.bold = cfg.get("bold", False)
        r.font.italic = cfg.get("italic", False)
        r.font.color.rgb = cfg.get("colour", BLACK)
    return box


def chrome(slide, eyebrow, footer_on_white=True):
    """Left amber stripe, black header band, footer line. Shared across slides."""
    bg(slide, WHITE)
    rect(slide, 0, 0, 0.1, 7.5, AMBER_B)
    rect(slide, 0.1, 0, 13.233, 1.0, BLACK)
    text(slide, eyebrow, 0.35, 0.34, 8.5, 0.4, size=12, bold=True, colour=OFFWHITE, font=BODY)
    text(slide, "PROFITPULSE", 9.5, 0.34, 3.4, 0.4, size=12, bold=True, colour=TEAL,
         align=PP_ALIGN.RIGHT, font=BODY)
    if footer_on_white:
        text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
             0.35, 7.08, 8.5, 0.3, size=8, colour=BLACK, font=BODY)
        text(slide, DATE_STAMP, 10.5, 7.08, 2.5, 0.3, size=8, colour=BLACK,
             align=PP_ALIGN.RIGHT, font=BODY)


def footer_dark_strip(slide):
    """Footer strip for slide two, where column fills vary."""
    rect(slide, 0.1, 7.05, 13.233, 0.35, BLACK)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
         0.35, 7.12, 8.5, 0.25, size=8, colour=OFFWHITE, font=BODY)
    text(slide, DATE_STAMP, 10.5, 7.12, 2.5, 0.25, size=8, colour=OFFWHITE,
         align=PP_ALIGN.RIGHT, font=BODY)


# ════════════════════════════════════════════════════════════════════════
# SLIDE ONE, THE COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(BLANK)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

text(s1, COMPANY, 0.35, 1.12, 9.0, 0.75, size=42, bold=True, colour=BLACK, font=HEAD)
text(s1, "National property compliance and maintenance platform, Burnley, Melbourne VIC",
     0.35, 1.85, 10.5, 0.35, size=12, colour=BLACK, font=BODY)

# Stat cards, six across, single row, shared top edge and height
cards = [
    ("$12.8M", "Revenue, FY2025", "Smart50 2025, rank 37"),
    ("31%",    "Revenue growth\nin FY2025", "Smart50 2025 citation"),
    ("19",     "Employees\nnationally", "Smart50 2025 citation"),
    ("2014",   "Founded, Melbourne", "Smart50 award profiles"),
    ("140K+",  "Compliance jobs\ndelivered since 2021", "Taskforce RentSafe site"),
    ("5,500+", "Trades in the\nnational network", "Taskforce Australia site"),
]
n = len(cards)
card_top = 2.35
card_h = 1.55
gap = 0.14
left_margin = 0.35
right_margin = 0.35
usable = 13.333 - left_margin - right_margin
card_w = (usable - gap * (n - 1)) / n

for i, (num, label, source) in enumerate(cards):
    cx = left_margin + i * (card_w + gap)
    rect(s1, cx, card_top, card_w, card_h, BLACK)
    rect(s1, cx, card_top, card_w, 0.05, TEAL)
    text(s1, num, cx + 0.12, card_top + 0.16, card_w - 0.24, 0.5,
         size=27, bold=True, colour=AMBER_B, font=HEAD)
    text(s1, label, cx + 0.12, card_top + 0.68, card_w - 0.24, 0.55,
         size=10, colour=OFFWHITE, font=BODY, line_spacing=1.0)
    text(s1, source, cx + 0.12, card_top + card_h - 0.28, card_w - 0.24, 0.24,
         size=7, italic=True, colour=WHITE, font=BODY)

# Key Commercial Signals
sig_top = card_top + card_h + 0.28
text(s1, "KEY COMMERCIAL SIGNALS", left_margin, sig_top, 6.0, 0.3,
     size=11, bold=True, colour=TEAL, font=BODY)

signals = [
    "Revenue moved from $7.43M to $9.7M to $12.8M across three straight Smart50 citations, 2023 to 2025.",
    "Target of 50 to 60% compound annual growth over the next three years, self funded organic growth (Smart50 2024 profile).",
    "New preferred supplier agreement with BigginScott Group real estate group, effective 1 July 2025 (Real+ partner listing).",
    "RentSafe platform has delivered 140,000+ jobs across 300 real estate offices and 180 major brands (Taskforce RentSafe site).",
    "Won the Proptech Association Award in 2024 and 2025; founder Jason Bright a 2026 Leader of the Year finalist (Proptech Australia).",
]
sig_runs = [{"text": "•  " + s, "size": 11, "colour": BLACK, "font": BODY} for s in signals]
multiline(s1, sig_runs, left_margin, sig_top + 0.32, usable, 2.15, line_spacing=1.12, space_after=6)

# ════════════════════════════════════════════════════════════════════════
# SLIDE TWO, THE OPPORTUNITY
# ════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(BLANK)
bg(s2, WHITE)
rect(s2, 0, 0, 0.1, 7.5, AMBER_B)
rect(s2, 0.1, 0, 13.233, 1.0, BLACK)
text(s2, "THE OPPORTUNITY", 0.35, 0.28, 7.0, 0.35, size=12, bold=True, colour=OFFWHITE, font=BODY)
text(s2, f"{COMPANY}. Three commercial observations from ProfitPulse.", 0.35, 0.6, 9.0, 0.35,
     size=10, colour=TEAL, font=BODY)
text(s2, "PROFITPULSE", 9.5, 0.34, 3.4, 0.4, size=12, bold=True, colour=TEAL,
     align=PP_ALIGN.RIGHT, font=BODY)

col_top = 1.0
col_h = 5.35
col_w = (13.233 - 0.1) / 3  # excludes stripe already accounted by starting at 0.1
col_x0 = 0.1

cols = [
    {
        "fill": TEAL, "text_colour": BLACK, "idx_colour": WHITE,
        "header": "Revenue is compounding faster than the team",
        "body": ("Revenue moved from $7.43 million to $9.7 million to $12.8 million across "
                 "three consecutive Smart50 citations, a 31 percent lift in the most recent "
                 "year, while headcount held at 19 across the same period. Revenue compounding "
                 "well ahead of headcount is a strong platform signal, but it leaves less "
                 "internal capacity to build the cash forecasting discipline the next phase needs."),
    },
    {
        "fill": BLACK, "text_colour": OFFWHITE, "idx_colour": AMBER_B,
        "header": "A three year target raises the cash stakes",
        "body": ("Taskforce has set a target of 50 to 60 percent compound annual growth for the "
                 "next three years, funded through self funded organic growth rather than a "
                 "further capital raise. Funding that pace from operating cash, on top of the "
                 "new BigginScott Group volume from 1 July 2025, narrows the margin for error on "
                 "receivables timing and contractor payment runs as the base scales."),
    },
    {
        "fill": GOLD, "text_colour": BLACK, "idx_colour": WHITE,
        "header": "Scale across 300 offices concentrates risk in the middle",
        "body": ("RentSafe has delivered more than 140,000 jobs across 300 real estate offices "
                 "and 180 major brands since 2021, and the new BigginScott Group relationship "
                 "adds to that book from July 2025. Growth this diversified is a genuine "
                 "strength, but running it well depends on knowing, week to week, which "
                 "invoicing cycles and contractor payment runs are keeping pace with volume."),
    },
]

for i, c in enumerate(cols):
    cx = col_x0 + i * col_w
    rect(s2, cx, col_top, col_w, col_h, c["fill"])
    pad = 0.28
    text(s2, f"0{i+1}", cx + pad, col_top + 0.22, col_w - 2 * pad, 0.8,
         size=40, bold=True, colour=c["idx_colour"], font=HEAD)
    text(s2, c["header"], cx + pad, col_top + 1.05, col_w - 2 * pad, 0.85,
         size=15, bold=True, colour=c["text_colour"], font=HEAD, line_spacing=1.05)
    text(s2, c["body"], cx + pad, col_top + 1.95, col_w - 2 * pad, 3.15,
         size=10.5, colour=c["text_colour"], font=BODY, line_spacing=1.18)

text(s2,
     "These observations are offered in good faith. Taskforce has built something genuinely "
     "well regarded in its category. The question is simply whether the weekly cash rhythm "
     "keeps pace with the growth target the business has set for itself.",
     0.35, col_top + col_h + 0.1, 12.6, 0.55, size=10.5, italic=True, colour=BLACK, font=BODY)

footer_dark_strip(s2)


# ════════════════════════════════════════════════════════════════════════
# SLIDE THREE, THE RECOMMENDATION AND HOW TO START
# ════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(BLANK)
chrome(s3, "THE RECOMMENDATION")

left_x = 0.35
left_w = 7.35

text(s3, "13 Week Cash Flow Build", left_x, 1.15, left_w, 0.55,
     size=28, bold=True, colour=BLACK, font=HEAD)
text(s3, "$2,650 one off", left_x, 1.82, left_w, 0.4,
     size=16, bold=True, colour=TEAL, font=BODY)
text(s3,
     "A rolling 13 week cash flow forecast built from your accounting data, with three "
     "scenarios and a weekly cadence playbook your team can run after handover.",
     left_x, 2.35, left_w, 0.65, size=11.5, colour=BLACK, font=BODY, line_spacing=1.2)

# Step one block
rect(s3, left_x, 3.25, left_w, 1.55, TEAL, line_colour=TEAL)
text(s3, "Step one, answer a few quick questions", left_x + 0.24, 3.46, left_w - 0.48, 0.35,
     size=13.5, bold=True, colour=BLACK, font=BODY)
text(s3, "See the solutions matched to your size and industry.",
     left_x + 0.24, 3.84, left_w - 0.48, 0.35, size=11.5, colour=BLACK, font=BODY)
link_box = text(s3, "profit-pulse.com.au/services/find-your-fit",
     left_x + 0.24, 4.24, left_w - 0.48, 0.35, size=12.5, bold=True, colour=BLACK, font=BODY)
run = link_box.text_frame.paragraphs[0].runs[0]
run.hyperlink.address = "https://profit-pulse.com.au/services/find-your-fit/"
run.font.underline = True

# Direct CTA
cta = rect(s3, left_x, 5.1, left_w, 0.62, WHITE, line_colour=AMBER_D, line_w=Pt(1.5))
cta_box = text(s3, "Purchase the suggested product now to get started",
     left_x + 0.2, 5.27, left_w - 0.4, 0.35, size=12.5, bold=True, colour=BLACK, font=BODY,
     align=PP_ALIGN.LEFT)
cta_run = cta_box.text_frame.paragraphs[0].runs[0]
cta_run.hyperlink.address = "https://buy.stripe.com/3cI4gA1qw5280yQgYT3ks0w"
cta_run.font.underline = True

text(s3, "Prefer a conversation first?", left_x, 5.98, left_w, 0.32,
     size=11.5, colour=BLACK, font=BODY)
book_box = text(s3, "Book a complimentary discovery call",
     left_x, 6.32, left_w, 0.32, size=12, bold=True, colour=BLACK, font=BODY)
book_run = book_box.text_frame.paragraphs[0].runs[0]
book_run.hyperlink.address = (
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/"
    "?ismsaljsauthenabled=true"
)
book_run.font.underline = True

# Right column, credibility panel
right_x = 8.05
right_w = 4.93
rect(s3, right_x, 1.15, right_w, 5.55, BLACK)
pad = 0.28
text(s3, "NITESH ROOPA", right_x + pad, 1.35, right_w - 2 * pad, 0.4,
     size=17, bold=True, colour=AMBER_B, font=HEAD)
text(s3, "CA, Managing Partner, ProfitPulse", right_x + pad, 1.78, right_w - 2 * pad, 0.32,
     size=11.5, colour=WHITE, font=BODY)
rect(s3, right_x + pad, 2.18, right_w - 2 * pad, 0.03, TEAL)

cred_runs = [
    {"text": "16 years of experience across 4 countries", "size": 11.5, "colour": OFFWHITE},
    {"text": "52 deals executed and managed", "size": 11.5, "colour": OFFWHITE},
    {"text": "Largest single deal, USD 1.3 billion, Cahora Bassa, Mozambique Government, Hydro",
     "size": 11.5, "colour": OFFWHITE},
    {"text": "Total GRBT project value over AUD 10 billion", "size": 11.5, "colour": OFFWHITE},
]
multiline(s3, cred_runs, right_x + pad, 2.45, right_w - 2 * pad, 2.2,
          line_spacing=1.25, space_after=14)

rect(s3, right_x + pad, 4.95, right_w - 2 * pad, 0.03, TEAL)
contact_runs = [
    {"text": "Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 12, "colour": WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 11, "colour": WHITE},
]
multiline(s3, contact_runs, right_x + pad, 5.2, right_w - 2 * pad, 1.6,
          line_spacing=1.3, space_after=10)

# Force the theme hyperlink colour to brand black, so renderers cannot inject
# the default PowerPoint hyperlink blue on any linked run.
from lxml import etree
NSMAP_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
theme_part = prs.slide_masters[0].part.part_related_by(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)
theme_el = etree.fromstring(theme_part.blob)
for tag in ("hlink", "folHlink"):
    node = theme_el.find(f".//{{{NSMAP_A}}}{tag}")
    if node is not None:
        srgb = node.find(f"{{{NSMAP_A}}}srgbClr")
        if srgb is not None:
            srgb.set("val", "000000")
        else:
            sys_clr = node.find(f"{{{NSMAP_A}}}sysClr")
            if sys_clr is not None:
                sys_clr.set("lastClr", "000000")
theme_part._blob = etree.tostring(theme_el, xml_declaration=True, encoding="UTF-8", standalone=True)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_22Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
