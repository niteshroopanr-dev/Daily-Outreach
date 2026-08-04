"""
ProfitPulse Brief Builder
Target: Orbitkey | Date: 05 Aug 2026
Three slide prospect facing deck. House style per Section 6.0 and 6.0A. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STR = "05 Aug 2026"
COMPANY = "Orbitkey"
SLUG = "orbitkey"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


def bg(slide, colour):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = colour


def rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, x, y, w, h, size=12, bold=False, colour=BLACK,
         align=PP_ALIGN.LEFT, font="Calibri", anchor=None, italic=False,
         line_spacing=None, shrink=True):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if shrink:
        from pptx.enum.text import MSO_AUTO_SIZE
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


def chrome(slide, eyebrow, header_right=COMPANY):
    bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.3), Inches(0.38), Inches(7), Inches(0.35),
         size=12, bold=True, colour=OFF_WHITE, font="Calibri")
    text(slide, header_right, Inches(8.5), Inches(0.3), Inches(4.53), Inches(0.45),
         size=15, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT, font="Georgia")
    text(slide, f"Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.3), Inches(7.05), Inches(9.5), Inches(0.3), size=8,
         colour=RGBColor(0x66, 0x66, 0x66), font="Calibri")
    text(slide, DATE_STR, Inches(11.5), Inches(7.05), Inches(1.53), Inches(0.3),
         size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.RIGHT, font="Calibri")


# ============================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ============================================================
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

text(s1, "Orbitkey", Inches(0.3), Inches(1.15), Inches(9.5), Inches(0.7),
     size=40, bold=True, colour=BLACK, font="Georgia")
text(s1, "Design led everyday carry accessories brand, Richmond, Melbourne VIC",
     Inches(0.3), Inches(2.0), Inches(11.5), Inches(0.4), size=14,
     colour=RGBColor(0x33, 0x33, 0x33), font="Calibri")

# Stat cards, 5 across
labels = [
    ("$24.6M", "Annual revenue", "Smart50 2025 award"),
    ("44", "Team members", "Smart50 2025 award"),
    ("40%", "3 year revenue growth", "Export Awards 2025"),
    ("100+", "Countries exported to", "Export Awards 2025"),
    ("2013", "Founded", "SmartCompany history"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
top = Inches(2.4)
card_h = Inches(1.6)
x = Inches(0.3)
for num, lbl, src in labels:
    rect(s1, x, top, card_w, card_h, BLACK)
    rect(s1, x, top, card_w, Pt(4), TEAL)
    text(s1, num, x + Inches(0.15), top + Inches(0.18), card_w - Inches(0.3), Inches(0.5),
         size=28, bold=True, colour=AMBER_B, font="Georgia")
    text(s1, lbl, x + Inches(0.15), top + Inches(0.78), card_w - Inches(0.3), Inches(0.5),
         size=11, colour=WHITE, font="Calibri", line_spacing=1.0)
    text(s1, src, x + Inches(0.15), top + Inches(1.34), card_w - Inches(0.3), Inches(0.22),
         size=8, colour=RGBColor(0xAA, 0xAA, 0xAA), italic=True, font="Calibri")
    x += card_w + gap

# Key commercial signals
text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.3), Inches(4.3), Inches(6), Inches(0.3),
     size=12, bold=True, colour=TEAL, font="Calibri")

signals = [
    "Ranked 50th nationally, 2025 Smart50 Awards, fastest growing small businesses (SmartCompany).",
    "Winner, Creative Industries, Australian Export Awards 2025, Victorian state final (Export Awards).",
    "Ships from five third party logistics warehouses to over 100 export countries (Export Awards 2025).",
    "Raised over USD 3.25 million across six Kickstarter campaigns from 47,068 backers (SmartCompany).",
    "Registered as Orbitkey Pty Ltd, ABN 15 634 506 059, Richmond VIC (ASIC public register).",
]
y = Inches(4.68)
for sig in signals:
    rect(s1, Inches(0.3), y + Inches(0.08), Inches(0.08), Inches(0.08), AMBER_D)
    text(s1, sig, Inches(0.52), y, Inches(12.3), Inches(0.32), size=11,
         colour=RGBColor(0x22, 0x22, 0x22), font="Calibri")
    y += Inches(0.36)

# ============================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY")
text(s2, "Orbitkey. Three commercial observations from ProfitPulse.",
     Inches(0.3), Inches(1.15), Inches(11.5), Inches(0.4), size=16, bold=True,
     colour=BLACK, font="Georgia")

col_w = Inches(4.15)
col_gap = Inches(0.08)
col_top = Inches(1.75)
col_h = Inches(5.0)
cols = [
    (TEAL, WHITE, "01", "Serial crowdfunding signals a cash gap",
     "Orbitkey's own founders say Kickstarter helps fund new products and manage cash "
     "flow around production. That is a direct signal that cash tied up in inventory, "
     "debtors and supplier terms across five global warehouses deserves a closer look.\n\n"
     "Working Capital Unlock maps exactly this: cash trapped in stock, work in progress "
     "and terms, with a prioritised release plan. Comparable clients typically release "
     "8 to 15 percent of revenue in freed up cash."),
    (BLACK, OFF_WHITE, "02", "Three product lines, one margin picture",
     "Orbitkey sells across Everyday Carry, Travel and Bag Organisation, and Work "
     "Organisation, through direct to consumer, Kickstarter and wholesale to over a "
     "thousand European stores. Three channels and three product families rarely carry "
     "identical margin.\n\n"
     "A line by line profitability view would show which products and channels are "
     "genuinely funding the growth, and which are being carried by the others."),
    (GOLD, BLACK, "03", "Fast international growth needs a costed plan",
     "Revenue has grown 40 percent over three years and Orbitkey now exports to more "
     "than 100 countries from five logistics warehouses. That pace of expansion usually "
     "outruns the financial plan supporting it.\n\n"
     "A Strategic Growth Diagnostic would map revenue, capacity and margin headroom "
     "into a costed 12 month plan, so the next stage of growth is funded on purpose."),
]
x = Inches(0.3)
for fill, txtcol, idx, head, body in cols:
    rect(s2, x, col_top, col_w, col_h, fill)
    text(s2, idx, x + Inches(0.25), col_top + Inches(0.2), col_w - Inches(0.5), Inches(0.6),
         size=28, bold=True, colour=txtcol, font="Georgia")
    text(s2, head, x + Inches(0.25), col_top + Inches(0.85), col_w - Inches(0.5), Inches(0.75),
         size=14, bold=True, colour=txtcol, font="Calibri", line_spacing=1.05)
    text(s2, body, x + Inches(0.25), col_top + Inches(1.65), col_w - Inches(0.5), Inches(3.1),
         size=10.5, colour=txtcol, font="Calibri", line_spacing=1.1)
    x += col_w + col_gap

text(s2, "These are observations offered in good faith. Orbitkey has built something "
     "genuinely impressive. The question is simply whether the financial architecture "
     "matches the pace of the growth.",
     Inches(0.3), Inches(6.85), Inches(12.7), Inches(0.35), size=10, italic=True,
     colour=RGBColor(0x44, 0x44, 0x44), font="Calibri")

# ============================================================
# SLIDE 3: THE RECOMMENDATION
# ============================================================
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION")

left_w = Inches(7.6)
text(s3, "Working Capital Unlock", Inches(0.3), Inches(1.2), left_w, Inches(0.5),
     size=24, bold=True, colour=BLACK, font="Georgia")
text(s3, "$6,000 one off", Inches(0.3), Inches(1.75), left_w, Inches(0.35),
     size=15, bold=True, colour=AMBER_D, font="Calibri")
text(s3, "A four week project mapping cash trapped in debtors, inventory, work in "
     "progress, supplier terms and bank facilities across Orbitkey's global operation, "
     "with a prioritised action list to release cash back into the business.",
     Inches(0.3), Inches(2.2), left_w, Inches(0.95), size=11.5,
     colour=RGBColor(0x22, 0x22, 0x22), font="Calibri", line_spacing=1.15)

rect(s3, Inches(0.3), Inches(3.3), left_w, Inches(1.55), RGBColor(0xF2, 0xF2, 0xF0),
     line=AMBER_D)
text(s3, "Step one, answer a few quick questions", Inches(0.5), Inches(3.45), left_w - Inches(0.4),
     Inches(0.35), size=13, bold=True, colour=BLACK, font="Calibri")
text(s3, "See the solutions matched to your size and industry.",
     Inches(0.5), Inches(3.82), left_w - Inches(0.4), Inches(0.3), size=11,
     colour=RGBColor(0x33, 0x33, 0x33), font="Calibri")
qtag = text(s3, "profit-pulse.com.au/services/find-your-fit", Inches(0.5), Inches(4.2),
     left_w - Inches(0.4), Inches(0.35), size=12, bold=True, colour=TEAL, font="Calibri")
qtag.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://profit-pulse.com.au/services/find-your-fit/"

cta = rect(s3, Inches(0.3), Inches(5.05), left_w, Inches(0.55), TEAL)
ctatxt = text(s3, "Purchase the suggested product now to get started", Inches(0.3), Inches(5.2),
     left_w, Inches(0.3), size=12, bold=True, colour=WHITE, align=PP_ALIGN.CENTER, font="Calibri")
ctatxt.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z"

text(s3, "Prefer a conversation first? Book a complimentary discovery call.",
     Inches(0.3), Inches(5.85), left_w, Inches(0.3), size=10.5,
     colour=RGBColor(0x33, 0x33, 0x33), font="Calibri")
booktag = text(s3, "Book a complimentary discovery call", Inches(0.3), Inches(6.15), left_w, Inches(0.3),
     size=10.5, bold=True, colour=TEAL, font="Calibri")
booktag.text_frame.paragraphs[0].runs[0].hyperlink.address = (
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# Right panel, About ProfitPulse
rx = Inches(8.1)
rw = Inches(4.93)
rect(s3, rx, Inches(1.2), rw, Inches(5.55), BLACK)
text(s3, "Nitesh Roopa", rx + Inches(0.25), Inches(1.4), rw - Inches(0.5), Inches(0.45),
     size=17, bold=True, colour=AMBER_B, font="Georgia")
text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.25), Inches(1.85), rw - Inches(0.5),
     Inches(0.3), size=11.5, colour=OFF_WHITE, font="Calibri")

pts = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
y = Inches(2.35)
for p in pts:
    rect(s3, rx + Inches(0.25), y + Inches(0.07), Inches(0.07), Inches(0.07), TEAL)
    text(s3, p, rx + Inches(0.45), y, rw - Inches(0.7), Inches(0.4), size=10.5,
         colour=WHITE, font="Calibri", line_spacing=1.05)
    y += Inches(0.5)

rect(s3, rx + Inches(0.25), Inches(4.55), rw - Inches(0.5), Pt(1.5), TEAL)
contact = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
y = Inches(4.75)
for c in contact:
    text(s3, c, rx + Inches(0.25), y, rw - Inches(0.5), Inches(0.35), size=11,
         colour=OFF_WHITE, font="Calibri")
    y += Inches(0.4)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Orbitkey_05Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
