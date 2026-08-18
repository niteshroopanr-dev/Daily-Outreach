"""
ProfitPulse Brief Builder
Target: Sniip | Date: 19 Aug 2026
Three-slide prospect-facing deck. House style v3.3. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours, official logo palette only
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Georgia"
SANS  = "Arial"

W = Inches(13.333)
H = Inches(7.5)

DATE_STR = "19 Aug 2026"
COMPANY  = "Sniip"


def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, colour, line_colour=None, line_w=None):
    shp = slide.shapes.add_shape(1, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, left, top, width, height, size=12, bold=False, italic=False,
         colour=BLACK, align=PP_ALIGN.LEFT, font=SANS, anchor=None, wrap=True,
         line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
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


def multiline(slide, lines, left, top, width, height, font=SANS, align=PP_ALIGN.LEFT,
              space_after=4, wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for ln in lines:
        cfg = {"text": ln, "size": 11, "colour": BLACK, "bold": False, "italic": False} \
            if isinstance(ln, str) else {
                "text": ln.get("text", ""), "size": ln.get("size", 11),
                "colour": ln.get("colour", BLACK), "bold": ln.get("bold", False),
                "italic": ln.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = cfg["text"]
        r.font.name = font
        r.font.size = Pt(cfg["size"])
        r.font.bold = cfg["bold"]
        r.font.italic = cfg["italic"]
        r.font.color.rgb = cfg["colour"]
    return box


def chrome(slide, eyebrow, header_right=None):
    """Fixed chrome: left accent stripe, black header band, footer line."""
    set_bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
         size=12, bold=True, colour=OFF_WHITE, font=SANS, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, header_right or "PROFITPULSE", Inches(9.0), Inches(0.32), Inches(4.0), Inches(0.4),
         size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT, font=SANS,
         anchor=MSO_ANCHOR.MIDDLE)
    # footer
    rect(slide, Inches(0.35), Inches(7.03), Inches(12.6), Pt(0.75), TEAL)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, "
                "Profit-Pulse.com.au",
         Inches(0.35), Inches(7.1), Inches(9.5), Inches(0.3),
         size=8, colour=BLACK, font=SANS)
    text(slide, DATE_STR, Inches(10.5), Inches(7.1), Inches(2.45), Inches(0.3),
         size=8, colour=BLACK, align=PP_ALIGN.RIGHT, font=SANS)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    rect(slide, left, top, width, height, BLACK)
    rect(slide, left, top, width, Pt(4), TEAL)
    text(slide, number, left + Inches(0.15), top + Inches(0.12), width - Inches(0.3), Inches(0.42),
         size=27, bold=True, colour=AMBER_B, font=SERIF)
    multiline(slide, [{"text": l, "size": 10.5, "colour": OFF_WHITE} for l in label_lines],
              left + Inches(0.15), top + Inches(0.8), width - Inches(0.3), Inches(0.5),
              font=SANS, space_after=0)
    text(slide, source, left + Inches(0.15), top + height - Inches(0.26), width - Inches(0.3),
         Inches(0.22), size=7, colour=OFF_WHITE, italic=True, font=SANS)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COMMERCIAL INTELLIGENCE BRIEF
# ═══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

text(s1, "Sniip", Inches(0.35), Inches(1.15), Inches(8), Inches(0.85),
     size=42, bold=True, colour=BLACK, font=SERIF)
text(s1, "Consumer and business bill payment platform, Brisbane, Queensland",
     Inches(0.35), Inches(2.05), Inches(11), Inches(0.3),
     size=13, colour=BLACK, font=SANS)

cards = [
    ("2014", ["Founded in", "Brisbane"], "Sniip company site"),
    ("200K+", ["Business and personal", "customers"], "Sniip company site"),
    ("$500M+", ["Processed in", "bill payments"], "Sniip company site"),
    ("2025", ["AFR Fast 100", "debut year"], "AFR Fast 100 2025 list"),
    ("5", ["Major payment partners", "added"], "Company & partner news"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
x = Inches(0.35)
y = Inches(2.4)
card_h = Inches(1.6)
for num, lbl, src in cards:
    stat_card(s1, x, y, card_w, card_h, num, lbl, src)
    x += card_w + gap

text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.35), Inches(6), Inches(0.3),
     size=12, bold=True, colour=TEAL, font=SANS)

signals = [
    "Debuted on the AFR Fast 100 2025, confirming FY25 revenue above five million dollars.",
    "Won Best Innovation in Payments at the Fintech Australia Finnies Awards, June 2025.",
    "Named a finalist again in the same category at the 2026 Finnies Awards.",
    "Added BPAY, American Express, Qantas Frequent Flyer, Virgin and WEX Motorpass as partners.",
    "LinkedIn lists a team of 11 to 50 people running the platform from Brisbane.",
    "Founded in 2014 by Damien Vasta, who remains Founder and Chief Executive Officer.",
]
multiline(s1, [{"text": "•  " + s, "size": 11.5, "colour": BLACK} for s in signals],
          Inches(0.35), Inches(4.72), Inches(12.6), Inches(2.1), font=SANS, space_after=6)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2 — THE OPPORTUNITY
# ═══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY", "SNIIP")

text(s2, "Three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.15), Inches(12.5), Inches(0.4),
     size=15, bold=True, colour=BLACK, font=SERIF)

col_w = Inches(4.0)
col_h = Inches(4.55)
col_top = Inches(1.75)
col_gap = Inches(0.13)
col_x = [Inches(0.35), Inches(0.35) + col_w + col_gap, Inches(0.35) + 2 * (col_w + col_gap)]
col_fill = [TEAL, BLACK, GOLD]
col_text = [WHITE, WHITE, BLACK]
col_accent = [BLACK, AMBER_B, BLACK]

observations = [
    ("01", "Five partners, one finance team",
     "In roughly a year Sniip added BPAY, American Express, Qantas Frequent Flyer, "
     "Virgin Australia Business Flyer and WEX Motorpass. Each partner carries its own "
     "settlement cycle and margin profile. Layering five economic models onto one "
     "finance function in twelve months tests reporting fast."),
    ("02", "A Fast 100 debut raises the bar",
     "Sniip's AFR Fast 100 2025 debut confirms revenue growth strong enough to clear "
     "the five million dollar threshold on a three year view. Lists like this draw "
     "investor and partner attention. Converting that attention well means board "
     "grade numbers are ready on request, not assembled after the fact."),
    ("03", "A rewards liability needs a steady hand",
     "Every bill paid can earn points across Qantas, Virgin and Amex programs. That "
     "is a growing liability sitting behind a fast growing transaction base. A "
     "monthly Fractional CFO Partnership keeps cash position, partner economics and "
     "the management pack as sharp as the product roadmap."),
]

for i, (num, head, para) in enumerate(observations):
    cx = col_x[i]
    rect(s2, cx, col_top, col_w, col_h, col_fill[i])
    text(s2, num, cx + Inches(0.25), col_top + Inches(0.2), col_w - Inches(0.5), Inches(0.65),
         size=30, bold=True, colour=col_accent[i], font=SERIF)
    text(s2, head, cx + Inches(0.25), col_top + Inches(0.95), col_w - Inches(0.5), Inches(0.75),
         size=14.5, bold=True, colour=col_text[i], font=SANS, wrap=True)
    text(s2, para, cx + Inches(0.25), col_top + Inches(1.75), col_w - Inches(0.5), Inches(2.6),
         size=10.5, colour=col_text[i], font=SANS, wrap=True, line_spacing=1.12)

text(s2, "These observations are offered in good faith. Sniip has built something "
         "genuinely impressive. The question is simply whether the financial "
         "architecture keeps pace with the partnership roadmap.",
     Inches(0.35), Inches(6.42), Inches(12.6), Inches(0.55),
     size=10.5, italic=True, colour=BLACK, font=SANS, wrap=True)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3 — THE RECOMMENDATION AND HOW TO START
# ═══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION", "SNIIP")

# Left column
lx = Inches(0.35)
lw = Inches(7.3)

text(s3, "Fractional CFO Partnership", lx, Inches(1.2), lw, Inches(0.55),
     size=24, bold=True, colour=BLACK, font=SERIF)
text(s3, "$4,950 per month", lx, Inches(1.78), lw, Inches(0.4),
     size=16, bold=True, colour=TEAL, font=SANS)
text(s3, "ProfitPulse verified price. The questionnaire confirms the exact fit for Sniip.",
     lx, Inches(2.16), lw, Inches(0.3), size=9.5, italic=True,
     colour=BLACK, font=SANS)

text(s3, "A senior financial partner at the table each month: management pack, "
         "quarterly board grade review, and support as new partnerships add complexity.",
     lx, Inches(2.5), lw, Inches(0.6), size=11.5, colour=BLACK,
     font=SANS, wrap=True)

# Step one block
rect(s3, lx, Inches(3.25), lw, Inches(1.35), OFF_WHITE, line_colour=TEAL, line_w=Pt(1))
text(s3, "Step one, answer a few quick questions", lx + Inches(0.2), Inches(3.4), lw - Inches(0.4),
     Inches(0.35), size=12.5, bold=True, colour=BLACK, font=SANS)
text(s3, "See the solutions matched to your size and industry.", lx + Inches(0.2), Inches(3.78),
     lw - Inches(0.4), Inches(0.3), size=10.5, colour=BLACK, font=SANS)
text(s3, "profit-pulse.com.au/services/find-your-fit", lx + Inches(0.2), Inches(4.1),
     lw - Inches(0.4), Inches(0.35), size=12, bold=True, colour=TEAL, font=SANS)

# Direct CTA
rect(s3, lx, Inches(4.85), lw, Inches(0.62), TEAL)
text(s3, "Purchase the suggested product now to get started", lx, Inches(4.85), lw, Inches(0.62),
     size=13, bold=True, colour=WHITE, align=PP_ALIGN.CENTER, font=SANS, anchor=MSO_ANCHOR.MIDDLE)

text(s3, "Prefer a conversation first?", lx, Inches(5.65), lw, Inches(0.3),
     size=10.5, colour=BLACK, font=SANS)
text(s3, "Book a complimentary discovery call", lx, Inches(5.95), lw, Inches(0.35),
     size=12, bold=True, colour=AMBER_D, font=SANS)

# Right column: credibility panel
rx = Inches(7.95)
rw = Inches(5.03)
rect(s3, rx, Inches(1.2), rw, Inches(5.1), BLACK)
rect(s3, rx, Inches(1.2), rw, Pt(4), TEAL)
text(s3, "Nitesh Roopa", rx + Inches(0.25), Inches(1.42), rw - Inches(0.5), Inches(0.4),
     size=17, bold=True, colour=AMBER_B, font=SERIF)
text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.25), Inches(1.82), rw - Inches(0.5),
     Inches(0.3), size=11, colour=OFF_WHITE, font=SANS)

cred = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "AUD 10 billion GRBT project value in Queensland",
]
multiline(s3, [{"text": "•  " + c, "size": 10.5, "colour": OFF_WHITE} for c in cred],
          rx + Inches(0.25), Inches(2.3), rw - Inches(0.5), Inches(1.5), font=SANS, space_after=7)

rect(s3, rx + Inches(0.25), Inches(3.95), rw - Inches(0.5), Pt(1), TEAL)

contact = [
    {"text": "Profit-Pulse.com.au", "size": 11, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 11, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 11, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 9.5, "colour": OFF_WHITE},
]
multiline(s3, contact, rx + Inches(0.25), Inches(4.15), rw - Inches(0.5), Inches(1.6),
          font=SANS, space_after=6)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Sniip_19Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
