"""
ProfitPulse Brief PPTX Builder, house style v3.3.
Target: Equality Media + Marketing | Date for: 17 Jul 2026
Mirrors build_pdf.py's geometry (same content, same layout, expressed in EMU/inches).
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import brief_content as bc

BLACK = RGBColor(0x00, 0x00, 0x00)
TEAL = RGBColor(0x01, 0xA2, 0x96)
AMBER_B = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D = RGBColor(0xF6, 0xA1, 0x02)
GOLD = RGBColor(0xE3, 0xA7, 0x12)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)
# Brand colours only (Rule 3): all "muted"/"divider"/"tile" needs below map onto
# the same seven approved colours, varied by size/weight rather than off-palette tints.
DARK_TILE = BLACK
MUTED = BLACK
MUTED2 = OFF_WHITE
DARK_TXT = BLACK
GREY333 = BLACK
LIGHT_TEAL_BG = OFF_WHITE
DIVIDER = TEAL

W = Inches(13.333)
H = Inches(7.5)
SERIF = "Cambria"
SANS = "Calibri"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


def set_bg(slide, colour):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = colour


def rect(slide, x, y, w, h, fill, line=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def txt(slide, s, x, y, w, h, size, colour, font=SANS, bold=False, italic=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True, hyperlink=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = s
    run.font.size = Pt(size)
    run.font.name = font
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    if hyperlink:
        run.hyperlink.address = hyperlink
    return box


def bullets(slide, items, x, y, w, size, colour, font=SANS, leading=15, bullet_gap=0, height=Inches(1.5)):
    box = slide.shapes.add_textbox(x, y, w, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"•  {it}"
        run.font.size = Pt(size)
        run.font.name = font
        run.font.color.rgb = colour
        p.space_after = Pt(bullet_gap)
    return box


def chrome(slide, eyebrow, right_label="PROFITPULSE"):
    set_bg(slide, WHITE)
    rect(slide, 0, 0, W, Inches(1.0), BLACK)
    txt(slide, eyebrow, Inches(0.35), Inches(0.28), Inches(7), Inches(0.4), 12, OFF_WHITE, bold=True)
    txt(slide, right_label, Inches(9.833), Inches(0.28), Inches(3.5), Inches(0.4), 12, OFF_WHITE,
        bold=True, align=PP_ALIGN.RIGHT)
    rect(slide, 0, 0, Inches(0.1), H, AMBER_D)


def footer(slide):
    txt(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
        Inches(0.35), Inches(7.05), Inches(7), Inches(0.3), 8, MUTED)
    txt(slide, bc.DATE_FOR, Inches(10.333), Inches(7.05), Inches(2.5), Inches(0.3), 8, MUTED,
        align=PP_ALIGN.RIGHT)


CONTENT_X0 = Inches(0.35)
CONTENT_W = Inches(13.0) - Inches(0.35)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(BLANK)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

txt(s1, bc.COMPANY, CONTENT_X0, Inches(1.15), Inches(11), Inches(0.6), 34, BLACK, font=SERIF, bold=True)
txt(s1, bc.LOCATION_LINE, CONTENT_X0, Inches(1.68), Inches(11), Inches(0.3), 12.5, GREY333)

n_cards = len(bc.STAT_CARDS)
gap = Inches(0.15)
card_w = (CONTENT_W - gap * (n_cards - 1)) / n_cards
card_top = Inches(2.12)
card_h = Inches(1.35)
for i, card in enumerate(bc.STAT_CARDS):
    cx = CONTENT_X0 + i * (card_w + gap)
    rect(s1, cx, card_top, card_w, card_h, DARK_TILE)
    rect(s1, cx, card_top, card_w, Pt(4), TEAL)
    txt(s1, card["number"], cx + Inches(0.14), card_top + Inches(0.18), card_w - Inches(0.28),
        Inches(0.42), 27, AMBER_B, bold=True)
    txt(s1, card["label"], cx + Inches(0.14), card_top + Inches(0.72), card_w - Inches(0.28),
        Inches(0.42), 10, OFF_WHITE)
    txt(s1, card["source"], cx + Inches(0.14), card_top + card_h - Inches(0.26), card_w - Inches(0.28),
        Inches(0.24), 7.5, MUTED2, italic=True)

chart_title_y = Inches(3.75)
txt(s1, "REVENUE TRAJECTORY, VERIFIED ($ MILLION)", CONTENT_X0, chart_title_y, Inches(8), Inches(0.25),
    11, TEAL, bold=True)
bar_area_top = chart_title_y + Inches(0.47)
bar_area_h = Inches(0.61)
max_val = max(v for _, v in bc.CHART_DATA)
bar_gap = Inches(0.5)
bar_w = Inches(1.1)
for i, (label, val) in enumerate(bc.CHART_DATA):
    bx = CONTENT_X0 + i * (bar_w + bar_gap)
    bh = Inches(0.61 * (val / max_val))
    rect(s1, bx, bar_area_top + (bar_area_h - bh), bar_w, bh, TEAL)
    txt(s1, f"${val:g}M", bx, bar_area_top + (bar_area_h - bh) - Inches(0.2), Inches(1.3), Inches(0.22),
        11, BLACK, bold=True)
category_y = bar_area_top + bar_area_h + Inches(0.19)
for i, (label, val) in enumerate(bc.CHART_DATA):
    bx = CONTENT_X0 + i * (bar_w + bar_gap)
    txt(s1, label, bx, category_y, Inches(1.4), Inches(0.2), 10, GREY333)
source_y = category_y + Inches(0.21)
txt(s1, bc.CHART_SOURCE, CONTENT_X0, source_y, Inches(11), Inches(0.2), 7.5, BLACK, italic=True)

sig_top = source_y + Inches(0.28)
txt(s1, "KEY COMMERCIAL SIGNALS", CONTENT_X0, sig_top, Inches(6), Inches(0.25), 11, TEAL, bold=True)
bullets(s1, bc.KEY_SIGNALS, CONTENT_X0, sig_top + Inches(0.28), CONTENT_W, 10.5, DARK_TXT,
        leading=10.5, bullet_gap=4, height=Inches(1.3))

footer(s1)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(BLANK)
chrome(s2, "THE OPPORTUNITY")
txt(s2, f"{bc.COMPANY}: Three commercial observations from ProfitPulse",
    CONTENT_X0, Inches(1.18), Inches(12), Inches(0.35), 14, BLACK, font=SERIF, bold=True)

col_gap = Inches(0.15)
col_w = (CONTENT_W - col_gap * 2) / 3
col_top = Inches(1.55)
col_h = Inches(4.75)
FILLS = {"TEAL": TEAL, "BLACK": BLACK, "GOLD": GOLD}
for i, obs in enumerate(bc.OBSERVATIONS):
    cx = CONTENT_X0 + i * (col_w + col_gap)
    fill = FILLS[obs["fill"]]
    rect(s2, cx, col_top, col_w, col_h, fill)
    text_colour = WHITE if obs["fill"] == "BLACK" else BLACK
    pad = Inches(0.22)
    txt(s2, obs["index"], cx + pad, col_top + Inches(0.25), col_w - pad * 2, Inches(0.55), 34,
        text_colour, font=SERIF, bold=True)
    txt(s2, obs["header"], cx + pad, col_top + Inches(0.95), col_w - pad * 2, Inches(0.6), 13.5,
        text_colour, bold=True)
    txt(s2, obs["body"], cx + pad, col_top + Inches(1.65), col_w - pad * 2, Inches(2.9), 10.5,
        text_colour)

txt(s2, bc.CLOSING_LINE, CONTENT_X0, col_top + col_h + Inches(0.2), CONTENT_W, Inches(0.5), 10.5,
    GREY333, italic=True)
footer(s2)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(BLANK)
chrome(s3, "THE RECOMMENDATION")

left_x = CONTENT_X0
left_w = Inches(7.45)
div_x = left_x + left_w + Inches(0.1)
right_x = div_x + Inches(0.15)
right_w = Inches(13.0) - right_x

y = Inches(1.25)
txt(s3, bc.PRIMARY_SERVICE_NAME, left_x, y, left_w, Inches(0.45), 23, BLACK, font=SERIF, bold=True)
y += Inches(0.5)
txt(s3, f"{bc.PRIMARY_SERVICE_PRICE} {bc.PRIMARY_SERVICE_CADENCE}", left_x, y, left_w, Inches(0.3),
    15.5, TEAL, bold=True)
y += Inches(0.42)
txt(s3, bc.PRIMARY_SERVICE_DESC, left_x, y, left_w, Inches(0.85), 11.5, DARK_TXT)
y += Inches(1.05)

step_h = Inches(1.3)
rect(s3, left_x, y, left_w, step_h, LIGHT_TEAL_BG, line=TEAL, line_w=Pt(1))
txt(s3, "Step one, answer a few quick questions", left_x + Inches(0.22), y + Inches(0.2),
    left_w - Inches(0.4), Inches(0.3), 13.5, TEAL, bold=True)
txt(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.22), y + Inches(0.58),
    left_w - Inches(0.4), Inches(0.3), 11.5, DARK_TXT)
txt(s3, bc.QUESTIONNAIRE_CLEAN, left_x + Inches(0.22), y + Inches(0.95), left_w - Inches(0.4),
    Inches(0.3), 13, TEAL, bold=True, hyperlink=bc.QUESTIONNAIRE_URL_HTML)
y += step_h + Inches(0.36)

cta_h = Inches(0.62)
rect(s3, left_x, y, left_w, cta_h, AMBER_D)
cta_box = txt(s3, "Purchase the suggested product now to get started", left_x, y, left_w, cta_h,
              13, BLACK, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
              hyperlink=bc.PRIMARY_SERVICE_STRIPE)
y += cta_h + Inches(0.36)

txt(s3, "Prefer a conversation first?", left_x, y, left_w, Inches(0.3), 11.5, DARK_TXT)
y += Inches(0.3)
txt(s3, "Book a complimentary discovery call", left_x, y, left_w, Inches(0.3), 12.5, TEAL, bold=True,
    hyperlink=bc.BOOKING_LINK)

rect(s3, div_x, Inches(1.25), Pt(1.2), Inches(5.5), DIVIDER)

ry_ = Inches(1.25)
txt(s3, "NITESH ROOPA", right_x, ry_, right_w, Inches(0.35), 17, BLACK, font=SERIF, bold=True)
ry_ += Inches(0.36)
txt(s3, "CA, Managing Partner, ProfitPulse", right_x, ry_, right_w, Inches(0.3), 12, DARK_TXT)
ry_ += Inches(0.6)
CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
bullets(s3, CREDS, right_x, ry_, right_w, 11, DARK_TXT, leading=11, bullet_gap=14, height=Inches(1.9))
ry_ += Inches(1.9)
rect(s3, right_x, ry_, right_w, Pt(1), DIVIDER)
ry_ += Inches(0.3)
CONTACT = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for cl in CONTACT:
    txt(s3, cl, right_x, ry_, right_w, Inches(0.28), 11.5, DARK_TXT)
    ry_ += Inches(0.33)

footer(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMedia_17Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
