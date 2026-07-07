"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 08 Jul 2026
Three slide prospect facing deck. House style per Section 6.0/6.0A: white body,
black header band, amber left stripe, black stat cards with teal top edge.
Brand colours only. Zero dashes in authored copy.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

# Brand colours
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STR = "08 Jul 2026"
COMPANY = "Taskforce Australia"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, font_name="Calibri",
             font_size=12, bold=False, colour=BLACK, align=PP_ALIGN.LEFT,
             italic=False, serif=False, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = "Georgia" if serif else font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return box


def add_paragraphs(slide, paras, left, top, width, height, font_name="Calibri",
                    font_size=12, colour=BLACK, align=PP_ALIGN.LEFT,
                    line_spacing=1.15, space_after=6, bold=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    first = True
    for para_text in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = para_text
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.color.rgb = colour
        run.font.bold = bold
    return box


def add_link_hotspot(slide, left, top, width, height, url):
    """Invisible clickable rectangle. Keeps hyperlink click behaviour separate
    from text run styling, since LibreOffice and some PowerPoint themes force
    hyperlinked run text to a theme blue and underline it regardless of any
    explicit run colour, which is off brand."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.background()
    shape.line.fill.background()
    shape.shadow.inherit = False
    shape.click_action.hyperlink.address = url
    return shape


def add_hyperlink_text(slide, text, url, left, top, width, height,
                        font_size=13, colour=TEAL, bold=True, align=PP_ALIGN.LEFT):
    box = add_text(slide, text, left, top, width, height, font_size=font_size,
                    colour=colour, bold=bold, align=align)
    add_link_hotspot(slide, left, top, width, height, url)
    return box


def add_chrome(slide, eyebrow, header_right="PROFITPULSE"):
    """Fixed chrome per Section 6.0A: left stripe, header band, footer line."""
    set_background(slide, WHITE)
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
              font_size=12, bold=True, colour=OFF_WHITE)
    add_text(slide, header_right, Inches(9.5), Inches(0.32), Inches(3.6), Inches(0.4),
              font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
    # Footer
    add_rect(slide, Inches(0.3), Inches(7.02), Inches(12.73), Pt(0.75), BLACK)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.3), Inches(7.08), Inches(9.5), Inches(0.35), font_size=8,
              colour=BLACK)
    add_text(slide, DATE_STR, Inches(10.5), Inches(7.08), Inches(2.53), Inches(0.35),
              font_size=8, colour=BLACK, align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
add_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, COMPANY, Inches(0.3), Inches(1.1), Inches(10), Inches(0.72),
          font_size=40, bold=True, colour=BLACK, serif=True)
add_text(s1, "Property technology and trade services platform, Burnley, Melbourne VIC",
          Inches(0.3), Inches(1.82), Inches(11.5), Inches(0.33), font_size=13, colour=TEAL, bold=True)

# Stat cards row: 6 cards
cards = [
    ("$12.8M", "Verified 2025 revenue", "SmartCompany Smart50"),
    ("31%", "One year revenue growth", "SmartCompany Smart50"),
    ("19", "Employees, flat two years", "SmartCompany Smart50"),
    ("#37", "Smart50 2025 national rank", "SmartCompany Smart50"),
    ("4,500", "Tradespeople in the network", "taskforce.com.au"),
    ("2014", "Year founded in Melbourne", "SmartCompany Smart50"),
]
card_y = Inches(2.28)
card_h = Inches(1.72)
margin = 0.3
gap = 0.15
n = len(cards)
total_w = 13.333 - 2 * margin
card_w_in = (total_w - (n - 1) * gap) / n
for i, (num, label, source) in enumerate(cards):
    x = Inches(margin + i * (card_w_in + gap))
    cw = Inches(card_w_in)
    add_rect(s1, x, card_y, cw, card_h, BLACK)
    add_rect(s1, x, card_y, cw, Pt(4), TEAL)
    num_box = s1.shapes.add_textbox(x + Inches(0.1), card_y + Inches(0.16), cw - Inches(0.2), Inches(0.48))
    num_tf = num_box.text_frame
    num_tf.word_wrap = False
    num_tf.auto_size = MSO_AUTO_SIZE.NONE
    num_p = num_tf.paragraphs[0]
    num_run = num_p.add_run()
    num_run.text = num
    num_run.font.size = Pt(27)
    num_run.font.bold = True
    num_run.font.color.rgb = AMBER_B
    num_run.font.name = "Calibri"
    add_paragraphs(s1, [label], x + Inches(0.14), card_y + Inches(0.65), cw - Inches(0.28), Inches(0.55),
                   font_size=11, colour=OFF_WHITE, line_spacing=1.05, space_after=0)
    add_text(s1, source, x + Inches(0.14), card_y + Inches(1.38), cw - Inches(0.28), Inches(0.26),
              font_size=8, colour=OFF_WHITE)

# Lower zone: Key Commercial Signals, full width, one line each
lower_y = Inches(4.32)
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.3), lower_y, Inches(8), Inches(0.3),
          font_size=12, bold=True, colour=TEAL)

signals = [
    "Revenue nearly doubled, 7.43 million to 12.8 million, in two years on flat headcount of 19.",
    "Public target of 50 to 60 percent revenue growth a year for three years, funded organically.",
    "RentSafe has completed over 140,000 jobs across 20 consumer brands and 300 real estate offices.",
    "RentRepair adds a new subscription revenue line at 59 dollars a month per property.",
    "Victorian State Winner, Outstanding Growth, 2024 Telstra Best of Business Awards.",
    "Winner, 2024 Proptech Association Awards, among 125 competing companies nationally.",
]
sig_box = s1.shapes.add_textbox(Inches(0.3), lower_y + Inches(0.38), Inches(12.73), Inches(2.15))
tf = sig_box.text_frame
tf.word_wrap = True
tf.auto_size = MSO_AUTO_SIZE.NONE
first = True
for sgl in signals:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.line_spacing = 1.05
    p.space_after = Pt(6)
    r1 = p.add_run()
    r1.text = "•  "
    r1.font.size = Pt(12)
    r1.font.color.rgb = AMBER_D
    r1.font.bold = True
    r2 = p.add_run()
    r2.text = sgl
    r2.font.size = Pt(12)
    r2.font.color.rgb = BLACK

add_text(s1, "Source: SmartCompany Smart50 award citations 2023 to 2025; Telstra Best of Business Awards 2024; taskforce.com.au",
          Inches(0.3), Inches(6.72), Inches(12.73), Inches(0.25), font_size=7.5, colour=BLACK)

# ════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY, THREE COMMERCIAL OBSERVATIONS
# ════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
add_chrome(s2, "THE OPPORTUNITY")

add_text(s2, f"{COMPANY}: three commercial observations from ProfitPulse",
          Inches(0.3), Inches(1.15), Inches(12.7), Inches(0.4),
          font_size=15, bold=True, colour=BLACK, serif=True)

col_y = Inches(1.75)
col_h = Inches(4.55)
col_gap = 0.14
col_w_in = (13.333 - 2 * 0.3 - 2 * col_gap) / 3
observations = [
    (TEAL, BLACK, "01", "Three products, one blended number",
     "TradieConnect prices by the job, RentSafe runs on real estate office "
     "contracts, and RentRepair charges a flat 59 dollars a month. The Smart50 "
     "citation reports one blended figure across all three, 12.8 million "
     "dollars and 31 percent growth. That number cannot show which line is "
     "funding the next stage, and which is being carried by the other two."),
    (BLACK, OFF_WHITE, "02", "Flat headcount, rising revenue",
     "Revenue rose from 7.43 million in 2023 to 12.8 million in 2025 while "
     "Smart50 citations show headcount unchanged at 19. That is a genuine "
     "productivity gain, but those same 19 people now run three product "
     "lines instead of one. Without a line by line view, it is hard to know "
     "if that gain is spread evenly or concentrated in a single line."),
    (GOLD, BLACK, "03", "A target without a stated plan",
     "Taskforce has told SmartCompany it is targeting 50 to 60 percent "
     "revenue growth a year for three years, funded organically rather than "
     "through a further raise. No public source breaks that target down by "
     "product line or states where the investment goes first. A costed, "
     "line by line view would turn the public ambition into a working plan."),
]
for i, (fill, txt_colour, idx, header, para) in enumerate(observations):
    x = Inches(0.3 + i * (col_w_in + col_gap))
    cw = Inches(col_w_in)
    add_rect(s2, x, col_y, cw, col_h, fill)
    add_text(s2, idx, x + Inches(0.25), col_y + Inches(0.22), cw - Inches(0.5), Inches(0.75),
              font_size=40, bold=True, colour=txt_colour)
    add_text(s2, header, x + Inches(0.25), col_y + Inches(1.0), cw - Inches(0.5), Inches(0.75),
              font_size=15, bold=True, colour=txt_colour, line_spacing=1.05)
    add_paragraphs(s2, [para], x + Inches(0.25), col_y + Inches(1.75), cw - Inches(0.5), Inches(2.6),
                   font_size=11, colour=txt_colour, line_spacing=1.18, space_after=0)

add_paragraphs(
    s2,
    ["These observations are offered in good faith. Taskforce has built something genuinely "
     "fast growing. The question is simply whether the margin behind each line matches the "
     "ambition in front of it."],
    Inches(0.3), Inches(6.4), Inches(12.73), Inches(0.55),
    font_size=11, colour=BLACK, line_spacing=1.1,
)

# ════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
add_chrome(s3, "THE RECOMMENDATION")

left_x = Inches(0.3)
left_w = Inches(7.6)

add_text(s3, "Product and Service Line Profitability", left_x, Inches(1.15), left_w, Inches(0.55),
          font_size=22, bold=True, colour=BLACK, serif=True)
add_text(s3, "$3,950 one off", left_x, Inches(1.68), left_w, Inches(0.4),
          font_size=18, bold=True, colour=AMBER_D)
add_paragraphs(
    s3,
    ["Ranks TradieConnect, RentSafe and RentRepair by gross margin, contribution margin "
     "and operational drag, so growth effort follows the line that actually pays."],
    left_x, Inches(2.12), left_w, Inches(0.75), font_size=12, colour=BLACK, line_spacing=1.15,
)

# Step one block
step_y = Inches(2.95)
add_rect(s3, left_x, step_y, left_w, Inches(1.55), OFF_WHITE, line_colour=TEAL, line_width=Pt(1))
add_text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), step_y + Inches(0.14),
          left_w - Inches(0.4), Inches(0.35), font_size=13, bold=True, colour=TEAL)
add_text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.2), step_y + Inches(0.52),
          left_w - Inches(0.4), Inches(0.3), font_size=11, colour=BLACK)
add_hyperlink_text(s3, "profit-pulse.com.au/services/find-your-fit",
                    "https://profit-pulse.com.au/services/find-your-fit/",
                    left_x + Inches(0.2), step_y + Inches(0.88), left_w - Inches(0.4), Inches(0.35),
                    font_size=13, colour=TEAL, bold=True)

# Direct CTA
cta_y = Inches(4.68)
add_rect(s3, left_x, cta_y, left_w, Inches(0.95), BLACK)
add_hyperlink_text(s3, "Purchase the suggested product now to get started",
                    "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
                    left_x + Inches(0.25), cta_y + Inches(0.28), left_w - Inches(0.5), Inches(0.4),
                    font_size=14, colour=AMBER_B, bold=True, align=PP_ALIGN.LEFT)

add_text(s3, "Prefer a conversation first?", left_x, Inches(5.85), left_w, Inches(0.32),
          font_size=12, colour=BLACK)
add_hyperlink_text(s3, "Book a complimentary discovery call",
                    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
                    left_x, Inches(6.18), left_w, Inches(0.35), font_size=12, colour=TEAL, bold=True)

# Right column: About ProfitPulse and Nitesh
right_x = Inches(8.15)
right_w = Inches(4.88)
right_h = Inches(4.75)
add_rect(s3, right_x, Inches(1.15), right_w, right_h, BLACK)
add_rect(s3, right_x, Inches(1.15), right_w, Pt(4), TEAL)

add_text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.4), right_w - Inches(0.5), Inches(0.45),
          font_size=19, bold=True, colour=AMBER_B, serif=True)
add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.88), right_w - Inches(0.5), Inches(0.35),
          font_size=12, colour=WHITE)

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
cy = Inches(2.42)
for cl in cred_lines:
    add_text(s3, cl, right_x + Inches(0.25), cy, right_w - Inches(0.5), Inches(0.35),
              font_size=11, colour=OFF_WHITE)
    cy += Inches(0.38)

add_rect(s3, right_x + Inches(0.25), cy + Inches(0.06), right_w - Inches(0.5), Pt(1), TEAL)
cy += Inches(0.32)

contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for cl in contact_lines:
    add_text(s3, cl, right_x + Inches(0.25), cy, right_w - Inches(0.5), Inches(0.35),
              font_size=11, colour=OFF_WHITE)
    cy += Inches(0.34)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_08Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
