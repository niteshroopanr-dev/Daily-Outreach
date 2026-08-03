"""
ProfitPulse Brief Builder
Target: Paire | Date: 04 Aug 2026
Three slide prospect facing deck. House style per Section 6. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STR = "04 Aug 2026"
COMPANY = "Paire"

HEADING_FONT = "Cambria"
BODY_FONT = "Calibri"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    shape.shadow.inherit = False
    if line_colour:
        shape.line.color.rgb = line_colour
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height, font_name=BODY_FONT, font_size=12,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, italic=False, anchor=None,
             line_spacing=None, shrink=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return box


def add_multiline(slide, lines, left, top, width, height, font_name=BODY_FONT,
                   default_size=12, default_colour=BLACK, default_bold=False,
                   align=PP_ALIGN.LEFT, spacing_after=4, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size, "colour": default_colour,
                   "bold": default_bold, "italic": False}
        else:
            cfg = {"text": line.get("text", ""), "size": line.get("size", default_size),
                   "colour": line.get("colour", default_colour), "bold": line.get("bold", default_bold),
                   "italic": line.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if spacing_after:
            p.space_after = Pt(spacing_after)
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
    return box


def add_chrome(slide, eyebrow, right_label=None):
    """Left accent stripe, black header band, footer line. right_label defaults to PROFITPULSE."""
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    add_rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.30), Inches(8.5), Inches(0.4),
              font_name=BODY_FONT, font_size=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, right_label or "PROFITPULSE", Inches(9.0), Inches(0.30), Inches(4.0), Inches(0.4),
              font_name=BODY_FONT, font_size=12, bold=True, colour=AMBER_B, align=PP_ALIGN.RIGHT)
    # footer
    add_rect(slide, Inches(0.35), Inches(7.05), Inches(12.6), Pt(1), RGBColor(0xCC, 0xCC, 0xCC))
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.35), Inches(7.10), Inches(9.0), Inches(0.3),
              font_size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.LEFT)
    add_text(slide, DATE_STR, Inches(10.5), Inches(7.10), Inches(2.4), Inches(0.3),
              font_size=8, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank_layout = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank_layout)
set_background(s1, WHITE)
add_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

add_text(s1, "Paire", Inches(0.35), Inches(1.15), Inches(7.0), Inches(0.85),
          font_name=HEADING_FONT, font_size=40, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s1, "Direct to consumer apparel brand, South Melbourne VIC",
          Inches(0.35), Inches(1.90), Inches(9.5), Inches(0.4),
          font_name=BODY_FONT, font_size=13, colour=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.LEFT)

# Stat cards row: 5 cards, y=2.4, height 1.6, width 2.3, gap 0.15
stat_cards = [
    ("$9.4M", "FY2025 revenue", "SmartCompany Smart50 2025"),
    ("64%", "FY2025 revenue growth", "SmartCompany Smart50 2025"),
    ("#15", "Smart50 2025 national rank", "SmartCompany, 2025"),
    ("2020", "Year founded", "FashionUnited 2026"),
    ("40", "People on the team", "FashionUnited 2026"),
]
card_w = Inches(2.3)
card_h = Inches(1.6)
card_gap = Inches(0.15)
card_y = Inches(2.4)
x = Inches(0.35)
for num, label, source in stat_cards:
    add_rect(s1, x, card_y, card_w, card_h, BLACK)
    add_rect(s1, x, card_y, card_w, Pt(4), TEAL)
    add_text(s1, num, x + Inches(0.15), card_y + Inches(0.18), card_w - Inches(0.3), Inches(0.55),
              font_name=HEADING_FONT, font_size=28, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_multiline(s1, [label], x + Inches(0.15), card_y + Inches(0.78), card_w - Inches(0.3), Inches(0.5),
                   default_size=11, default_colour=OFF_WHITE, spacing_after=0)
    add_text(s1, source, x + Inches(0.15), card_y + Inches(1.30), card_w - Inches(0.3), Inches(0.25),
              font_size=7, italic=True, colour=RGBColor(0x9A, 0x9A, 0x9A), align=PP_ALIGN.LEFT)
    x += card_w + card_gap

# Revenue trend chart (verified 3 data points) — simple bar chart, y 4.25 to 5.85
add_text(s1, "REVENUE TRAJECTORY, SMART50 AWARD CITATIONS", Inches(0.35), Inches(4.30), Inches(6.0), Inches(0.3),
          font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
chart_data = [("Smart50 2023", 2.15), ("Smart50 2024", 5.9), ("Smart50 2025", 9.4)]
max_val = 9.4
bar_area_x = Inches(0.35)
bar_area_w = Inches(6.0)
bar_w = Inches(1.5)
bar_gap = Inches(0.5)
base_y = Inches(6.55)
max_bar_h = Inches(1.5)
bx = bar_area_x
for label, val in chart_data:
    bar_h = Emu(int(max_bar_h * (val / max_val))) if False else Inches(1.5 * (val / max_val))
    add_rect(s1, bx, base_y - bar_h, bar_w, bar_h, TEAL)
    add_text(s1, f"${val}M", bx, base_y - bar_h - Inches(0.32), bar_w, Inches(0.3),
              font_size=11, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
    add_text(s1, label, bx, base_y + Inches(0.05), bar_w, Inches(0.3),
              font_size=8, colour=RGBColor(0x55, 0x55, 0x55), align=PP_ALIGN.CENTER)
    bx += bar_w + bar_gap

# Key Commercial Signals, right column
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(6.85), Inches(4.30), Inches(6.1), Inches(0.3),
          font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
signals = [
    "Smart50 2025 rank 15, revenue grew 64% to $9.4M (SmartCompany, 2025)",
    "Smart50 2024 double winner, Rising Star and Retail Award, 88% growth (SmartCompany)",
    "Pitched on Shark Tank Australia for a $500,000 equity investment (SmartCompany)",
    "Singapore distributor Seager Inc signed for Southeast Asia wholesale, 2023 (SmartCompany)",
    "Converted QV Melbourne pop up into a permanent flagship store (SmartCompany, 2025)",
    "Launched direct to consumer in the United States, 2026 range (FashionUnited)",
]
add_multiline(s1, signals, Inches(6.85), Inches(4.65), Inches(6.1), Inches(2.3),
               default_size=10.5, default_colour=BLACK, spacing_after=6, line_spacing=1.05)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_04Aug2026.pptx"

# ══════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════
from pptx.util import Emu
s2 = prs.slides.add_slide(blank_layout)
set_background(s2, WHITE)
add_chrome(s2, "THE OPPORTUNITY")
add_text(s2, "Paire: three commercial observations from ProfitPulse", Inches(0.35), Inches(1.10), Inches(12.5), Inches(0.4),
          font_name=HEADING_FONT, font_size=17, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)

col_y = Inches(1.70)
col_h = Inches(4.9)
col_w = Inches(4.05)
col_gap = Inches(0.09)
cols = [
    (TEAL, WHITE, "01", "Three growth engines, one capital plan",
     "Paire is scaling revenue at 64% a year while running three capital hungry expansions at once: "
     "a permanent flagship plus planned Sydney stores, a new Singapore wholesale relationship, and a "
     "2026 direct to consumer launch in the United States. Each draws on the same pool of cash and "
     "management time. No public evidence points to a single, costed plan that sequences these three "
     "bets against the capacity of the business to fund them."),
    (BLACK, WHITE, "02", "A funding question still open",
     "The Shark Tank Australia pitch shows the founders have already tested appetite for outside equity, "
     "reportedly seeking $500,000, but no source confirms a completed deal. Meanwhile headcount has grown "
     "from roughly ten to a stated forty. Whether the next stage of growth is funded from trading cash flow, "
     "debt, or equity is an open question that shapes every other decision this year."),
    (GOLD, BLACK, "03", "Structure has not caught up with scale",
     "Paire now sells across Australia, Singapore, and the United States, from a single Melbourne base whose "
     "exact corporate structure for the US market is not publicly disclosed. As revenue heads toward eight "
     "figures, the reporting rhythm and capital allocation discipline that suited a two million dollar startup "
     "may not suit a multi country retail and wholesale business."),
]
cx = Inches(0.35)
for fill, txt_colour, idx, header, body in cols:
    add_rect(s2, cx, col_y, col_w, col_h, fill)
    add_text(s2, idx, cx + Inches(0.25), col_y + Inches(0.20), col_w - Inches(0.5), Inches(0.7),
              font_name=HEADING_FONT, font_size=30, bold=True, colour=txt_colour, align=PP_ALIGN.LEFT)
    add_text(s2, header, cx + Inches(0.25), col_y + Inches(0.95), col_w - Inches(0.5), Inches(0.75),
              font_name=HEADING_FONT, font_size=14, bold=True, colour=txt_colour, align=PP_ALIGN.LEFT)
    add_text(s2, body, cx + Inches(0.25), col_y + Inches(1.75), col_w - Inches(0.5), Inches(2.9),
              font_name=BODY_FONT, font_size=10.5, colour=txt_colour, align=PP_ALIGN.LEFT, line_spacing=1.08)
    cx += col_w + col_gap

add_text(s2, "These are observations offered in good faith. Paire has built something impressive. "
              "The question is simply whether the financial architecture matches the ambition.",
          Inches(0.35), Inches(6.66), Inches(12.6), Inches(0.3),
          font_size=10.5, italic=True, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)

# ══════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank_layout)
set_background(s3, WHITE)
add_chrome(s3, "THE RECOMMENDATION")

# LEFT COLUMN
lx = Inches(0.35)
lw = Inches(7.4)
add_text(s3, "Strategic Growth Diagnostic", lx, Inches(1.15), lw, Inches(0.55),
          font_name=HEADING_FONT, font_size=22, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "$5,000 one off", lx, Inches(1.68), lw, Inches(0.4),
          font_name=BODY_FONT, font_size=15, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, "A six week engagement that maps revenue, capacity, and margin headroom across the "
              "retail, wholesale, and United States channels, then produces a 12 month growth plan "
              "with funding and capital allocation steps spelled out for each.",
          lx, Inches(2.10), lw, Inches(1.05), font_size=11.5, colour=BLACK, line_spacing=1.1)

add_rect(s3, lx, Inches(3.25), lw, Inches(1.35), RGBColor(0xF2, 0xF2, 0xF0), line_colour=TEAL, line_width=Pt(1))
add_text(s3, "Step one, answer a few quick questions", lx + Inches(0.2), Inches(3.35), lw - Inches(0.4), Inches(0.35),
          font_size=12.5, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry.", lx + Inches(0.2), Inches(3.70), lw - Inches(0.4), Inches(0.3),
          font_size=10.5, colour=BLACK, align=PP_ALIGN.LEFT)
link1 = add_text(s3, "profit-pulse.com.au/services/find-your-fit", lx + Inches(0.2), Inches(4.00), lw - Inches(0.4), Inches(0.35),
          font_size=11.5, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
link1.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://profit-pulse.com.au/services/find-your-fit/"

add_rect(s3, lx, Inches(4.80), lw, Inches(0.65), AMBER_B)
cta = add_text(s3, "Purchase the suggested product now to get started", lx, Inches(4.97), lw, Inches(0.35),
          font_size=13, bold=True, colour=BLACK, align=PP_ALIGN.CENTER)
cta.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"

add_text(s3, "Prefer a conversation first? Book a complimentary discovery call.", lx, Inches(5.65), lw, Inches(0.35),
          font_size=10.5, colour=BLACK, align=PP_ALIGN.LEFT)
link2 = add_text(s3, "bookings.cloud.microsoft/book/ProfitPulse1", lx, Inches(5.98), lw, Inches(0.35),
          font_size=10.5, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
link2.text_frame.paragraphs[0].runs[0].hyperlink.address = (
    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

# RIGHT COLUMN: About panel
rx = Inches(8.05)
rw = Inches(4.9)
add_rect(s3, rx, Inches(1.15), rw, Inches(5.75), BLACK)
add_rect(s3, rx, Inches(1.15), rw, Pt(4), TEAL)
add_text(s3, "Nitesh Roopa", rx + Inches(0.25), Inches(1.40), rw - Inches(0.5), Inches(0.45),
          font_name=HEADING_FONT, font_size=18, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.25), Inches(1.85), rw - Inches(0.5), Inches(0.35),
          font_size=11.5, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
add_rect(s3, rx + Inches(0.25), Inches(2.30), rw - Inches(0.5), Pt(1), TEAL)
credibility = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
add_multiline(s3, credibility, rx + Inches(0.25), Inches(2.65), rw - Inches(0.5), Inches(2.1),
               default_size=11, default_colour=OFF_WHITE, spacing_after=16, line_spacing=1.1)
add_rect(s3, rx + Inches(0.25), Inches(4.75), rw - Inches(0.5), Pt(1), TEAL)
contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
add_multiline(s3, contact_lines, rx + Inches(0.25), Inches(5.05), rw - Inches(0.5), Inches(1.9),
               default_size=11, default_colour=WHITE, spacing_after=16)

prs.save(out_path)
print(f"PPTX saved: {out_path}")
