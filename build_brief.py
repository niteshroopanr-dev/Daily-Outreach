"""
ProfitPulse Brief Builder, v3.3 house style
Target: Taskforce Australia | Date: 07 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# Brand colours, exactly seven
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

SERIF = "Cambria"
SANS  = "Calibri"

NEXT_DAY = "07 Aug 2026"
COMPANY  = "Taskforce Australia"


def _add_alpha(srgbClr_el, alpha_pct):
    from pptx.oxml.ns import qn as _qn
    from lxml import etree
    alpha_el = etree.SubElement(srgbClr_el, _qn('a:alpha'))
    alpha_el.set('val', str(int(alpha_pct * 1000)))


def set_run_alpha(run, alpha_pct):
    """Reduce opacity of a text run's colour without introducing a new hex colour."""
    rPr = run._r.find(qn('a:rPr'))
    srgb = rPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
    _add_alpha(srgb, alpha_pct)


def set_shape_alpha(shape, alpha_pct):
    """Reduce opacity of a shape's fill without introducing a new hex colour."""
    srgb = shape.fill.fore_color._xFill.find(qn('a:srgbClr'))
    _add_alpha(srgb, alpha_pct)


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None, alpha=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.shadow.inherit = False
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if alpha is not None:
        set_shape_alpha(shape, alpha)
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height,
             font_name=SANS, font_size=12, bold=False,
             colour=BLACK, align=PP_ALIGN.LEFT, wrap=True,
             italic=False, anchor=None, shrink=True, alpha=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    if shrink:
        # shrink text on overflow as a safety net only
        from pptx.enum.text import MSO_AUTO_SIZE
        tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor:
        tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    if alpha is not None:
        set_run_alpha(run, alpha)
    return box


def add_multiline(slide, lines, left, top, width, height,
                   font_name=SANS, default_size=12, default_colour=BLACK,
                   align=PP_ALIGN.LEFT, line_spacing=None, space_after=4):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        cfg = {"text": line, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False, "alpha": None} if isinstance(line, str) else {
            "text": line.get("text", ""), "size": line.get("size", default_size),
            "colour": line.get("colour", default_colour), "bold": line.get("bold", False),
            "italic": line.get("italic", False), "alpha": line.get("alpha", None),
        }
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if space_after is not None:
            p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]
        if cfg["alpha"] is not None:
            set_run_alpha(run, cfg["alpha"])
    return box


def add_hyperlink_text(slide, text, url, left, top, width, height,
                        font_name=SANS, font_size=11, bold=False,
                        colour=TEAL, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.hyperlink.address = url
    return box


def chrome(slide, eyebrow, company_label=None):
    """Fixed chrome: left accent stripe, black header band, footer line."""
    set_background(slide, WHITE)
    # Left accent stripe, amber, full height
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    # Header band, black, full width, 1 inch tall
    add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
              font_name=SANS, font_size=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, company_label or "PROFITPULSE", Inches(9.3), Inches(0.32), Inches(3.8), Inches(0.4),
              font_name=SANS, font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
    # Footer line
    add_rect(slide, Inches(0.35), Inches(7.05), Inches(12.6), Pt(0.75), BLACK, alpha=15)
    add_text(slide, "Prepared by Nitesh Roopa, CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.35), Inches(7.12), Inches(9.5), Inches(0.3),
              font_name=SANS, font_size=8, colour=BLACK, align=PP_ALIGN.LEFT, alpha=55)
    add_text(slide, NEXT_DAY, Inches(10.5), Inches(7.12), Inches(2.45), Inches(0.3),
              font_name=SANS, font_size=8, colour=BLACK, align=PP_ALIGN.RIGHT, alpha=55)


def stat_card(slide, left, top, width, height, number, label_lines, source):
    add_rect(slide, left, top, width, height, BLACK)
    add_rect(slide, left, top, width, Pt(4), TEAL)
    add_text(slide, number, left + Inches(0.15), top + Inches(0.16), width - Inches(0.3), Inches(0.5),
              font_name=SERIF, font_size=27, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_multiline(slide, label_lines, left + Inches(0.15), top + Inches(0.68),
                  width - Inches(0.3), Inches(0.55),
                  font_name=SANS, default_size=10.5, default_colour=OFF_WHITE, space_after=0)
    add_text(slide, source, left + Inches(0.15), top + height - Inches(0.28), width - Inches(0.3), Inches(0.24),
              font_name=SANS, font_size=7.5, colour=OFF_WHITE, align=PP_ALIGN.LEFT, italic=True, alpha=55)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

add_text(s1, COMPANY, Inches(0.35), Inches(1.15), Inches(9.5), Inches(0.75),
          font_name=SERIF, font_size=40, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s1, "Property compliance and maintenance network, Burnley, Melbourne VIC",
          Inches(0.35), Inches(1.88), Inches(12.4), Inches(0.35),
          font_name=SANS, font_size=13, colour=BLACK, align=PP_ALIGN.LEFT, alpha=65)

# Stat card row: 6 cards, y=2.4, height=1.6, evenly spaced
cards = [
    ("$12.8M", ["Revenue, FY2025", "reported"], "SmartCompany Smart50 2025"),
    ("31%", ["Revenue growth,", "year on year"], "SmartCompany Smart50 2025"),
    ("19", ["Internal team", "members"], "SmartCompany Smart50 2025"),
    ("5,000", ["Tradespeople in", "national network"], "Taskforce, Real plus release"),
    ("500", ["Real estate agencies", "served nationally"], "PropertyMe integrator profile"),
    ("2014", ["Founded, Burnley,", "Melbourne VIC"], "SmartCompany Smart50 2025"),
]
n = len(cards)
card_w = Inches(2.03)
gap = Inches(0.15)
total_w = card_w * n + gap * (n - 1)
start_x = (W - total_w) / 2
card_top = Inches(2.4)
card_h = Inches(1.6)
for i, (num, lbl, src) in enumerate(cards):
    x = int(start_x + i * (card_w + gap))
    stat_card(s1, x, card_top, card_w, card_h, num, lbl, src)

# Key Commercial Signals
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.35), Inches(6), Inches(0.3),
          font_name=SANS, font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    "Revenue grew 31 percent year on year to $12.8 million, Smart50 2025 rank 37.",
    "New 2025 partnership with Real plus extends reach across property agencies nationally.",
    "Platform integrates with PropertyMe, MRI Property Tree and Console, serving 500 agencies.",
    "National network of 5,000 tradespeople delivers compliance, maintenance and warranty work.",
    "Named Victorian State Winner for Outstanding Growth, 2024 Telstra Best of Business Awards.",
    "Winner, Proptech Association Awards 2024, for the RentSafe and RentRepair platform.",
]
sig_lines = []
for sline in signals:
    sig_lines.append({"text": "▪  " + sline, "size": 11.5, "colour": BLACK})
add_multiline(s1, sig_lines, Inches(0.35), Inches(4.72), Inches(12.6), Inches(2.2),
              font_name=SANS, default_size=11.5, space_after=6)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY, THREE COMMERCIAL OBSERVATIONS
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY", COMPANY.upper())

add_text(s2, "Three commercial observations from ProfitPulse",
          Inches(0.35), Inches(1.12), Inches(12.6), Inches(0.4),
          font_name=SERIF, font_size=20, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)

col_top = Inches(1.7)
col_h = Inches(4.75)
col_w = Inches(4.03)
gap2 = Inches(0.12)
col_fills = [TEAL, BLACK, GOLD]
col_text_colours = [WHITE, OFF_WHITE, BLACK]
col_index_colours = [BLACK, TEAL, BLACK]

observations = [
    ("01", "Three lines, three margins",
     "Taskforce runs three lines: RentSafe compliance checks, RentRepair maintenance and manufacturer warranty servicing. Each carries a different cost base and margin profile. Past $12.8 million in revenue, the real question is whether every line is growing profit as fast as it is growing revenue."),
    ("02", "A lean team, a wide network",
     "Nineteen internal staff coordinate a national network of 5,000 tradespeople across compliance, maintenance and warranty work. That leverage is a real asset, and it also means margin now lives in how well each job type and service line is priced and tracked at scale."),
    ("03", "Three new doors just opened",
     "The 2025 partnership with Real plus, alongside integrations with PropertyMe, MRI Property Tree and Console, opens the platform to more of the 500 agencies it already serves. New distribution is only as valuable as the margin behind the services it sells."),
]

for i, (idx, head, body) in enumerate(observations):
    x = Inches(0.35) + i * (col_w + gap2)
    add_rect(s2, x, col_top, col_w, col_h, col_fills[i])
    add_text(s2, idx, x + Inches(0.25), col_top + Inches(0.2), col_w - Inches(0.5), Inches(0.7),
              font_name=SERIF, font_size=34, bold=True, colour=col_index_colours[i], align=PP_ALIGN.LEFT)
    add_text(s2, head, x + Inches(0.25), col_top + Inches(0.95), col_w - Inches(0.5), Inches(0.7),
              font_name=SERIF, font_size=15, bold=True, colour=col_text_colours[i], align=PP_ALIGN.LEFT)
    add_multiline(s2, [body], x + Inches(0.25), col_top + Inches(1.65), col_w - Inches(0.5), Inches(2.9),
                  font_name=SANS, default_size=10.5, default_colour=col_text_colours[i], space_after=0)

add_text(s2,
          "These observations are offered in good faith. Taskforce has built something genuinely "
          "impressive since 2014. The question is simply whether the margin behind each service "
          "line is as clear as the growth number.",
          Inches(0.35), Inches(6.55), Inches(12.6), Inches(0.45),
          font_name=SANS, font_size=10.5, italic=True, colour=BLACK, align=PP_ALIGN.LEFT, alpha=65)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION", COMPANY.upper())

# Left column
lx = Inches(0.35)
lw = Inches(7.6)

add_text(s3, "Product and Service Line Profitability", lx, Inches(1.15), lw, Inches(0.55),
          font_name=SERIF, font_size=21, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "$3,950 one off", lx, Inches(1.68), lw, Inches(0.4),
          font_name=SANS, font_size=16, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_multiline(s3, [
    "A three week project ranking every product or service line by gross margin, "
    "contribution margin and operational drag. For Taskforce, that means a clear, "
    "sourced view of what RentSafe, RentRepair and manufacturer warranty work each "
    "contribute once network payouts, platform costs and service time are counted."
], lx, Inches(2.12), lw, Inches(1.0), font_name=SANS, default_size=11, default_colour=BLACK, space_after=0)

# Step one block
add_rect(s3, lx, Inches(3.25), lw, Inches(1.25), OFF_WHITE)
add_rect(s3, lx, Inches(3.25), Inches(0.06), Inches(1.25), TEAL)
add_text(s3, "Step one, answer a few quick questions", lx + Inches(0.25), Inches(3.4), lw - Inches(0.5), Inches(0.35),
          font_name=SANS, font_size=13, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry.", lx + Inches(0.25), Inches(3.72), lw - Inches(0.5), Inches(0.3),
          font_name=SANS, font_size=11, colour=BLACK, align=PP_ALIGN.LEFT, alpha=75)
add_hyperlink_text(s3, "profit-pulse.com.au/services/find-your-fit",
                    "https://profit-pulse.com.au/services/find-your-fit/",
                    lx + Inches(0.25), Inches(4.02), lw - Inches(0.5), Inches(0.35),
                    font_size=12, bold=True, colour=TEAL)

# Direct CTA
add_rect(s3, lx, Inches(4.7), lw, Inches(0.55), AMBER_D)
add_hyperlink_text(s3, "Purchase the suggested product now to get started",
                    "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D",
                    lx + Inches(0.25), Inches(4.85), lw - Inches(0.5), Inches(0.3),
                    font_size=13, bold=True, colour=BLACK)

# Booking
add_text(s3, "Prefer a conversation first?", lx, Inches(5.45), lw, Inches(0.3),
          font_name=SANS, font_size=11, colour=BLACK, align=PP_ALIGN.LEFT, alpha=75)
add_hyperlink_text(s3, "Book a complimentary discovery call",
                    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
                    lx, Inches(5.75), lw, Inches(0.3), font_size=11, bold=True, colour=TEAL)

# Right column: credibility panel
rx = Inches(8.25)
rw = Inches(4.73)
add_rect(s3, rx, Inches(1.15), rw, Inches(5.6), BLACK)
add_rect(s3, rx, Inches(1.15), rw, Pt(4), TEAL)

add_text(s3, "Nitesh Roopa", rx + Inches(0.28), Inches(1.4), rw - Inches(0.56), Inches(0.4),
          font_name=SERIF, font_size=18, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.28), Inches(1.82), rw - Inches(0.56), Inches(0.35),
          font_name=SANS, font_size=12, colour=WHITE, align=PP_ALIGN.LEFT)

add_rect(s3, rx + Inches(0.28), Inches(2.28), rw - Inches(0.56), Pt(1), TEAL)

cred_lines = [
    {"text": "16 years of experience across 4 countries", "size": 11, "colour": OFF_WHITE},
    {"text": "52 deals executed and managed across the career", "size": 11, "colour": OFF_WHITE},
    {"text": "Largest single deal USD 1.3 billion, Cahora Bassa", "size": 11, "colour": OFF_WHITE},
    {"text": "Total GRBT project value over AUD 10 billion", "size": 11, "colour": OFF_WHITE},
]
add_multiline(s3, cred_lines, rx + Inches(0.28), Inches(2.45), rw - Inches(0.56), Inches(1.7),
              font_name=SANS, default_size=11, space_after=8)

add_rect(s3, rx + Inches(0.28), Inches(4.25), rw - Inches(0.56), Pt(1), OFF_WHITE, alpha=25)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 11, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 11, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 11, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10, "colour": OFF_WHITE},
]
add_multiline(s3, contact_lines, rx + Inches(0.28), Inches(4.45), rw - Inches(0.56), Inches(1.5),
              font_name=SANS, default_size=11, space_after=6)


# ══════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════
out_dir = "/home/user/Daily-Outreach/Out-reach efforts"
out_path = out_dir + "/Brief_TaskforceAustralia_07Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
