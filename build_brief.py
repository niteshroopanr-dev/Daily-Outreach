"""
ProfitPulse Brief Builder
Target: Attekus | Date: 29 Aug 2026
Three slide prospect facing deck. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

# Brand colours (the only seven colours used anywhere in this deck)
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Liberation Serif"
SANS  = "Liberation Sans"

W = Inches(13.333)
H = Inches(7.5)

COMPANY = "Attekus"
DATE_STAMP = "29 Aug 2026"
QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL = "https://profit-pulse.com.au/services/find-your-fit/"
STRIPE_URL = "https://buy.stripe.com/cNi9AUd9e9iodlC8sn3ks0k"
BOOKING_URL = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour=None, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill_colour is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_colour
    else:
        shape.fill.background()
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, font_name=SANS, font_size=18,
             bold=False, colour=WHITE, align=PP_ALIGN.LEFT, wrap=True, italic=False,
             hyperlink=None, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
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
    if hyperlink:
        run.hyperlink.address = hyperlink
    return box


def add_multiline(slide, lines, left, top, width, height, font_name=SANS, default_size=14,
                   default_colour=OFF_WHITE, default_bold=False, align=PP_ALIGN.LEFT,
                   spacing_after=None, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size, "colour": default_colour,
                   "bold": default_bold, "italic": False, "hyperlink": None}
        else:
            cfg = {
                "text": line.get("text", ""), "size": line.get("size", default_size),
                "colour": line.get("colour", default_colour), "bold": line.get("bold", default_bold),
                "italic": line.get("italic", False), "hyperlink": line.get("hyperlink"),
            }
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if spacing_after is not None:
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
        if cfg["hyperlink"]:
            run.hyperlink.address = cfg["hyperlink"]
    return box


def footer(slide):
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.4), Inches(7.14), Inches(9.5), Inches(0.3),
              font_size=9, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, DATE_STAMP, Inches(10.6), Inches(7.14), Inches(2.3), Inches(0.3),
              font_size=9, colour=OFF_WHITE, align=PP_ALIGN.RIGHT)


def base_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(s, BLACK)
    add_rect(s, Inches(0), Inches(0), Inches(0.09), H, fill_colour=AMBER_D)
    return s


prs = Presentation()
prs.slide_width = W
prs.slide_height = H

# ============================================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ============================================================================
s1 = base_slide(prs)

add_text(s1, "COMMERCIAL INTELLIGENCE BRIEF", Inches(0.4), Inches(0.32), Inches(7), Inches(0.35),
          font_size=12, bold=True, colour=GOLD, align=PP_ALIGN.LEFT)
add_text(s1, "PROFITPULSE", Inches(9.5), Inches(0.32), Inches(3.4), Inches(0.35),
          font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

add_text(s1, COMPANY, Inches(0.4), Inches(0.72), Inches(9), Inches(1.0),
          font_name=SERIF, font_size=54, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

add_text(s1, "Cloud based booking and event software for councils, Brisbane based",
          Inches(0.4), Inches(1.68), Inches(11.5), Inches(0.4),
          font_size=16, colour=WHITE, align=PP_ALIGN.LEFT)

add_rect(s1, Inches(0.4), Inches(2.12), Inches(12.5), Pt(1.5), fill_colour=TEAL)

# Stat tiles
stats = [
    ("$9.2M", ["FY2025 revenue", "reported"], "getlatka.com, Sept 2025"),
    ("$5M",   ["First external", "capital raised"], "QIC Ventures, Oct 2024"),
    ("20%",   ["Of ANZ councils", "on Bookable"], "Five V Capital profile"),
    ("42",    ["People across", "the business"], "getlatka.com profile"),
    ("2017",  ["Year founded,", "Brisbane"], "Startup Daily report"),
    ("WINNER",["ANZ High Growth", "Award 2025"], "Lord Mayor Business Awards"),
]
tile_w = Inches(1.98)
tile_h = Inches(1.85)
gap = Inches(0.11)
x0 = Inches(0.4)
y0 = Inches(2.34)
for i, (num, label_lines, source) in enumerate(stats):
    x = Inches(x0.inches + i * (tile_w.inches + gap.inches))
    add_rect(s1, x, y0, tile_w, tile_h, fill_colour=BLACK, line_colour=OFF_WHITE, line_width=Pt(0.75))
    add_rect(s1, x, y0, tile_w, Inches(0.07), fill_colour=TEAL)
    add_text(s1, num, x + Inches(0.12), y0 + Inches(0.2), tile_w - Inches(0.24), Inches(0.55),
              font_name=SERIF, font_size=30, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_multiline(s1, label_lines, x + Inches(0.12), y0 + Inches(0.82), tile_w - Inches(0.24), Inches(0.6),
                  default_size=11, default_colour=OFF_WHITE, line_spacing=1.0)
    add_text(s1, source, x + Inches(0.12), y0 + Inches(1.5), tile_w - Inches(0.24), Inches(0.3),
              font_size=8, colour=GOLD, italic=True, align=PP_ALIGN.LEFT)

# Key commercial signals
add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.4), Inches(4.42), Inches(6), Inches(0.3),
          font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

signals = [
    "Raised AUD 5 million first ever external round, led by Five V Capital, October 2024.",
    "QIC Ventures joined as co investor in the same institutional funding round.",
    "Named 2025 ANZ High Growth Business at Brisbane Lord Mayor Business Awards.",
    "Bookable platform now used by close to 20 percent of councils across Australia and NZ.",
    "New capital earmarked for UK council entry plus ANZ education and government sectors.",
    "CEO Peter Suchting joined in 2023, leading beyond founder only leadership.",
]
sig_lines = [{"text": "•  " + t, "size": 13, "colour": OFF_WHITE} for t in signals]
add_multiline(s1, sig_lines, Inches(0.4), Inches(4.78), Inches(12.4), Inches(2.2),
              default_size=13, spacing_after=7)

footer(s1)

# ============================================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================================
s2 = base_slide(prs)

add_text(s2, "THE OPPORTUNITY", Inches(0.4), Inches(0.32), Inches(7), Inches(0.35),
          font_name=SERIF, font_size=24, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s2, "Attekus, three commercial observations from ProfitPulse",
          Inches(0.4), Inches(0.82), Inches(12.4), Inches(0.35),
          font_size=14, colour=WHITE, align=PP_ALIGN.LEFT)
add_rect(s2, Inches(0.4), Inches(1.22), Inches(12.5), Pt(1.5), fill_colour=TEAL)

columns = [
    {
        "idx": "01", "header": "New Capital, New Scrutiny", "fill": TEAL, "text_col": BLACK,
        "tag": "Aligned to the recommended service",
        "body": ("Five V Capital and QIC Ventures backed Attekus with its first ever "
                 "external raise in October 2024, funding entry into the United Kingdom "
                 "council market plus ANZ education and government agencies. A Capital "
                 "Allocation Review ranks each expansion path by expected return before "
                 "further capital is committed."),
    },
    {
        "idx": "02", "header": "A Board That Expects More", "fill": BLACK, "text_col": OFF_WHITE,
        "tag": None,
        "body": ("Peter Suchting joined as Chief Executive Officer in 2023, and two "
                 "institutional shareholders now sit alongside the founders. Reporting "
                 "built for founder only oversight rarely satisfies venture capital "
                 "standards. An Annual Plan and Board Pack lifts that reporting rhythm "
                 "to the standard Five V Capital and QIC Ventures now expect."),
    },
    {
        "idx": "03", "header": "Three Markets, One Plan", "fill": GOLD, "text_col": BLACK,
        "tag": None,
        "body": ("Attekus already serves close to 20 percent of councils across "
                 "Australia and New Zealand. The new capital targets three further "
                 "markets at once: the United Kingdom, education, and government "
                 "agencies. A Strategic Growth Diagnostic sequences that expansion "
                 "into a costed 12 month plan."),
    },
]

col_w = Inches(4.03)
col_gap = Inches(0.19)
col_y = Inches(1.5)
col_h = Inches(4.55)
for i, col in enumerate(columns):
    x = Inches(0.4 + i * (col_w.inches + col_gap.inches))
    line_col = OFF_WHITE if col["fill"] == BLACK else BLACK
    add_rect(s2, x, col_y, col_w, col_h, fill_colour=col["fill"],
             line_colour=(OFF_WHITE if col["fill"] == BLACK else None),
             line_width=Pt(0.75) if col["fill"] == BLACK else None)
    add_text(s2, col["idx"], x + Inches(0.22), col_y + Inches(0.2), Inches(1.5), Inches(0.6),
              font_name=SERIF, font_size=32, bold=True, colour=col["text_col"], align=PP_ALIGN.LEFT)
    add_text(s2, col["header"], x + Inches(0.22), col_y + Inches(0.85), col_w - Inches(0.44), Inches(0.7),
              font_name=SERIF, font_size=17, bold=True, colour=col["text_col"], align=PP_ALIGN.LEFT)
    body_top = col_y + Inches(1.55)
    if col["tag"]:
        add_text(s2, col["tag"].upper(), x + Inches(0.22), col_y + Inches(1.35), col_w - Inches(0.44), Inches(0.3),
                  font_size=9, bold=True, colour=col["text_col"], align=PP_ALIGN.LEFT)
        body_top = col_y + Inches(1.62)
    add_text(s2, col["body"], x + Inches(0.22), body_top, col_w - Inches(0.44), Inches(2.7),
              font_size=12.5, colour=col["text_col"], align=PP_ALIGN.LEFT, line_spacing=1.18)

add_text(s2, ("Seven years of disciplined, capital light growth is a rare foundation. ProfitPulse "
              "would welcome the chance to help Attekus put its first outside capital to its best "
              "possible use."),
          Inches(0.4), Inches(6.28), Inches(12.5), Inches(0.7),
          font_size=13, italic=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT, line_spacing=1.15)

footer(s2)

# ============================================================================
# SLIDE 3: THE RECOMMENDATION
# ============================================================================
s3 = base_slide(prs)

add_text(s3, "THE RECOMMENDATION", Inches(0.4), Inches(0.32), Inches(8), Inches(0.35),
          font_name=SERIF, font_size=24, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_rect(s3, Inches(0.4), Inches(0.82), Inches(12.5), Pt(1.5), fill_colour=TEAL)

# Left column
lx = Inches(0.4)
lw = Inches(7.4)

add_text(s3, "Capital Allocation Review", lx, Inches(1.05), lw, Inches(0.55),
          font_name=SERIF, font_size=27, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "AUD 6,000 one off", lx, Inches(1.62), lw, Inches(0.4),
          font_size=16, bold=True, colour=WHITE, align=PP_ALIGN.LEFT)
add_text(s3, ("An independent ranking of the United Kingdom, education and government "
              "expansion paths by expected return, with a redeployment plan the board can act on."),
          lx, Inches(2.08), lw, Inches(0.75),
          font_size=13, colour=OFF_WHITE, align=PP_ALIGN.LEFT, line_spacing=1.2)

add_rect(s3, lx, Inches(2.95), lw, Inches(1.55), fill_colour=BLACK, line_colour=TEAL, line_width=Pt(1))
add_text(s3, "STEP ONE", lx + Inches(0.22), Inches(3.1), lw - Inches(0.44), Inches(0.3),
          font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
add_text(s3, "See the solutions matched to your size and industry",
          lx + Inches(0.22), Inches(3.42), lw - Inches(0.44), Inches(0.4),
          font_size=13, colour=WHITE, align=PP_ALIGN.LEFT)
add_text(s3, QUESTIONNAIRE_CLEAN, lx + Inches(0.22), Inches(3.82), lw - Inches(0.44), Inches(0.5),
          font_size=14, bold=True, colour=TEAL, align=PP_ALIGN.LEFT, hyperlink=QUESTIONNAIRE_URL)

add_rect(s3, lx, Inches(4.65), lw, Inches(0.95), fill_colour=BLACK, line_colour=AMBER_D, line_width=Pt(1))
add_text(s3, "Purchase the suggested product now to get started",
          lx + Inches(0.22), Inches(4.85), lw - Inches(0.44), Inches(0.6),
          font_size=14, bold=True, colour=AMBER_D, align=PP_ALIGN.LEFT, hyperlink=STRIPE_URL)

add_text(s3, "Prefer a conversation first", lx, Inches(5.78), lw, Inches(0.35),
          font_size=12, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
add_text(s3, "Book a complimentary discovery call", lx, Inches(6.12), lw, Inches(0.4),
          font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.LEFT, hyperlink=BOOKING_URL)

# Right column: About ProfitPulse and Nitesh
rx = Inches(8.15)
rw = Inches(4.75)
add_rect(s3, rx, Inches(1.05), rw, Inches(5.55), fill_colour=BLACK, line_colour=OFF_WHITE, line_width=Pt(0.75))
add_rect(s3, rx, Inches(1.05), rw, Inches(0.06), fill_colour=GOLD)

add_text(s3, "Nitesh Roopa, CA, Managing Partner, ProfitPulse",
          rx + Inches(0.25), Inches(1.28), rw - Inches(0.5), Inches(0.85),
          font_name=SERIF, font_size=16, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT, line_spacing=1.1)

cred_lines = [
    "16 years of experience across 4 countries",
    "Over 52 deals executed and managed",
    "Queensland GRBT project value over AUD 10 billion",
    "CA qualification, SAICA South Africa",
]
cred_fmt = [{"text": "•  " + t, "size": 12.5, "colour": OFF_WHITE} for t in cred_lines]
add_multiline(s3, cred_fmt, rx + Inches(0.25), Inches(2.28), rw - Inches(0.5), Inches(1.7),
              spacing_after=8, line_spacing=1.15)

add_rect(s3, rx + Inches(0.25), Inches(4.15), rw - Inches(0.5), Pt(1.5), fill_colour=TEAL)

contact_fmt = [
    {"text": "Profit-Pulse.com.au", "size": 13, "colour": OFF_WHITE, "bold": False},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 13, "colour": TEAL, "bold": False},
    {"text": "+61 411 876 267", "size": 13, "colour": OFF_WHITE, "bold": False},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 11, "colour": OFF_WHITE, "bold": False},
    {"text": "Brisbane, Australia", "size": 11, "colour": OFF_WHITE, "bold": False},
]
add_multiline(s3, contact_fmt, rx + Inches(0.25), Inches(4.35), rw - Inches(0.5), Inches(2.1),
              spacing_after=8, line_spacing=1.15)

footer(s3)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Attekus_29Aug2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
