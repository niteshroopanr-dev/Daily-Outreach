"""
ProfitPulse Brief Builder
Target: AdUnion | Date: 09 Jun 2026
Three-slide prospect-facing deck. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
import copy

# ── Brand colours ─────────────────────────────────────────────────────────────
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)
DARK_PANEL = RGBColor(0x11, 0x11, 0x11)
TEAL_DARK  = RGBColor(0x04, 0x1A, 0x18)

# Slide dimensions: widescreen 13.333" x 7.5"
W = Inches(13.333)
H = Inches(7.5)


def set_background(slide, colour):
    """Fill the slide background with a solid colour."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None, line_width=None):
    """Add a filled rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()  # no line
    return shape


def add_text(slide, text, left, top, width, height,
             font_name="Arial", font_size=18, bold=False,
             colour=WHITE, align=PP_ALIGN.LEFT,
             wrap=True, italic=False):
    """Add a text box."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return txBox


def add_multiline(slide, lines, left, top, width, height,
                  font_name="Arial", default_size=14,
                  default_colour=OFF_WHITE, default_bold=False,
                  align=PP_ALIGN.LEFT, spacing_after=None):
    """
    Add a text box with multiple lines. Each element of `lines` is either:
      - a string (uses defaults), or
      - a dict with keys: text, size, colour, bold, italic
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    first = True
    for line in lines:
        if isinstance(line, str):
            cfg = {"text": line, "size": default_size,
                   "colour": default_colour, "bold": default_bold,
                   "italic": False}
        else:
            cfg = {
                "text":   line.get("text", ""),
                "size":   line.get("size", default_size),
                "colour": line.get("colour", default_colour),
                "bold":   line.get("bold", default_bold),
                "italic": line.get("italic", False),
            }

        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()

        p.alignment = align
        if spacing_after:
            p.space_after = Pt(spacing_after)
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = font_name
        run.font.size = Pt(cfg["size"])
        run.font.bold = cfg["bold"]
        run.font.italic = cfg["italic"]
        run.font.color.rgb = cfg["colour"]

    return txBox


# ══════════════════════════════════════════════════════════════════════════════
# BUILD PRESENTATION
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank_layout = prs.slide_layouts[6]  # completely blank


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 1: THE OPENING
# ────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank_layout)
set_background(s1, BLACK)

# Teal left accent bar
add_rect(s1, Inches(0), Inches(0), Inches(0.08), H, TEAL)

# ProfitPulse brand mark, top left
add_text(s1, "PROFITPULSE",
         left=Inches(0.25), top=Inches(0.32), width=Inches(4), height=Inches(0.4),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

# Teal divider line under brand mark
add_rect(s1, Inches(0.25), Inches(0.76), Inches(3.0), Pt(2), TEAL)

# Company name: AdUnion, large amber heading
add_text(s1, "AdUnion",
         left=Inches(0.25), top=Inches(1.15), width=Inches(9), height=Inches(1.2),
         font_size=68, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

# Opportunity line 1
add_text(s1, "Three years of compounding growth.",
         left=Inches(0.25), top=Inches(2.45), width=Inches(9.5), height=Inches(0.6),
         font_size=26, bold=False, colour=WHITE, align=PP_ALIGN.LEFT)

# Opportunity line 2
add_text(s1, "The next stage belongs to margin clarity.",
         left=Inches(0.25), top=Inches(3.05), width=Inches(9.5), height=Inches(0.6),
         font_size=26, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

# Teal horizontal accent line
add_rect(s1, Inches(0.25), Inches(3.72), Inches(5.0), Pt(2), TEAL)

# Verified metrics row
add_multiline(
    s1,
    [
        {"text": "$9.3M revenue   •   81% three year growth   •   10 people",
         "size": 15, "colour": OFF_WHITE, "bold": False},
        {"text": "Smart50 2025 #8   •   AFR Fast 100 2025 #26   •   Deloitte Tech Fast 50 2025 #31",
         "size": 13, "colour": TEAL, "bold": False},
    ],
    left=Inches(0.25), top=Inches(3.85), width=Inches(12.5), height=Inches(0.9),
    default_size=14, default_colour=OFF_WHITE,
)

# Right side: large teal accent word
add_text(s1, "MELBOURNE",
         left=Inches(9.5), top=Inches(1.5), width=Inches(3.5), height=Inches(0.5),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)
add_text(s1, "2016",
         left=Inches(9.5), top=Inches(2.0), width=Inches(3.5), height=Inches(0.8),
         font_size=48, bold=True, colour=RGBColor(0x22, 0x22, 0x22), align=PP_ALIGN.RIGHT)
add_text(s1, "Founded",
         left=Inches(9.5), top=Inches(2.8), width=Inches(3.5), height=Inches(0.4),
         font_size=11, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.RIGHT)

# Date, bottom right
add_text(s1, "09 Jun 2026",
         left=Inches(10), top=Inches(6.9), width=Inches(3.0), height=Inches(0.4),
         font_size=12, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.RIGHT)

# ProfitPulse tag, bottom left
add_text(s1, "Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(6.9), width=Inches(4), height=Inches(0.4),
         font_size=12, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 2: THE INSIGHT
# ────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank_layout)
set_background(s2, BLACK)

# Teal left accent bar
add_rect(s2, Inches(0), Inches(0), Inches(0.08), H, TEAL)

# Header bar
add_rect(s2, Inches(0.08), Inches(0), Inches(W - Inches(0.08)), Inches(0.95), DARK_PANEL)

# Slide title
add_text(s2, "THE INSIGHT",
         left=Inches(0.25), top=Inches(0.15), width=Inches(7), height=Inches(0.6),
         font_size=13, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

# Company name in header
add_text(s2, "AdUnion",
         left=Inches(9), top=Inches(0.15), width=Inches(4), height=Inches(0.6),
         font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# ── LEFT COLUMN: Verified data ──────────────────────────────────────────────
left_panel_x = Inches(0.25)
left_panel_w = Inches(4.5)

add_text(s2, "Verified facts",
         left=left_panel_x, top=Inches(1.05), width=left_panel_w, height=Inches(0.35),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

# Data items
data_items = [
    ("Revenue (FY2024)",        "$9.3 million"),
    ("3 year growth",           "81% average"),
    ("Revenue since 2023",      "Tripled"),
    ("Team size",               "10 people"),
    ("Smart50 2025",            "Rank 8 of 50"),
    ("AFR Fast 100 2025",       "Rank 26"),
    ("Deloitte Tech Fast 50",   "Rank 31"),
    ("Founded",                 "2016"),
    ("Location",                "Cremorne, Melbourne VIC"),
    ("Managing Director",       "Robert Ong"),
]

y = Inches(1.45)
for label, val in data_items:
    # Label
    add_text(s2, label,
             left=left_panel_x, top=y, width=Inches(2.1), height=Inches(0.38),
             font_size=11, bold=False, colour=RGBColor(0x88, 0x88, 0x88), align=PP_ALIGN.LEFT)
    # Value
    add_text(s2, val,
             left=Inches(2.45), top=y, width=Inches(2.3), height=Inches(0.38),
             font_size=11, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    y += Inches(0.38)

# Source note
add_text(s2, "Source: SmartCompany Smart50 2025 award citation and public industry publications",
         left=left_panel_x, top=Inches(5.42), width=left_panel_w, height=Inches(0.5),
         font_size=9, bold=False, colour=RGBColor(0x55, 0x55, 0x55), align=PP_ALIGN.LEFT,
         italic=True)

# ── Vertical divider ────────────────────────────────────────────────────────
add_rect(s2, Inches(4.9), Inches(1.0), Pt(1.5), Inches(5.7), TEAL)

# ── RIGHT COLUMN: Wedge ─────────────────────────────────────────────────────
right_x = Inches(5.1)
right_w = Inches(7.9)

add_text(s2, "The commercial observation",
         left=right_x, top=Inches(1.05), width=right_w, height=Inches(0.35),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

wedge = (
    "AdUnion has compounded at 81 percent over three years and now manages "
    "streaming campaigns for Samsung, Tubi, Tangerine Telecom, and a growing "
    "roster across retail, travel, and financial services. With ten people "
    "generating $9.3 million in revenue, productivity is exceptional.\n\n"
    "The question that follows rapid, diversified growth is: which clients "
    "drive the real margin after the full cost of service, team time, and "
    "platform investment are allocated? In most agencies at this stage, the "
    "answer is uneven. The clients that look largest by revenue are not always "
    "the ones that yield the most to the business.\n\n"
    "The Customer Concentration and Profitability Map answers that question "
    "in three weeks: every client ranked by revenue, gross margin contribution, "
    "and effort to serve. The output is a clear action list: grow these, "
    "reprice these, reset these."
)

add_text(s2, wedge,
         left=right_x, top=Inches(1.45), width=right_w, height=Inches(4.3),
         font_size=13, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT,
         wrap=True)

# ── Bottom: Matched service bar ──────────────────────────────────────────────
add_rect(s2, Inches(0.08), Inches(6.3), W - Inches(0.08), Inches(0.9), TEAL_DARK)
add_rect(s2, Inches(0.08), Inches(6.3), Inches(0.4), Inches(0.9), TEAL)

add_multiline(
    s2,
    [
        {"text": "Matched Service: Customer Concentration and Profitability Map  (G2)",
         "size": 14, "colour": AMBER_B, "bold": True},
        {"text": "Command tier  •  $3,950 one off  •  ProfitPulse verified price  •  Questionnaire confirms exact fit",
         "size": 11, "colour": OFF_WHITE, "bold": False},
    ],
    left=Inches(0.7), top=Inches(6.35), width=Inches(12.3), height=Inches(0.85),
    default_size=12, default_colour=OFF_WHITE,
)

# Footer
add_text(s2, "Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(7.15), width=Inches(4), height=Inches(0.3),
         font_size=10, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)
add_text(s2, "09 Jun 2026",
         left=Inches(10), top=Inches(7.15), width=Inches(3), height=Inches(0.3),
         font_size=10, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 3: THE CALL TO ACTION
# ────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank_layout)
set_background(s3, BLACK)

# Teal left accent bar
add_rect(s3, Inches(0), Inches(0), Inches(0.08), H, TEAL)

# Header bar
add_rect(s3, Inches(0.08), Inches(0), W - Inches(0.08), Inches(0.95), DARK_PANEL)
add_text(s3, "FIND THE RIGHT FIX FOR YOUR BUSINESS",
         left=Inches(0.25), top=Inches(0.15), width=Inches(9), height=Inches(0.6),
         font_size=13, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
add_text(s3, "AdUnion",
         left=Inches(9.5), top=Inches(0.15), width=Inches(3.5), height=Inches(0.6),
         font_size=13, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT)

# ── PRIMARY CTA ─────────────────────────────────────────────────────────────
add_rect(s3, Inches(0.25), Inches(1.1), Inches(7.9), Inches(2.4), TEAL_DARK,
         line_colour=TEAL, line_width=Pt(1))

add_text(s3, "Step 1: Answer a few quick questions",
         left=Inches(0.45), top=Inches(1.18), width=Inches(7.5), height=Inches(0.4),
         font_size=14, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

add_text(s3,
         "See the solutions matched to your size and industry, each with a direct purchase option.",
         left=Inches(0.45), top=Inches(1.6), width=Inches(7.5), height=Inches(0.45),
         font_size=12, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

add_text(s3,
         "profit-pulse.com.au/full-suite-of-products?"
         "utm_source=outreach&utm_medium=pptx&utm_campaign=nightly_outreach&utm_content=adunion",
         left=Inches(0.45), top=Inches(2.08), width=Inches(7.5), height=Inches(0.55),
         font_size=11, bold=False, colour=TEAL, align=PP_ALIGN.LEFT)

add_text(s3, "QUESTIONNAIRE",
         left=Inches(0.45), top=Inches(2.68), width=Inches(2.5), height=Inches(0.5),
         font_size=11, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

# ── SEPARATOR ───────────────────────────────────────────────────────────────
add_rect(s3, Inches(0.25), Inches(3.62), Inches(7.9), Pt(1.5), TEAL)
add_text(s3, "Or start directly",
         left=Inches(0.45), top=Inches(3.69), width=Inches(4), height=Inches(0.35),
         font_size=11, bold=False, colour=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.LEFT)

# ── DIRECT CHECKOUT ─────────────────────────────────────────────────────────
add_rect(s3, Inches(0.25), Inches(4.1), Inches(7.9), Inches(1.85),
         RGBColor(0x1a, 0x15, 0x00),
         line_colour=AMBER_D, line_width=Pt(1))

add_text(s3, "Customer Concentration and Profitability Map",
         left=Inches(0.45), top=Inches(4.18), width=Inches(7.5), height=Inches(0.45),
         font_size=14, bold=True, colour=AMBER_D, align=PP_ALIGN.LEFT)

add_multiline(
    s3,
    [
        {"text": "Command tier  •  $3,950 one off  •  ProfitPulse verified price",
         "size": 12, "colour": OFF_WHITE, "bold": False},
        {"text": "Questionnaire confirms exact tier before purchase if preferred.",
         "size": 11, "colour": RGBColor(0x88, 0x88, 0x88), "bold": False},
    ],
    left=Inches(0.45), top=Inches(4.65), width=Inches(7.5), height=Inches(0.6),
)

add_text(s3,
         "buy.stripe.com/14AbJ21qw2U0ftK0ZV3ks1A",
         left=Inches(0.45), top=Inches(5.27), width=Inches(7.5), height=Inches(0.45),
         font_size=11, bold=False, colour=AMBER_D, align=PP_ALIGN.LEFT)

# ── BOOKING LINK ─────────────────────────────────────────────────────────────
add_text(s3, "Prefer a conversation first?",
         left=Inches(0.45), top=Inches(6.05), width=Inches(7.5), height=Inches(0.35),
         font_size=11, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT)

add_text(s3,
         "Book a complimentary discovery call:  "
         "bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/",
         left=Inches(0.45), top=Inches(6.4), width=Inches(7.5), height=Inches(0.45),
         font_size=10, bold=False, colour=TEAL, align=PP_ALIGN.LEFT)

# ── SIGNATURE BLOCK (right side) ────────────────────────────────────────────
add_rect(s3, Inches(8.45), Inches(1.1), Inches(4.6), Inches(5.7),
         RGBColor(0x0a, 0x0a, 0x0a),
         line_colour=RGBColor(0x22, 0x22, 0x22), line_width=Pt(1))

add_text(s3, "NITESH ROOPA",
         left=Inches(8.65), top=Inches(1.25), width=Inches(4.2), height=Inches(0.5),
         font_size=18, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)

add_text(s3, "CA, Managing Partner",
         left=Inches(8.65), top=Inches(1.75), width=Inches(4.2), height=Inches(0.35),
         font_size=13, bold=False, colour=WHITE, align=PP_ALIGN.LEFT)

add_text(s3, "ProfitPulse",
         left=Inches(8.65), top=Inches(2.1), width=Inches(4.2), height=Inches(0.45),
         font_size=20, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)

add_rect(s3, Inches(8.65), Inches(2.62), Inches(3.8), Pt(1.5), TEAL)

sig_lines = [
    {"text": "Profit-Pulse.com.au",          "size": 12, "colour": OFF_WHITE, "bold": False},
    {"text": "Nitesh@Profit-Pulse.com.au",   "size": 12, "colour": TEAL,      "bold": False},
    {"text": "+61 411 876 267",              "size": 12, "colour": OFF_WHITE, "bold": False},
    {"text": "",                              "size": 8,  "colour": WHITE,     "bold": False},
    {"text": "16 years across 4 countries",  "size": 11, "colour": RGBColor(0x88, 0x88, 0x88), "bold": False},
    {"text": "52 deals executed and managed","size": 11, "colour": RGBColor(0x88, 0x88, 0x88), "bold": False},
    {"text": "CA (SAICA, South Africa)",     "size": 11, "colour": RGBColor(0x88, 0x88, 0x88), "bold": False},
]
add_multiline(
    s3, sig_lines,
    left=Inches(8.65), top=Inches(2.75), width=Inches(4.2), height=Inches(3.5),
    default_size=12, default_colour=OFF_WHITE,
    spacing_after=2,
)

# ── Footer ───────────────────────────────────────────────────────────────────
add_text(s3, "Profit-Pulse.com.au",
         left=Inches(0.25), top=Inches(7.15), width=Inches(4), height=Inches(0.3),
         font_size=10, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.LEFT)
add_text(s3, "09 Jun 2026",
         left=Inches(10), top=Inches(7.15), width=Inches(3), height=Inches(0.3),
         font_size=10, bold=False, colour=RGBColor(0x44, 0x44, 0x44), align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_AdUnion_09Jun2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")
