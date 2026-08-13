"""
ProfitPulse Brief Builder v3.3 house style
Target: iCatalyst | Date: 14 Aug 2026
Three slide prospect facing deck. White body background, black header band,
amber left accent stripe, teal topped stat cards. Brand colours only. Zero dashes
except required literal brand domain strings.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Brand colours
BLACK      = RGBColor(0x00, 0x00, 0x00)
TEAL       = RGBColor(0x01, 0xA2, 0x96)
AMBER_B    = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D    = RGBColor(0xF6, 0xA1, 0x02)
GOLD       = RGBColor(0xE3, 0xA7, 0x12)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xE6, 0xE5, 0xDE)

SERIF = "Liberation Serif"
SANS  = "Liberation Sans"

W = Inches(13.333)
H = Inches(7.5)
DATE_STR = "14 Aug 2026"
COMPANY = "iCatalyst"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


def set_bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, x, y, w, h, colour, line_colour=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = colour
    if line_colour:
        shp.line.color.rgb = line_colour
        shp.line.width = line_w or Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    # Strip the default <p:style> effectRef so no theme shadow is applied.
    sp = shp._element
    style_el = sp.find("{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style_el is not None:
        sp.remove(style_el)
    return shp


def text(slide, s, x, y, w, h, size, colour, bold=False, italic=False,
         font=SANS, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=None,
         wrap=True, autofit_shrink=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if autofit_shrink:
        from pptx.enum.text import MSO_AUTO_SIZE
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
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


def multiline(slide, runs, x, y, w, h, font=SANS, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, space_after=4):
    """runs: list of dicts {text,size,colour,bold,italic,font}"""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, cfg in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = cfg["text"]
        r.font.name = cfg.get("font", font)
        r.font.size = Pt(cfg.get("size", 12))
        r.font.bold = cfg.get("bold", False)
        r.font.italic = cfg.get("italic", False)
        r.font.color.rgb = cfg.get("colour", BLACK)
    return box


def add_hyperlink_text(slide, s, x, y, w, h, size, colour, url, bold=True,
                        font=SANS, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = colour
    r.hyperlink.address = url
    return box


def chrome(slide, eyebrow, right_label=BLACK):
    """Fixed chrome: left accent stripe, black header band, footer line."""
    set_bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.4), Inches(0.36), Inches(7.5), Inches(0.35),
         12, OFF_WHITE, bold=True, font=SANS, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, COMPANY, Inches(9.0), Inches(0.32), Inches(3.93), Inches(0.4),
         14, TEAL, bold=True, font=SERIF, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # Footer line
    rect(slide, Inches(0.4), Inches(7.02), Inches(12.5), Pt(0.75), TEAL)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.4), Inches(7.08), Inches(9.5), Inches(0.3), 8, BLACK)
    text(slide, DATE_STR, Inches(10.4), Inches(7.08), Inches(2.5), Inches(0.3),
         8, BLACK, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF")

text(s1, "iCatalyst", Inches(0.4), Inches(1.12), Inches(9), Inches(0.72),
     40, BLACK, bold=True, font=SERIF)
text(s1, "Microsoft Dynamics 365 and Power Platform consultancy, Subiaco, Perth WA",
     Inches(0.4), Inches(1.86), Inches(11.5), Inches(0.32), 13, BLACK, font=SANS)

# Stat cards row: 5 cards, 2.3in wide, 0.15in gap, top y=2.4, height 1.6
card_y = Inches(2.35)
card_w = Inches(2.3)
card_h = Inches(1.55)
gap = Inches(0.15)
start_x = Inches(0.4)
cards = [
    ("$13.5M", "FY2025 verified\nrevenue", "SmartCompany Smart50 2025"),
    ("60%",    "Revenue growth\nyear on year", "SmartCompany Smart50 2025"),
    ("50",     "Team members\nnationally", "SmartCompany Smart50 2025"),
    ("4",      "Offices across\nAustralia", "ITBrief, iTWire 2025"),
    ("2020",   "Year founded\nin Perth", "SmartCompany Smart50 2023"),
]
for i, (num, label, src) in enumerate(cards):
    cx = start_x + i * (card_w + gap)
    rect(s1, cx, card_y, card_w, card_h, BLACK)
    rect(s1, cx, card_y, card_w, Pt(4), TEAL)
    text(s1, num, cx + Inches(0.15), card_y + Inches(0.14), card_w - Inches(0.3), Inches(0.5),
         28, AMBER_B, bold=True, font=SERIF)
    text(s1, label, cx + Inches(0.15), card_y + Inches(0.68), card_w - Inches(0.3), Inches(0.5),
         11, OFF_WHITE, font=SANS, line_spacing=1.0)
    text(s1, src, cx + Inches(0.15), card_y + Inches(1.28), card_w - Inches(0.3), Inches(0.24),
         7.5, OFF_WHITE, italic=True, font=SANS)

# Section: chart (left) + Key Commercial Signals (right)
sec_y = Inches(4.18)
text(s1, "VERIFIED REVENUE GROWTH", Inches(0.4), sec_y, Inches(4.4), Inches(0.28),
     11, TEAL, bold=True, font=SANS)
text(s1, "KEY COMMERCIAL SIGNALS", Inches(5.15), sec_y, Inches(7.6), Inches(0.28),
     11, TEAL, bold=True, font=SANS)

# Bar chart, manual shapes, 3 verified data points
baseline_y = 6.35  # inches, screen coords top based
max_bar_h = 1.45   # inches
data = [("2023", 5.22), ("2024", 9.5), ("2025", 13.5)]
bar_w = Inches(0.85)
bar_gap = Inches(0.35)
bar_colours = [TEAL, AMBER_D, GOLD]
chart_x0 = Inches(0.75)
for i, (yr, val) in enumerate(data):
    bh = max_bar_h * (val / 13.5)
    bx = chart_x0 + i * (bar_w + bar_gap)
    by = Inches(baseline_y - bh)
    rect(s1, bx, by, bar_w, Inches(bh), bar_colours[i])
    text(s1, f"${val}M", bx - Inches(0.15), Inches(baseline_y - bh - 0.3),
         Inches(1.15), Inches(0.26), 10, BLACK, bold=True, align=PP_ALIGN.CENTER)
    text(s1, yr, bx - Inches(0.15), Inches(baseline_y + 0.06), Inches(1.15), Inches(0.22),
         10, BLACK, align=PP_ALIGN.CENTER)
# baseline
rect(s1, chart_x0 - Inches(0.05), Inches(baseline_y), Inches(3.5), Pt(1), TEAL)
text(s1, "Source: SmartCompany Smart50 award citations, 2023 to 2025",
     Inches(0.4), Inches(6.72), Inches(4.4), Inches(0.22), 7.5, BLACK,
     italic=True, line_spacing=1.0)

signals = [
    "Acquired Brisbane based Beyond CRM in 2025, adding a fourth national office (ITBrief, iTWire)",
    "Average project value grew from $100,000 to $350,000 as delivery scaled up (Microsoft Customer Story)",
    "Achieved Microsoft Solutions Partner status within 9 months of launch (Microsoft Customer Story)",
    "FY26 target: 30% growth in managed services revenue, 25% growth in customers (ARN, techpartner.news)",
    "Three straight Smart50 appearances: rank 40 in 2023, rank 4 in 2024, rank 16 in 2025 (SmartCompany)",
]
runs = []
for sgl in signals:
    runs.append({"text": "\u25a0  " + sgl, "size": 11.5, "colour": BLACK})
multiline(s1, runs, Inches(5.15), Inches(4.55), Inches(7.6), Inches(2.05), space_after=10)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY, THREE COMMERCIAL OBSERVATIONS
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY")
text(s2, "iCatalyst: three commercial observations from ProfitPulse",
     Inches(0.4), Inches(1.12), Inches(12.4), Inches(0.4), 16, BLACK, bold=True, font=SERIF)

col_top = Inches(1.68)
col_h = Inches(4.68)
col_w = Inches(4.24)
col_x = [Inches(0.1), Inches(4.44), Inches(8.78)]
col_fill = [TEAL, BLACK, GOLD]
col_text = [BLACK, WHITE, BLACK]
col_index_colour = [WHITE, AMBER_B, WHITE]

observations = [
    ("01", "Four offices, one growth plan?",
     "iCatalyst added its fourth office this year through the Beyond CRM "
     "acquisition in Brisbane, joining Perth, Melbourne and Sydney. Each "
     "site carries its own hiring, marketing and integration cost. A "
     "Capital Allocation Review would rank where the next dollar across "
     "these four sites should go before FY26 spend is committed."),
    ("02", "Average deal size just tripled",
     "Average project value grew from $100,000 to $350,000 as headcount "
     "scaled to 50, with seventy percent of revenue still tied to time "
     "based project work. A Workforce Capacity and Utilisation Review "
     "would show which offices carry spare delivery capacity and which "
     "are stretched thin."),
    ("03", "FY26 targets need a costed pathway",
     "iCatalyst has set FY26 targets of thirty percent growth in managed "
     "services revenue and twenty five percent growth in customers, plus "
     "new proprietary IP. A Strategic Growth Diagnostic would turn those "
     "public targets into a costed, scenario tested twelve month "
     "roadmap before the year begins."),
]

for i, (idx, head, para) in enumerate(observations):
    rect(s2, col_x[i], col_top, col_w, col_h, col_fill[i])
    text(s2, idx, col_x[i] + Inches(0.28), col_top + Inches(0.22), Inches(1.6), Inches(0.7),
         40, col_index_colour[i], bold=True, font=SERIF)
    text(s2, head, col_x[i] + Inches(0.28), col_top + Inches(0.98), col_w - Inches(0.56), Inches(0.75),
         15, col_text[i], bold=True, font=SANS, line_spacing=1.05)
    text(s2, para, col_x[i] + Inches(0.28), col_top + Inches(1.8), col_w - Inches(0.56), Inches(2.7),
         11, col_text[i], font=SANS, line_spacing=1.18)

text(s2, "These observations are offered in good faith. iCatalyst has built something "
         "rare in five years. The question is simply whether the capital plan matches the ambition.",
     Inches(0.4), Inches(6.46), Inches(12.5), Inches(0.5), 10.5, BLACK,
     italic=True, font=SANS)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION")

lx = Inches(0.4)
lw = Inches(7.6)

text(s3, "Capital Allocation Review", lx, Inches(1.15), lw, Inches(0.55),
     26, BLACK, bold=True, font=SERIF)
text(s3, "$7,500 one off", lx, Inches(1.72), lw, Inches(0.36),
     18, AMBER_D, bold=True, font=SANS)
text(s3, "Independent review of where capital is deployed across people, offices and "
         "the Beyond CRM integration, ranked against the return each is generating.",
     lx, Inches(2.14), lw, Inches(0.62), 12, BLACK, font=SANS, line_spacing=1.15)

# Step one block
rect(s3, lx, Inches(2.9), lw, Inches(1.55), OFF_WHITE, line_colour=TEAL, line_w=Pt(1))
text(s3, "Step one: answer a few quick questions", lx + Inches(0.22), Inches(3.06), lw - Inches(0.44),
     Inches(0.34), 13, BLACK, bold=True, font=SANS)
text(s3, "See the solutions matched to your size and industry.",
     lx + Inches(0.22), Inches(3.42), lw - Inches(0.44), Inches(0.3), 11, BLACK, font=SANS)
add_hyperlink_text(s3, "profit-pulse.com.au/services/find-your-fit",
                    lx + Inches(0.22), Inches(3.76), lw - Inches(0.44), Inches(0.32),
                    13, TEAL, "https://profit-pulse.com.au/services/find-your-fit/")
add_hyperlink_text(s3, "Purchase the suggested product now to get started",
                    lx + Inches(0.22), Inches(4.14), lw - Inches(0.44), Inches(0.28),
                    11, AMBER_D, "https://buy.stripe.com/5kQ9AU7OU9io95m8sn3ks0l")

text(s3, "Prefer a conversation first?", lx, Inches(4.66), lw, Inches(0.3),
     11.5, BLACK, bold=True, font=SANS)
add_hyperlink_text(s3, "Book a complimentary discovery call",
                    lx, Inches(4.98), lw, Inches(0.3), 11.5, TEAL,
                    "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

text(s3, "Supporting services identified: Strategic Growth Diagnostic, Workforce Capacity and Utilisation Review",
     lx, Inches(5.5), lw, Inches(0.4), 9, BLACK, italic=True, font=SANS)

# Right column: About ProfitPulse and Nitesh
rx = Inches(8.28)
rw = Inches(4.62)
rect(s3, rx, Inches(1.15), rw, Inches(5.55), BLACK)
rect(s3, rx, Inches(1.15), rw, Pt(4), TEAL)

text(s3, "Nitesh Roopa", rx + Inches(0.25), Inches(1.4), rw - Inches(0.5), Inches(0.42),
     18, AMBER_B, bold=True, font=SERIF)
text(s3, "CA, Managing Partner, ProfitPulse", rx + Inches(0.25), Inches(1.86), rw - Inches(0.5), Inches(0.3),
     12, OFF_WHITE, font=SANS)
rect(s3, rx + Inches(0.25), Inches(2.24), rw - Inches(0.5), Pt(1.2), TEAL)

cred_runs = [
    {"text": "16 years of experience across 4 countries", "size": 11, "colour": OFF_WHITE},
    {"text": "52 deals executed and managed", "size": 11, "colour": OFF_WHITE},
    {"text": "Largest deal USD 1.3 billion, Cahora Bassa", "size": 11, "colour": OFF_WHITE},
    {"text": "AUD 10 billion Gympie Road Bypass Tunnel business case, QIC", "size": 11, "colour": OFF_WHITE},
]
multiline(s3, cred_runs, rx + Inches(0.25), Inches(2.42), rw - Inches(0.5), Inches(1.7), space_after=8)

rect(s3, rx + Inches(0.25), Inches(4.2), rw - Inches(0.5), Pt(1), TEAL)
contact_runs = [
    {"text": "Profit-Pulse.com.au", "size": 11.5, "colour": OFF_WHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 11.5, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 11.5, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 9.5, "colour": OFF_WHITE},
]
multiline(s3, contact_runs, rx + Inches(0.25), Inches(4.38), rw - Inches(0.5), Inches(1.5), space_after=7)

OUT_PATH = "/home/user/Daily-Outreach/Out-reach efforts/Brief_iCatalyst_14Aug2026.pptx"
prs.save(OUT_PATH)

# ── Post process: force theme hyperlink colours to brand teal ──────────────
# LibreOffice renders hyperlink run text using the theme's hlink/folHlink
# colour regardless of explicit run colour, so retint the theme itself.
import zipfile, shutil, re

tmp_path = OUT_PATH + ".tmp"
shutil.copy(OUT_PATH, tmp_path)
with zipfile.ZipFile(tmp_path, "r") as zin:
    names = zin.namelist()
    theme_names = [n for n in names if re.match(r"ppt/theme/theme\d+\.xml", n)]
    contents = {n: zin.read(n) for n in names}

for tn in theme_names:
    xml = contents[tn].decode("utf-8")
    xml = re.sub(r"(<a:hlink>\s*<a:srgbClr val=\")[0-9A-Fa-f]{6}(\"\s*/>\s*</a:hlink>)",
                 r"\g<1>01A296\g<2>", xml)
    xml = re.sub(r"(<a:folHlink>\s*<a:srgbClr val=\")[0-9A-Fa-f]{6}(\"\s*/>\s*</a:folHlink>)",
                 r"\g<1>01A296\g<2>", xml)
    contents[tn] = xml.encode("utf-8")

with zipfile.ZipFile(OUT_PATH, "w", zipfile.ZIP_DEFLATED) as zout:
    for n in names:
        zout.writestr(n, contents[n])
import os
os.remove(tmp_path)
print("Saved PPTX with retinted theme hyperlink colours")
