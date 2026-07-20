"""
ProfitPulse Brief Builder
Target: Equality Media + Marketing | Date: 21 Jul 2026
Three slide prospect facing deck. House style per Section 6. Brand colours only. Zero dashes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE  = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

DATE_STAMP = "21 Jul 2026"
COMPANY = "Equality Media + Marketing"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]


def bg(slide, colour):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = colour


def rect(slide, x, y, w, h, colour, line=None):
    shp = slide.shapes.add_shape(1, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = colour
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, x, y, w, h, size=12, bold=False, colour=BLACK,
         align=PP_ALIGN.LEFT, font="Calibri", anchor=None, italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    r.font.color.rgb = colour
    return box


def multiline(slide, lines, x, y, w, h, align=PP_ALIGN.LEFT, spacing=4, font="Calibri"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = ln["text"]
        r.font.size = Pt(ln.get("size", 12))
        r.font.bold = ln.get("bold", False)
        r.font.italic = ln.get("italic", False)
        r.font.name = ln.get("font", font)
        r.font.color.rgb = ln.get("colour", BLACK)
    return box


def chrome(slide, eyebrow, header_right, dark_footer=False):
    """Fixed chrome per Section 6.0A: left stripe, header band, footer line."""
    bg(slide, WHITE)
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)
    rect(slide, Inches(0), Inches(0), W, Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.28), Inches(7), Inches(0.45),
         size=12, bold=True, colour=OFFWHITE, font="Cambria")
    text(slide, header_right, Inches(6.5), Inches(0.28), Inches(6.5), Inches(0.45),
         size=12, bold=True, colour=TEAL, align=PP_ALIGN.RIGHT, font="Calibri")
    # Footer line
    rect(slide, Inches(0.35), Inches(7.05), Inches(12.6), Pt(0.75), OFFWHITE)
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.35), Inches(7.1), Inches(9.5), Inches(0.3), size=8, colour=BLACK)
    text(slide, DATE_STAMP, Inches(10.5), Inches(7.1), Inches(2.45), Inches(0.3),
         size=8, colour=BLACK, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1: THE COMMERCIAL INTELLIGENCE BRIEF
# ══════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

text(s1, COMPANY, Inches(0.35), Inches(1.2), Inches(11), Inches(0.75),
     size=40, bold=True, colour=AMBER_D, font="Cambria")
text(s1, "Property marketing, media and creative agency, Richmond, Melbourne VIC.",
     Inches(0.35), Inches(1.9), Inches(11.5), Inches(0.4), size=13, colour=BLACK)

# Stat cards, 5 cards, y=2.4, h=1.6
stat_cards = [
    ("$13M", "Revenue, FY2025 profile", "SmartCompany Smart50 2025"),
    ("49%",  "Three year growth, Smart50", "SmartCompany Smart50 2025"),
    ("#24",  "Smart50 2025 rank, up from 45th", "SmartCompany Smart50 2025"),
    ("30",   "People at the agency now", "Mediaweek, AFR BOSS 2026"),
    ("2018", "Founded, Richmond VIC", "Company website, Smart50"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
x0 = Inches(0.35)
y0 = Inches(2.4)
card_h = Inches(1.6)
for i, (num, label, src) in enumerate(stat_cards):
    cx = x0 + i * (card_w + gap)
    rect(s1, cx, y0, card_w, card_h, BLACK)
    rect(s1, cx, y0, card_w, Pt(4), TEAL)
    text(s1, num, cx + Inches(0.12), y0 + Inches(0.15), card_w - Inches(0.24), Inches(0.5),
         size=28, bold=True, colour=AMBER_B, font="Cambria")
    text(s1, label, cx + Inches(0.12), y0 + Inches(0.68), card_w - Inches(0.24), Inches(0.5),
         size=11, colour=OFFWHITE)
    text(s1, src, cx + Inches(0.12), y0 + Inches(1.28), card_w - Inches(0.24), Inches(0.28),
         size=7, italic=True, colour=OFFWHITE)

# Chart: revenue by Smart50 profile year (left half)
chart_y = Inches(4.25)
chart_h = Inches(1.75)
chart_w = Inches(5.6)
cd = CategoryChartData()
cd.categories = ["2024 profile", "2025 profile"]
cd.add_series("Revenue $M", (7.3, 13.0))
gframe = s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x0, chart_y, chart_w, chart_h, cd)
chart = gframe.chart
chart.has_legend = False
chart.has_title = True
chart.chart_title.text_frame.text = "Revenue, $ millions, Smart50 profile years"
chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
chart.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = BLACK
plot = chart.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '"$"0.0"M"'
plot.data_labels.number_format_is_linked = False
plot.data_labels.font.size = Pt(10)
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = TEAL
cat_ax = chart.category_axis
cat_ax.tick_labels.font.size = Pt(9)
val_ax = chart.value_axis
val_ax.visible = False
val_ax.has_major_gridlines = False

text(s1, "Source: SmartCompany Smart50 2024 and 2025 profiles",
     x0, chart_y + chart_h + Inches(0.02), chart_w, Inches(0.25),
     size=7, italic=True, colour=BLACK)

# Key Commercial Signals (right half)
sig_x = Inches(6.25)
sig_w = Inches(6.7)
text(s1, "KEY COMMERCIAL SIGNALS", sig_x, chart_y - Inches(0.02), sig_w, Inches(0.3),
     size=11, bold=True, colour=TEAL, font="Cambria")
signals = [
    "Revenue $7.3M to $13M in a year, headcount 15 to 30 in the same window. (Smart50 24/25)",
    "Founded 2018 as a media buyer, now also offers strategy and creative in house. (Smart50)",
    "About half of all clients already buy more than one service line. (Smart50 2025)",
    "AFR BOSS Best Places to Work winner, third time in four years, 2026. (Mediaweek)",
    "AdNews Agency of the Year and a Wellbeing Initiative award winner. (AdNews)",
]
multiline(
    s1,
    [{"text": "• " + t, "size": 10.5, "colour": BLACK} for t in signals],
    sig_x, chart_y + Inches(0.3), sig_w, Inches(1.9), spacing=6,
)

# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
chrome(s2, "THE OPPORTUNITY", COMPANY)

text(s2, "Three commercial observations from ProfitPulse",
     Inches(0.35), Inches(1.15), Inches(11), Inches(0.4), size=14, bold=True,
     colour=BLACK, font="Cambria")

col_y = Inches(1.7)
col_h = Inches(4.35)
col_w = Inches(4.14)
gap2 = Inches(0.1)
cols = [
    dict(fill=TEAL, text_colour=BLACK, idx="01",
         header="Growth has outpaced margin visibility",
         body=("Revenue grew 78 percent in a year, $7.3 million to $13 million, while the "
               "team roughly doubled, 15 to 30 people. (Smart50 2024, 2025 profiles). When "
               "headcount grows faster than revenue, margin usually moves before it shows on "
               "a P&L. A Product and Service Line Profitability review gives a ranked answer "
               "inside three weeks.")),
    dict(fill=BLACK, text_colour=WHITE, idx="02",
         header="Three service lines, one blended number",
         body=("Equality Media started in 2018 as a media buying practice and has since "
               "formally added strategy and creative, with about half of clients now buying "
               "more than one line. (Smart50 2025). Media, creative, and strategy carry "
               "different margins and staffing needs, so growth decisions are currently being "
               "made on revenue alone.")),
    dict(fill=GOLD, text_colour=BLACK, idx="03",
         header="One sector carries the whole growth story",
         body=("The agency is built specifically around the property sector, the foundation "
               "of its Smart50 and AFR BOSS results. (Company website, Mediaweek). That focus "
               "is a real strength, but it also means one property cycle carries more weight "
               "than it would for a diversified agency. A Strategic Growth Diagnostic can "
               "quantify how much of next year's plan should lean on adjacent sectors.")),
]
for i, c in enumerate(cols):
    cx = Inches(0.35) + i * (col_w + gap2)
    rect(s2, cx, col_y, col_w, col_h, c["fill"])
    text(s2, c["idx"], cx + Inches(0.25), col_y + Inches(0.2), col_w - Inches(0.5), Inches(0.7),
         size=30, bold=True, colour=c["text_colour"], font="Cambria")
    text(s2, c["header"], cx + Inches(0.25), col_y + Inches(0.95), col_w - Inches(0.5), Inches(0.9),
         size=14, bold=True, colour=c["text_colour"], font="Cambria")
    text(s2, c["body"], cx + Inches(0.25), col_y + Inches(1.9), col_w - Inches(0.5), Inches(2.3),
         size=10.5, colour=c["text_colour"])

text(s2, ("These are observations offered in good faith. Equality Media has built something "
          "impressive in seven years. The question is simply whether the financial "
          "architecture matches the pace of the last twelve months."),
     Inches(0.35), Inches(6.2), Inches(12.6), Inches(0.75), size=10.5, italic=True,
     colour=BLACK)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ══════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
chrome(s3, "THE RECOMMENDATION", COMPANY)

left_x = Inches(0.35)
left_w = Inches(7.6)

text(s3, "Product and Service Line Profitability", left_x, Inches(1.2), left_w, Inches(0.6),
     size=22, bold=True, colour=AMBER_D, font="Cambria")
text(s3, "$2,950 one off. ProfitPulse verified figure.", left_x, Inches(1.8), left_w, Inches(0.4),
     size=15, bold=True, colour=TEAL)
text(s3, ("Ranks media, creative, and strategy by gross margin, contribution margin, and "
          "operational drag, in three weeks flat."),
     left_x, Inches(2.25), left_w, Inches(0.6), size=11.5, colour=BLACK)

# Step one block
rect(s3, left_x, Inches(2.95), left_w, Inches(1.55), OFFWHITE)
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.25), Inches(3.1),
     left_w - Inches(0.5), Inches(0.4), size=13, bold=True, colour=TEAL, font="Cambria")
text(s3, "See the solutions matched to your size and industry.",
     left_x + Inches(0.25), Inches(3.55), left_w - Inches(0.5), Inches(0.4), size=11, colour=BLACK)
link_box = text(s3, "profit-pulse.com.au/services/find-your-fit",
     left_x + Inches(0.25), Inches(3.95), left_w - Inches(0.5), Inches(0.4), size=12, bold=True, colour=TEAL)
run = link_box.text_frame.paragraphs[0].runs[0]
run.hyperlink.address = "https://profit-pulse.com.au/services/find-your-fit/"
run.font.underline = False
run.font.color.rgb = TEAL

# Direct CTA
cta_box = text(s3, "Purchase the suggested product now to get started",
     left_x, Inches(4.75), left_w, Inches(0.5), size=14, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)
cta_rect = rect(s3, left_x, Inches(4.65), left_w, Inches(0.65), AMBER_D)
# move CTA text above rect z order by re-adding after rect
s3.shapes._spTree.remove(cta_box._element)
s3.shapes._spTree.append(cta_box._element)
run2 = cta_box.text_frame.paragraphs[0].runs[0]
run2.hyperlink.address = "https://buy.stripe.com/eVq7sM6KQ1PW95m7oj3ks1C"
run2.font.underline = False
run2.font.color.rgb = WHITE
cta_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

text(s3, "Prefer a conversation first?", left_x, Inches(5.55), left_w, Inches(0.35),
     size=11, colour=BLACK)
call_box = text(s3, "Book a complimentary discovery call", left_x, Inches(5.9), left_w, Inches(0.4),
     size=12, bold=True, colour=TEAL)
run3 = call_box.text_frame.paragraphs[0].runs[0]
run3.hyperlink.address = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"
run3.font.underline = False
run3.font.color.rgb = TEAL

# Right column: credibility panel
right_x = Inches(8.25)
right_w = Inches(4.7)
rect(s3, right_x, Inches(1.2), right_w, Inches(5.55), BLACK)
text(s3, "Nitesh Roopa", right_x + Inches(0.3), Inches(1.4), right_w - Inches(0.6), Inches(0.45),
     size=18, bold=True, colour=AMBER_B, font="Cambria")
text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.3), Inches(1.85),
     right_w - Inches(0.6), Inches(0.35), size=12, colour=WHITE)

cred_lines = [
    {"text": "16 years across 4 countries", "size": 11, "colour": OFFWHITE},
    {"text": "52 deals executed and managed", "size": 11, "colour": OFFWHITE},
    {"text": "Largest single deal, USD 1.3 billion, Cahora Bassa", "size": 11, "colour": OFFWHITE},
    {"text": "Over AUD 10 billion in Queensland GRBT project value", "size": 11, "colour": OFFWHITE},
]
multiline(s3, cred_lines, right_x + Inches(0.3), Inches(2.4), right_w - Inches(0.6), Inches(1.7), spacing=8)

rect(s3, right_x + Inches(0.3), Inches(4.15), right_w - Inches(0.6), Pt(1), TEAL)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 11.5, "colour": OFFWHITE},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 11.5, "colour": TEAL},
    {"text": "+61 411 876 267", "size": 11.5, "colour": OFFWHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10.5, "colour": OFFWHITE},
]
multiline(s3, contact_lines, right_x + Inches(0.3), Inches(4.35), right_w - Inches(0.6), Inches(1.8), spacing=8)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_EqualityMediaMarketing_21Jul2026.pptx"
prs.save(out_path)

# Patch the theme's hyperlink colours so LibreOffice/PowerPoint render hyperlinked
# text in brand colours instead of the default theme blue/purple (Rule 3).
import zipfile
import shutil

tmp_path = out_path + ".tmp"
with zipfile.ZipFile(out_path, "r") as zin:
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/theme/theme1.xml":
                data = data.replace(
                    b'<a:hlink><a:srgbClr val="0000FF"/></a:hlink>',
                    b'<a:hlink><a:srgbClr val="01A296"/></a:hlink>',
                )
                data = data.replace(
                    b'<a:folHlink><a:srgbClr val="800080"/></a:folHlink>',
                    b'<a:folHlink><a:srgbClr val="F6A102"/></a:folHlink>',
                )
            zout.writestr(item, data)
shutil.move(tmp_path, out_path)
print(f"PPTX saved: {out_path}")
