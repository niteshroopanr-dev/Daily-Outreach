"""
ProfitPulse Brief Builder
Target: Taskforce Australia | Date: 30 Jul 2026
Three slide prospect facing deck. Brand colours only. Zero dashes. No tier names.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

BLACK = RGBColor(0x00, 0x00, 0x00)
TEAL = RGBColor(0x01, 0xA2, 0x96)
AMBER_B = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D = RGBColor(0xF6, 0xA1, 0x02)
GOLD = RGBColor(0xE3, 0xA7, 0x12)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
DATE_STAMP = "30 Jul 2026"
COMPANY = "Taskforce Australia"


def bg(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def rect(slide, left, top, width, height, fill_colour, line_colour=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def text(slide, s, left, top, width, height, size=14, bold=False, colour=BLACK,
         align=PP_ALIGN.LEFT, font="Calibri", italic=False, autosize_safety=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if autosize_safety:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = s
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return box


def multiline(slide, lines, left, top, width, height, align=PP_ALIGN.LEFT,
              font="Calibri", space_after=4):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = line["text"]
        run.font.name = font
        run.font.size = Pt(line.get("size", 12))
        run.font.bold = line.get("bold", False)
        run.font.italic = line.get("italic", False)
        run.font.color.rgb = line.get("colour", BLACK)
    return box


def add_hyperlink(box, url, colour=BLACK):
    run = box.text_frame.paragraphs[0].runs[0]
    run.hyperlink.address = url
    run.font.underline = False
    run.font.color.rgb = colour


def chrome(slide, eyebrow, right_label, right_colour=WHITE):
    """Fixed chrome: left accent stripe, black header band, footer line."""
    rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_B)
    rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    text(slide, eyebrow, Inches(0.35), Inches(0.28), Inches(8.5), Inches(0.45),
         size=12, bold=True, colour=OFF_WHITE, font="Georgia")
    text(slide, right_label, Inches(8.5), Inches(0.28), Inches(4.55), Inches(0.45),
         size=12, bold=True, colour=right_colour, align=PP_ALIGN.RIGHT, font="Georgia")
    text(slide, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         Inches(0.35), Inches(7.05), Inches(9.5), Inches(0.3), size=8, colour=BLACK)
    text(slide, DATE_STAMP, Inches(10.5), Inches(7.05), Inches(2.55), Inches(0.3),
         size=8, colour=BLACK, align=PP_ALIGN.RIGHT)


prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# ============================================================
# SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
# ============================================================
s1 = prs.slides.add_slide(blank)
bg(s1, WHITE)
chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE", right_colour=TEAL)

text(s1, "Taskforce Australia", Inches(0.35), Inches(1.15), Inches(9.5), Inches(0.75),
     size=40, bold=True, colour=BLACK, font="Georgia")
text(s1, "Property compliance and repair platform, Burnley, Melbourne VIC",
     Inches(0.35), Inches(1.85), Inches(11.5), Inches(0.4), size=14, colour=TEAL)

# Stat cards row: 5 cards, evenly spaced
cards = [
    ("$12.8M", "FY25 revenue", "SmartCompany Smart50 2025"),
    ("31%", "FY25 revenue growth", "SmartCompany Smart50 2025"),
    ("#37", "Smart50 national rank", "SmartCompany Smart50 2025"),
    ("20", "People, Melbourne team", "LinkedIn company page"),
    ("2014", "Year founded", "Smart50 award profile"),
]
card_w = Inches(2.3)
gap = Inches(0.15)
top = Inches(2.4)
card_h = Inches(1.6)
x = Inches(0.35)
for num, label, src in cards:
    rect(s1, x, top, card_w, card_h, BLACK)
    rect(s1, x, top, card_w, Pt(4), TEAL)
    text(s1, num, x + Inches(0.15), top + Inches(0.18), card_w - Inches(0.3), Inches(0.55),
         size=28, bold=True, colour=WHITE, font="Georgia")
    text(s1, label, x + Inches(0.15), top + Inches(0.78), card_w - Inches(0.3), Inches(0.5),
         size=11, colour=OFF_WHITE)
    text(s1, src, x + Inches(0.15), top + Inches(1.28), card_w - Inches(0.3), Inches(0.28),
         size=8, italic=True, colour=OFF_WHITE)
    x = x + card_w + gap

text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.25), Inches(6), Inches(0.3),
     size=12, bold=True, colour=TEAL)

signals = [
    "Revenue grew $7.43M to $12.8M, FY23 to FY25, three straight Smart50 profiles.",
    "Victorian State Winner, Outstanding Growth, 2024 Telstra Best of Business Awards.",
    "RentSafe has completed over 140,000 jobs across 180 real estate brands, 300 offices.",
    "About 4,500 to 5,000 vetted, insured tradespeople deliver RentSafe and RentRepair jobs.",
    "New distribution partnership announced with Real+ to expand into agency channels.",
    "Stated target: 50 to 60 percent compound growth over three years, no external capital.",
]
sig_lines = [{"text": "•  " + s, "size": 12, "colour": BLACK} for s in signals]
multiline(s1, sig_lines, Inches(0.35), Inches(4.6), Inches(12.6), Inches(2.35), space_after=6)

# ============================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================
s2 = prs.slides.add_slide(blank)
bg(s2, WHITE)
chrome(s2, "THE OPPORTUNITY", COMPANY.upper(), right_colour=TEAL)

text(s2, "Three commercial observations from ProfitPulse, for Taskforce Australia",
     Inches(0.35), Inches(1.12), Inches(12.6), Inches(0.4), size=14, bold=True, colour=BLACK, font="Georgia")

col_w = Inches(4.13)
col_gap = Inches(0.07)
col_top = Inches(1.65)
col_h = Inches(3.9)
cols_x = [Inches(0.35), Inches(0.35) + col_w + col_gap, Inches(0.35) + 2 * (col_w + col_gap)]
col_fills = [TEAL, BLACK, GOLD]
col_text_colours = [BLACK, WHITE, BLACK]
col_index = ["01", "02", "03"]
col_headers = [
    "Self funded growth is a cash strategy, not a slogan",
    "Revenue nearly doubled while the team barely grew",
    "A national contractor network is a timing problem",
]
col_bodies = [
    ("Taskforce has told the market it plans fifty to sixty percent compound growth over "
     "three years without raising external capital. That target means the working capital "
     "cycle, how fast agency clients pay against how fast the trade network is paid, is now "
     "doing the job a capital raise would otherwise do. A Working Capital Unlock maps where "
     "cash sits across debtors, supplier terms and facilities, and prioritises what releases "
     "it fastest."),
    ("Revenue rose from $7.43 million in FY23 to $12.8 million in FY25, over seventy percent, "
     "while headcount held near nineteen to twenty people per Smart50 and LinkedIn. That is "
     "strong productivity, and a sign the finance function built for a seven million dollar "
     "business is now running one worth nearly double. A 13 Week Cash Flow Build gives the "
     "next stage a forward view it currently runs without."),
    ("RentSafe and RentRepair route work through roughly 4,500 to 5,000 vetted tradespeople "
     "to serve 180 real estate brands and 300 offices. Every job creates two clocks, one for "
     "paying the tradesperson and one for collecting from the client, and the gap between them "
     "is where growth capital quietly leaks. A Fractional CFO Partnership keeps a senior "
     "finance voice on that gap every month, not just at each award cycle."),
]

for i in range(3):
    rect(s2, cols_x[i], col_top, col_w, col_h, col_fills[i])
    tc = col_text_colours[i]
    text(s2, col_index[i], cols_x[i] + Inches(0.25), col_top + Inches(0.2), col_w - Inches(0.5), Inches(0.6),
         size=30, bold=True, colour=tc, font="Georgia")
    text(s2, col_headers[i], cols_x[i] + Inches(0.25), col_top + Inches(0.85), col_w - Inches(0.5), Inches(0.85),
         size=14, bold=True, colour=tc, font="Georgia")
    text(s2, col_bodies[i], cols_x[i] + Inches(0.25), col_top + Inches(1.75), col_w - Inches(0.5), Inches(2.0),
         size=11, colour=tc)

text(s2, "These are observations offered in good faith. Taskforce has built something "
     "genuinely impressive in a short time. The question is simply whether the financial "
     "architecture keeps pace with the growth target.",
     Inches(0.35), Inches(5.85), Inches(12.6), Inches(0.42), size=10, italic=True, colour=TEAL)

# ============================================================
# SLIDE 3: THE RECOMMENDATION AND HOW TO START
# ============================================================
s3 = prs.slides.add_slide(blank)
bg(s3, WHITE)
chrome(s3, "THE RECOMMENDATION", COMPANY.upper(), right_colour=TEAL)

left_x = Inches(0.35)
left_w = Inches(7.6)

text(s3, "Working Capital Unlock", left_x, Inches(1.15), left_w, Inches(0.55),
     size=26, bold=True, colour=BLACK, font="Georgia")
text(s3, "$6,000 one off", left_x, Inches(1.68), left_w, Inches(0.4),
     size=18, bold=True, colour=AMBER_D, font="Georgia")
text(s3, "Maps cash trapped in debtors, contractor terms and facilities, with a prioritised "
     "list to release it, funding growth without external capital.",
     left_x, Inches(2.15), left_w, Inches(0.65), size=11.5, colour=BLACK)

rect(s3, left_x, Inches(2.9), left_w, Inches(1.85), TEAL)
text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2), Inches(3.04),
     left_w - Inches(0.4), Inches(0.35), size=13, bold=True, colour=BLACK, font="Georgia")
text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.2), Inches(3.42),
     left_w - Inches(0.4), Inches(0.35), size=11, colour=BLACK)
qbox = text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.2), Inches(3.8),
            left_w - Inches(0.4), Inches(0.4), size=13, bold=True, colour=BLACK)
add_hyperlink(qbox, "https://profit-pulse.com.au/services/find-your-fit/", colour=BLACK)
ctabox = text(s3, "Purchase the suggested product now to get started", left_x + Inches(0.2), Inches(4.22),
              left_w - Inches(0.4), Inches(0.35), size=11, bold=True, colour=BLACK)
add_hyperlink(ctabox, "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z", colour=BLACK)

text(s3, "Prefer a conversation first?", left_x, Inches(5.0), left_w, Inches(0.35),
     size=11, colour=BLACK)
bookbox = text(s3, "Book a complimentary discovery call", left_x, Inches(5.35), left_w, Inches(0.35),
               size=12, bold=True, colour=BLACK)
add_hyperlink(bookbox, "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true", colour=BLACK)

# Right column: credibility panel
right_x = Inches(8.25)
right_w = Inches(4.75)
rect(s3, right_x, Inches(1.15), right_w, Inches(4.75), BLACK)
rect(s3, right_x, Inches(1.15), right_w, Pt(4), TEAL)
text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.4), right_w - Inches(0.5), Inches(0.5),
     size=18, bold=True, colour=AMBER_B, font="Georgia")
text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.85), right_w - Inches(0.5), Inches(0.35),
     size=12, colour=WHITE)

cred_lines = [
    {"text": "•  16 years across 4 countries", "size": 11, "colour": OFF_WHITE},
    {"text": "•  52 deals executed and managed", "size": 11, "colour": OFF_WHITE},
    {"text": "•  Largest deal USD 1.3 billion, Cahora Bassa", "size": 11, "colour": OFF_WHITE},
    {"text": "•  Over AUD 10 billion in GRBT project value", "size": 11, "colour": OFF_WHITE},
]
multiline(s3, cred_lines, right_x + Inches(0.25), Inches(2.35), right_w - Inches(0.5), Inches(1.5), space_after=6)

rect(s3, right_x + Inches(0.25), Inches(4.0), right_w - Inches(0.5), Pt(1.5), TEAL)

contact_lines = [
    {"text": "Profit-Pulse.com.au", "size": 12, "colour": TEAL, "bold": True},
    {"text": "Nitesh@Profit-Pulse.com.au", "size": 12, "colour": OFF_WHITE},
    {"text": "+61 411 876 267", "size": 12, "colour": OFF_WHITE},
    {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10, "colour": OFF_WHITE},
]
multiline(s3, contact_lines, right_x + Inches(0.25), Inches(4.2), right_w - Inches(0.5), Inches(1.6), space_after=6)

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_30Jul2026.pptx"
prs.save(out_path)
print(f"PPTX saved: {out_path}")

# Force the theme hyperlink colours to brand black, since LibreOffice's PDF
# export renders hyperlinked runs using the theme hlink colour regardless of
# explicit run level font colour, and the default theme value is off brand blue.
import zipfile
import shutil

tmp_path = out_path + ".tmp"
with zipfile.ZipFile(out_path, "r") as zin:
    names = zin.namelist()
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            data = zin.read(name)
            if name == "ppt/theme/theme1.xml":
                text_xml = data.decode("utf-8")
                text_xml = text_xml.replace(
                    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>',
                    '<a:hlink><a:srgbClr val="000000"/></a:hlink>',
                )
                text_xml = text_xml.replace(
                    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>',
                    '<a:folHlink><a:srgbClr val="000000"/></a:folHlink>',
                )
                data = text_xml.encode("utf-8")
            zout.writestr(name, data)
shutil.move(tmp_path, out_path)
print("Theme hyperlink colours patched to brand black.")
