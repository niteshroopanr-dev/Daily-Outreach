"""
ProfitPulse nightly outreach brief builder, V3.3 house style.
Three slide deck. White body, black header band, amber left stripe,
black stat cards with a teal top edge. Brand colours only. Zero dashes.
Data driven: edit the DATA dict at the bottom (or import and call build()).
"""

import copy
import zipfile
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

# Brand colours, the only seven allowed anywhere in this deck.
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)
HEADER_H = Inches(1.0)
FOOTER_Y = Inches(7.05)
CONTENT_TOP = Inches(1.15)
CONTENT_BOTTOM = Inches(6.9)
STRIPE_W = Inches(0.1)


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
    # Strip the theme <p:style> block entirely, its effectRef otherwise renders a
    # shadow in LibreOffice even with an empty effectLst on spPr.
    style_el = shape._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style"
    )
    if style_el is not None:
        shape._element.remove(style_el)
    return shape


def add_text(slide, text, left, top, width, height, font_name="Calibri", font_size=12,
             bold=False, colour=BLACK, align=PP_ALIGN.LEFT, wrap=True, italic=False,
             line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n")
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = colour
    return box


def add_multiline(slide, lines, left, top, width, height, font_name="Calibri",
                   default_size=12, default_colour=BLACK, align=PP_ALIGN.LEFT,
                   space_after=2, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for line in lines:
        cfg = {"text": line, "size": default_size, "colour": default_colour,
               "bold": False, "italic": False} if isinstance(line, str) else {
            "text": line.get("text", ""), "size": line.get("size", default_size),
            "colour": line.get("colour", default_colour), "bold": line.get("bold", False),
            "italic": line.get("italic", False)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
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
    return box


def add_hyperlink_note(shape, url):
    """Attach a hyperlink to the run(s) inside a textbox shape (whole frame)."""
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.hyperlink.address = url


def fix_theme_hyperlink_colours(pptx_path):
    """PowerPoint themes default a:hlink/a:folHlink to blue/purple. Renderers such
    as LibreOffice apply that theme colour to any hyperlinked run regardless of
    direct run formatting, which would put a non brand blue on the slide. Force
    both to brand black so hyperlinked text stays within the seven allowed colours."""
    tmp_path = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/theme/theme") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text = text.replace(
                    "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>",
                    "<a:hlink><a:srgbClr val=\"000000\"/></a:hlink>",
                )
                text = text.replace(
                    "<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>",
                    "<a:folHlink><a:srgbClr val=\"000000\"/></a:folHlink>",
                )
                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp_path, pptx_path)


def house_chrome(slide, eyebrow, right_label, date_str):
    """Left accent stripe, black header band, footer line. Call first on every slide."""
    set_background(slide, WHITE)
    add_rect(slide, Inches(0), Inches(0), STRIPE_W, H, AMBER_D)
    add_rect(slide, STRIPE_W, Inches(0), W - STRIPE_W, HEADER_H, BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(8.5), Inches(0.4),
              font_name="Calibri", font_size=12, bold=True, colour=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, right_label, Inches(9.0), Inches(0.32), Inches(4.0), Inches(0.4),
              font_name="Calibri", font_size=12, bold=True, colour=WHITE, align=PP_ALIGN.RIGHT)
    add_rect(slide, Inches(0.35), FOOTER_Y, Inches(12.6), Pt(0.75), TEAL)
    add_text(slide, "Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse, Profit-Pulse.com.au",
              Inches(0.35), Inches(7.14), Inches(9.0), Inches(0.3),
              font_size=8, colour=BLACK, align=PP_ALIGN.LEFT)
    add_text(slide, date_str, Inches(9.6), Inches(7.14), Inches(3.35), Inches(0.3),
              font_size=8, colour=BLACK, align=PP_ALIGN.RIGHT)


def stat_card(slide, x, y, w, h, number, label, source):
    add_rect(slide, x, y, w, h, BLACK)
    add_rect(slide, x, y, w, Pt(4), TEAL)
    add_text(slide, number, x + Inches(0.12), y + Inches(0.14), w - Inches(0.24), Inches(0.5),
              font_name="Cambria", font_size=27, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_text(slide, label, x + Inches(0.12), y + Inches(0.68), w - Inches(0.24), Inches(0.55),
              font_name="Calibri", font_size=10.5, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT,
              line_spacing=1.0)
    add_text(slide, source, x + Inches(0.12), y + h - Inches(0.32), w - Inches(0.24), Inches(0.28),
              font_name="Calibri", font_size=7.5, bold=False, colour=OFF_WHITE, align=PP_ALIGN.LEFT,
              italic=True)


def build(data, out_path):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]
    date_str = data["date_str"]
    company = data["company"]

    # ---------------- SLIDE 1 ----------------
    s1 = prs.slides.add_slide(blank)
    house_chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE", date_str)

    add_text(s1, company, Inches(0.35), Inches(1.2), Inches(9.0), Inches(0.75),
              font_name="Cambria", font_size=40, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
    add_text(s1, data["descriptor"], Inches(0.35), Inches(1.95), Inches(12.6), Inches(0.35),
              font_name="Calibri", font_size=13, bold=False, colour=TEAL, align=PP_ALIGN.LEFT)

    cards = data["stat_cards"]
    n = len(cards)
    gap = Inches(0.15)
    card_w = Emu(int((W - Inches(0.7) - gap * (n - 1)) / n))
    card_h = Inches(1.6)
    card_y = Inches(2.4)
    x = Inches(0.35)
    for c in cards:
        stat_card(s1, x, card_y, card_w, card_h, c["number"], c["label"], c["source"])
        x = Emu(int(x + card_w + gap))

    add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.28), Inches(6), Inches(0.3),
              font_name="Calibri", font_size=12, bold=True, colour=TEAL, align=PP_ALIGN.LEFT)
    signal_lines = [{"text": "•  " + s, "size": 11.5, "colour": BLACK} for s in data["signals"]]
    add_multiline(s1, signal_lines, Inches(0.35), Inches(4.68), Inches(12.6), Inches(2.15),
                   font_name="Calibri", default_size=11.5, default_colour=BLACK, space_after=6,
                   line_spacing=1.05)

    # ---------------- SLIDE 2 ----------------
    s2 = prs.slides.add_slide(blank)
    house_chrome(s2, "THE OPPORTUNITY", company, date_str)
    add_text(s2, "Three commercial observations from ProfitPulse", Inches(0.35), Inches(1.02),
              Inches(9.0), Inches(0.3), font_name="Calibri", font_size=12, bold=False,
              colour=BLACK, italic=True, align=PP_ALIGN.LEFT)

    col_top = Inches(1.42)
    col_h = Inches(4.95)
    col_gap = Inches(0.12)
    col_w = Emu(int((W - Inches(0.7) - col_gap * 2) / 3))
    fills = [TEAL, BLACK, GOLD]
    text_colours = [BLACK, OFF_WHITE, BLACK]
    idx_colours = [WHITE, AMBER_B, WHITE]
    x = Inches(0.35)
    for i, obs in enumerate(data["observations"]):
        add_rect(s2, x, col_top, col_w, col_h, fills[i])
        add_text(s2, f"0{i+1}", x + Inches(0.2), col_top + Inches(0.18), col_w - Inches(0.4), Inches(0.7),
                  font_name="Cambria", font_size=34, bold=True, colour=idx_colours[i], align=PP_ALIGN.LEFT)
        add_text(s2, obs["header"], x + Inches(0.2), col_top + Inches(0.95), col_w - Inches(0.4), Inches(0.75),
                  font_name="Calibri", font_size=14, bold=True, colour=text_colours[i], align=PP_ALIGN.LEFT,
                  line_spacing=1.0)
        add_text(s2, obs["body"], x + Inches(0.2), col_top + Inches(1.8), col_w - Inches(0.4), Inches(3.0),
                  font_name="Calibri", font_size=10.5, bold=False, colour=text_colours[i], align=PP_ALIGN.LEFT,
                  line_spacing=1.08)
        x = Emu(int(x + col_w + col_gap))

    add_text(s2, data["warm_line"], Inches(0.35), Inches(6.5), Inches(12.6), Inches(0.45),
              font_name="Calibri", font_size=10.5, bold=False, italic=True, colour=BLACK,
              align=PP_ALIGN.LEFT)

    # ---------------- SLIDE 3 ----------------
    s3 = prs.slides.add_slide(blank)
    house_chrome(s3, "THE RECOMMENDATION", company, date_str)

    left_x = Inches(0.35)
    left_w = Inches(7.5)
    add_text(s3, data["service_name"], left_x, Inches(1.25), left_w, Inches(0.65),
              font_name="Cambria", font_size=22, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
    add_text(s3, data["service_price_line"], left_x, Inches(1.9), left_w, Inches(0.4),
              font_name="Calibri", font_size=15, bold=True, colour=AMBER_D, align=PP_ALIGN.LEFT)
    add_text(s3, data["service_desc"], left_x, Inches(2.35), left_w, Inches(1.0),
              font_name="Calibri", font_size=11, bold=False, colour=BLACK, align=PP_ALIGN.LEFT,
              line_spacing=1.1)

    add_rect(s3, left_x, Inches(3.5), left_w, Inches(1.35), TEAL)
    add_text(s3, "STEP ONE, ANSWER A FEW QUICK QUESTIONS", left_x + Inches(0.2), Inches(3.62),
              left_w - Inches(0.4), Inches(0.3), font_name="Calibri", font_size=11, bold=True,
              colour=BLACK, align=PP_ALIGN.LEFT)
    add_text(s3, "See the solutions matched to your size and industry.", left_x + Inches(0.2),
              Inches(3.95), left_w - Inches(0.4), Inches(0.3), font_name="Calibri", font_size=10.5,
              colour=BLACK, align=PP_ALIGN.LEFT)
    q_box = add_text(s3, "profit-pulse.com.au/services/find-your-fit", left_x + Inches(0.2),
              Inches(4.3), left_w - Inches(0.4), Inches(0.4), font_name="Calibri", font_size=12,
              bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
    add_hyperlink_note(q_box, "https://profit-pulse.com.au/services/find-your-fit/")

    cta_box = add_rect(s3, left_x, Inches(5.05), Inches(4.6), Inches(0.55), AMBER_D)
    cta_text = add_text(s3, "Purchase the suggested product now to get started",
              left_x + Inches(0.15), Inches(5.18), Inches(4.3), Inches(0.32),
              font_name="Calibri", font_size=10.5, bold=True, colour=BLACK, align=PP_ALIGN.LEFT)
    add_hyperlink_note(cta_text, data["stripe_link"])

    add_text(s3, "Prefer a conversation first? Book a complimentary discovery call.",
              left_x, Inches(5.85), left_w, Inches(0.35), font_name="Calibri", font_size=10.5,
              colour=BLACK, align=PP_ALIGN.LEFT)
    book_box = add_text(s3, "Book a complimentary discovery call", left_x, Inches(6.18),
              left_w, Inches(0.35), font_name="Calibri", font_size=11, bold=True, colour=TEAL,
              align=PP_ALIGN.LEFT)
    add_hyperlink_note(book_box, data["booking_link"])

    right_x = Inches(8.35)
    right_w = Inches(4.6)
    add_rect(s3, right_x, Inches(1.25), right_w, Inches(5.5), BLACK)
    add_text(s3, "NITESH ROOPA", right_x + Inches(0.25), Inches(1.45), right_w - Inches(0.5), Inches(0.4),
              font_name="Cambria", font_size=17, bold=True, colour=AMBER_B, align=PP_ALIGN.LEFT)
    add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.88),
              right_w - Inches(0.5), Inches(0.35), font_name="Calibri", font_size=12, colour=WHITE,
              align=PP_ALIGN.LEFT)
    add_rect(s3, right_x + Inches(0.25), Inches(2.28), right_w - Inches(0.7), Pt(1.5), TEAL)
    cred_lines = [
        {"text": "16 years across 4 countries", "size": 11, "colour": OFF_WHITE},
        {"text": "52 deals executed and managed", "size": 11, "colour": OFF_WHITE},
        {"text": "Largest single deal, USD 1.3 billion, Cahora Bassa", "size": 11, "colour": OFF_WHITE},
        {"text": "Total GRBT project value over AUD 10 billion", "size": 11, "colour": OFF_WHITE},
    ]
    add_multiline(s3, cred_lines, right_x + Inches(0.25), Inches(2.45), right_w - Inches(0.5), Inches(1.7),
                   font_name="Calibri", default_size=11, default_colour=OFF_WHITE, space_after=5,
                   line_spacing=1.05)
    contact_lines = [
        {"text": "Profit-Pulse.com.au", "size": 11.5, "colour": TEAL},
        {"text": "Nitesh@Profit-Pulse.com.au", "size": 11.5, "colour": TEAL},
        {"text": "+61 411 876 267", "size": 11.5, "colour": OFF_WHITE},
        {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10.5, "colour": OFF_WHITE},
    ]
    add_multiline(s3, contact_lines, right_x + Inches(0.25), Inches(4.35), right_w - Inches(0.5), Inches(1.6),
                   font_name="Calibri", default_size=11.5, default_colour=OFF_WHITE, space_after=6,
                   line_spacing=1.1)

    prs.save(out_path)
    fix_theme_hyperlink_colours(out_path)
    print(f"PPTX saved: {out_path}")
    return out_path
