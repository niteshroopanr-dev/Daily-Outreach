"""
ProfitPulse Brief Builder, Version 3 house style.
Three slide prospect facing deck. White body, black header band, amber left
stripe, black stat tiles with teal top accent. Brand colours only. No dashes.
Fill in DATA below and run: python3 build_brief_v3.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_AUTO_SIZE

# Brand colours, the only seven permitted anywhere in this deck.
BLACK     = RGBColor(0x00, 0x00, 0x00)
TEAL      = RGBColor(0x01, 0xA2, 0x96)
AMBER_B   = RGBColor(0xF8, 0xC8, 0x06)
AMBER_D   = RGBColor(0xF6, 0xA1, 0x02)
GOLD      = RGBColor(0xE3, 0xA7, 0x12)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xE6, 0xE5, 0xDE)

W = Inches(13.333)
H = Inches(7.5)

SERIF = "Cambria"
SANS = "Calibri"


def set_background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_rect(slide, left, top, width, height, fill_colour, line_colour=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, font_name=SANS,
             font_size=14, bold=False, colour=BLACK, align=PP_ALIGN.LEFT,
             italic=False, line_spacing=None):
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
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    return box


def add_multiline(slide, paras, left, top, width, height, align=PP_ALIGN.LEFT,
                   line_spacing=None):
    """paras: list of dicts text,size,colour,bold,italic,font,space_after"""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    first = True
    for cfg in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        if cfg.get("space_after"):
            p.space_after = Pt(cfg["space_after"])
        run = p.add_run()
        run.text = cfg["text"]
        run.font.name = cfg.get("font", SANS)
        run.font.size = Pt(cfg.get("size", 12))
        run.font.bold = cfg.get("bold", False)
        run.font.italic = cfg.get("italic", False)
        run.font.color.rgb = cfg.get("colour", BLACK)
    return box


def add_hyperlink_text(slide, text, url, left, top, width, height,
                        font_size=12, bold=False, colour=TEAL,
                        font_name=SANS, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
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


def chrome(slide, eyebrow, right_label, date_str, prepared_line):
    """Fixed chrome: left stripe, header band, footer line. Identical geometry
    on every slide per Section 6.0A."""
    set_background(slide, WHITE)
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), H, AMBER_D)
    add_rect(slide, Inches(0.1), Inches(0), W - Inches(0.1), Inches(1.0), BLACK)
    add_text(slide, eyebrow, Inches(0.35), Inches(0.32), Inches(7.5), Inches(0.4),
              font_name=SANS, font_size=12, bold=True, colour=OFF_WHITE)
    add_text(slide, right_label, Inches(7.5), Inches(0.32), Inches(5.5), Inches(0.4),
              font_name=SANS, font_size=12, bold=True, colour=AMBER_B,
              align=PP_ALIGN.RIGHT)
    add_rect(slide, Inches(0.35), Inches(7.0), Inches(8.0), Pt(0.75),
             RGBColor(0xCC, 0xCC, 0xCC))
    add_multiline(
        slide,
        [{"text": f"{prepared_line}    {date_str}", "size": 8,
          "colour": RGBColor(0x66, 0x66, 0x66), "font": SANS}],
        Inches(0.35), Inches(7.08), Inches(12.6), Inches(0.3),
    )


def stat_card(slide, left, top, width, height, number, label_lines, source):
    add_rect(slide, left, top, width, height, BLACK)
    add_rect(slide, left, top, width, Pt(4), TEAL)
    add_text(slide, number, left + Inches(0.15), top + Inches(0.18),
              width - Inches(0.3), Inches(0.5), font_name=SERIF, font_size=27,
              bold=True, colour=AMBER_B)
    add_multiline(
        slide,
        [{"text": label_lines, "size": 11, "colour": OFF_WHITE, "font": SANS}],
        left + Inches(0.15), top + Inches(0.72), width - Inches(0.3), Inches(0.55),
        line_spacing=1.05,
    )
    add_text(slide, source, left + Inches(0.15), top + height - Inches(0.32),
              width - Inches(0.3), Inches(0.28), font_name=SANS, font_size=7,
              italic=True, colour=RGBColor(0x9A, 0x9A, 0x9A))


def build(data, out_path):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]

    date_str = data["date_str"]
    prepared_line = "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au"

    # ───────────────────────── SLIDE 1 ─────────────────────────
    s1 = prs.slides.add_slide(blank)
    chrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE", date_str, prepared_line)

    add_text(s1, data["company_name"], Inches(0.35), Inches(1.15), Inches(11.5),
              Inches(0.7), font_name=SERIF, font_size=40, bold=True, colour=BLACK)
    add_text(s1, data["descriptor"], Inches(0.35), Inches(1.85), Inches(11.5),
              Inches(0.4), font_name=SANS, font_size=14, colour=RGBColor(0x33, 0x33, 0x33))

    cards = data["stat_cards"]
    n = len(cards)
    gap = Inches(0.15)
    card_w = Inches((13.333 - 0.35 * 2 - 0.15 * (n - 1)) / n)
    card_top = Inches(2.4)
    card_h = Inches(1.6)
    x = Inches(0.35)
    for c in cards:
        stat_card(s1, x, card_top, card_w, card_h, c["number"], c["label"], c["source"])
        x = Inches(x.inches + card_w.inches + gap.inches)

    add_text(s1, "KEY COMMERCIAL SIGNALS", Inches(0.35), Inches(4.35), Inches(8),
              Inches(0.3), font_name=SANS, font_size=12, bold=True, colour=TEAL)
    sig_paras = [{"text": "•  " + s, "size": 11.5, "colour": BLACK, "font": SANS,
                  "space_after": 7} for s in data["signals"]]
    add_multiline(s1, sig_paras, Inches(0.35), Inches(4.7), Inches(12.6), Inches(2.2),
                  line_spacing=1.08)

    # ───────────────────────── SLIDE 2 ─────────────────────────
    s2 = prs.slides.add_slide(blank)
    chrome(s2, "THE OPPORTUNITY", data["company_name"], date_str, prepared_line)
    add_text(s2, "Three commercial observations from ProfitPulse",
              Inches(0.35), Inches(1.12), Inches(12.6), Inches(0.35),
              font_name=SANS, font_size=13, italic=True, colour=RGBColor(0x44, 0x44, 0x44))

    col_w = Inches(4.05)
    col_gap = Inches(0.12)
    col_top = Inches(1.6)
    col_h = Inches(4.95)
    fills = [TEAL, BLACK, GOLD]
    text_colours = [WHITE, WHITE, BLACK]
    idx_colours = [BLACK, AMBER_B, BLACK]
    x = Inches(0.35)
    for i, obs in enumerate(data["observations"]):
        add_rect(s2, x, col_top, col_w, col_h, fills[i])
        add_text(s2, obs["index"], x + Inches(0.25), col_top + Inches(0.2),
                  Inches(1.5), Inches(0.6), font_name=SERIF, font_size=30, bold=True,
                  colour=idx_colours[i])
        add_text(s2, obs["header"], x + Inches(0.25), col_top + Inches(0.95),
                  col_w - Inches(0.5), Inches(0.7), font_name=SERIF, font_size=15.5,
                  bold=True, colour=text_colours[i], line_spacing=1.05)
        add_text(s2, obs["body"], x + Inches(0.25), col_top + Inches(1.75),
                  col_w - Inches(0.5), Inches(3.0), font_name=SANS, font_size=10.5,
                  colour=text_colours[i], line_spacing=1.18)
        x = Inches(x.inches + col_w.inches + col_gap.inches)

    add_text(s2, data["warm_line"], Inches(0.35), Inches(6.65), Inches(12.6),
              Inches(0.35), font_name=SANS, font_size=10.5, italic=True,
              colour=RGBColor(0x44, 0x44, 0x44))

    # ───────────────────────── SLIDE 3 ─────────────────────────
    s3 = prs.slides.add_slide(blank)
    chrome(s3, "THE RECOMMENDATION", data["company_name"], date_str, prepared_line)

    left_x = Inches(0.35)
    left_w = Inches(7.6)

    add_text(s3, data["service_name"], left_x, Inches(1.15), left_w, Inches(0.65),
              font_name=SERIF, font_size=21, bold=True, colour=BLACK, line_spacing=1.0)
    add_text(s3, data["price_line"], left_x, Inches(1.78), left_w, Inches(0.35),
              font_name=SANS, font_size=14, bold=True, colour=TEAL)
    add_text(s3, data["service_blurb"], left_x, Inches(2.15), left_w, Inches(0.85),
              font_name=SANS, font_size=11, colour=BLACK, line_spacing=1.15)

    add_rect(s3, left_x, Inches(3.1), left_w, Inches(1.5), RGBColor(0xF4, 0xF4, 0xF1),
             line_colour=TEAL)
    add_text(s3, "Step one, answer a few quick questions", left_x + Inches(0.2),
              Inches(3.25), left_w - Inches(0.4), Inches(0.35), font_name=SANS,
              font_size=12, bold=True, colour=BLACK)
    add_text(s3, "See the solutions matched to your size and industry.",
              left_x + Inches(0.2), Inches(3.62), left_w - Inches(0.4), Inches(0.35),
              font_name=SANS, font_size=10.5, colour=RGBColor(0x33, 0x33, 0x33))
    add_hyperlink_text(s3, "profit-pulse.com.au/full-suite-of-products",
                        "https://profit-pulse.com.au/full-suite-of-products",
                        left_x + Inches(0.2), Inches(4.0), left_w - Inches(0.4), Inches(0.4),
                        font_size=12, bold=True, colour=TEAL)

    add_hyperlink_text(s3, "Purchase the suggested product now to get started",
                        data["stripe_link"], left_x, Inches(4.85), left_w, Inches(0.45),
                        font_size=13, bold=True, colour=AMBER_D)

    add_text(s3, "Prefer a conversation first?", left_x, Inches(5.55), left_w,
              Inches(0.32), font_name=SANS, font_size=11, colour=BLACK)
    add_hyperlink_text(s3, "Book a complimentary discovery call",
                        data["booking_link"], left_x, Inches(5.88), left_w, Inches(0.35),
                        font_size=11, bold=True, colour=TEAL)

    # Right column, credibility panel
    right_x = Inches(8.25)
    right_w = Inches(4.73)
    add_rect(s3, right_x, Inches(1.15), right_w, Inches(5.55), BLACK)
    add_text(s3, "Nitesh Roopa", right_x + Inches(0.25), Inches(1.4), right_w - Inches(0.5),
              Inches(0.4), font_name=SERIF, font_size=17, bold=True, colour=AMBER_B)
    add_text(s3, "CA, Managing Partner, ProfitPulse", right_x + Inches(0.25), Inches(1.82),
              right_w - Inches(0.5), Inches(0.32), font_name=SANS, font_size=11,
              colour=WHITE)
    cred_paras = [{"text": "•  " + c, "size": 10, "colour": OFF_WHITE, "font": SANS,
                   "space_after": 6} for c in data["credibility_points"]]
    add_multiline(s3, cred_paras, right_x + Inches(0.25), Inches(2.3), right_w - Inches(0.5),
                  Inches(1.9), line_spacing=1.1)

    add_rect(s3, right_x + Inches(0.25), Inches(4.3), right_w - Inches(0.5), Pt(1), TEAL)
    contact_paras = [
        {"text": "Profit-Pulse.com.au", "size": 10.5, "colour": OFF_WHITE, "font": SANS, "space_after": 4},
        {"text": "Nitesh@Profit-Pulse.com.au", "size": 10.5, "colour": TEAL, "font": SANS, "space_after": 4},
        {"text": "+61 411 876 267", "size": 10.5, "colour": OFF_WHITE, "font": SANS, "space_after": 4},
        {"text": "linkedin.com/in/nitesh-roopa-77594163", "size": 10.5, "colour": TEAL, "font": SANS},
    ]
    add_multiline(s3, contact_paras, right_x + Inches(0.25), Inches(4.5), right_w - Inches(0.5),
                  Inches(1.5))

    prs.save(out_path)
    print(f"PPTX saved: {out_path}")


if __name__ == "__main__":
    print("Import this module and call build(data, out_path) with real data.")
