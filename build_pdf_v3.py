# -*- coding: utf-8 -*-
"""
ProfitPulse Brief PDF Builder, Version 3 house style. Direct PDF generation
(LibreOffice headless conversion is unavailable in this environment), built
to match build_brief_v3.py's PPTX layout exactly, driven by the same data
dict. Page size 960pt x 540pt, matching 13.333in x 7.5in at 72dpi.
"""
from reportlab.lib.colors import HexColor, Color
from reportlab.pdfgen import canvas

W = 960.0
H = 540.0
PT_PER_IN = 72.0

BLACK = HexColor("#000000")
TEAL = HexColor("#01A296")
AMBER_B = HexColor("#F8C806")
AMBER_D = HexColor("#F6A102")
GOLD = HexColor("#E3A712")
WHITE = HexColor("#FFFFFF")
OFF_WHITE = HexColor("#E6E5DE")
LIGHT_PANEL = HexColor("#F4F4F1")
MUTED = HexColor("#666666")
MUTED2 = HexColor("#9A9A9A")
DARK_GREY = HexColor("#333333")


def rl_y(y_top):
    return H - y_top


def fill_rect(c, x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.setStrokeColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, y_top, w, h, stroke_colour, line_width=1, fill_colour=None):
    c.setStrokeColor(stroke_colour)
    c.setLineWidth(line_width)
    if fill_colour:
        c.setFillColor(fill_colour)
        c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=1)
    else:
        c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def txt(c, text, x, y_top, size, colour, font="Helvetica-Bold", align="left", max_width=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline_y = rl_y(y_top + size)
    if align == "right" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + max_width - tw
    elif align == "center" and max_width:
        tw = c.stringWidth(text, font, size)
        x = x + (max_width - tw) / 2
    c.drawString(x, baseline_y, text)


def wrap_lines(c, text, max_w, size, font):
    words = text.split()
    lines, current = [], ""
    for word in words:
        test = (current + " " + word).strip()
        if c.stringWidth(test, font, size) <= max_w:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def txt_wrapped(c, text, x, y_top, max_w, size, colour, font="Helvetica",
                 leading=None, bullet=False):
    if leading is None:
        leading = size * 1.32
    c.setFillColor(colour)
    c.setFont(font, size)
    lines = wrap_lines(c, text, max_w, size, font)
    y = y_top
    for line in lines:
        c.drawString(x, rl_y(y + size), line)
        y += leading
    return y


def stat_card(c, x, y_top, w, h, number, label, source):
    fill_rect(c, x, y_top, w, h, BLACK)
    fill_rect(c, x, y_top, w, 3, TEAL)
    txt(c, number, x + 10, y_top + 13, 27, AMBER_B, font="Times-Bold")
    txt_wrapped(c, label, x + 10, y_top + 55, w - 20, 10.5, OFF_WHITE,
                font="Helvetica", leading=13.5)
    c.setFont("Helvetica-Oblique", 7.5)
    c.setFillColor(MUTED2)
    c.drawString(x + 10, rl_y(y_top + h - 12), source)


def chrome(c, eyebrow, right_label, date_str, prepared_line):
    fill_rect(c, 0, 0, W, H, WHITE)
    fill_rect(c, 0, 0, 7, H, AMBER_D)
    fill_rect(c, 7, 0, W - 7, 72, BLACK)
    txt(c, eyebrow, 26, 26, 12.5, OFF_WHITE, font="Helvetica-Bold")
    txt(c, right_label, W - 26, 26, 12.5, AMBER_B, font="Helvetica-Bold",
        align="right", max_width=400)
    c.setStrokeColor(HexColor("#CCCCCC"))
    c.setLineWidth(0.75)
    ry = rl_y(508)
    c.line(26, ry, 600, ry)
    txt(c, f"{prepared_line}    {date_str}", 26, 512, 7.5, MUTED, font="Helvetica")


def observation_column(c, x, y_top, w, h, fill, idx_colour, text_colour, index, header, body):
    fill_rect(c, x, y_top, w, h, fill)
    txt(c, index, x + 18, y_top + 14, 26, idx_colour, font="Times-Bold")
    hy = y_top + 66
    c.setFont("Times-Bold", 14)
    c.setFillColor(text_colour)
    hlines = wrap_lines(c, header, w - 36, 14, "Times-Bold")
    for line in hlines:
        c.drawString(x + 18, rl_y(hy + 14), line)
        hy += 17
    txt_wrapped(c, body, x + 18, hy + 10, w - 36, 10.5, text_colour,
                font="Helvetica", leading=14)


def link_rect(c, x, y_top, w, h, url):
    c.linkURL(url, (x, rl_y(y_top + h), x + w, rl_y(y_top)), relative=0, thickness=0)


def build(data, out_path):
    date_str = data["date_str"]
    prepared_line = ("Prepared by Nitesh Roopa CA, Managing Partner and Founder, "
                      "ProfitPulse, Profit-Pulse.com.au")
    c = canvas.Canvas(out_path, pagesize=(W, H))
    c.setTitle(f"{data['company_name']} | ProfitPulse Brief | {date_str}")
    c.setAuthor("ProfitPulse")
    c.setSubject(data["service_name"])

    # ───────────────────────── SLIDE 1 ─────────────────────────
    chrome(c, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE", date_str, prepared_line)
    txt(c, data["company_name"], 26, 83, 30, BLACK, font="Times-Bold")
    txt(c, data["descriptor"], 26, 133, 11, DARK_GREY, font="Helvetica")

    cards = data["stat_cards"]
    n = len(cards)
    gap = 11
    card_w = (W - 26 * 2 - gap * (n - 1)) / n
    card_top = 173
    card_h = 115
    x = 26
    for card in cards:
        stat_card(c, x, card_top, card_w, card_h, card["number"], card["label"], card["source"])
        x += card_w + gap

    txt(c, "KEY COMMERCIAL SIGNALS", 26, 313, 11.5, TEAL, font="Helvetica-Bold")
    y = 340
    for sig in data["signals"]:
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 10.5)
        c.drawString(26, rl_y(y + 10.5), "•")
        y_after = txt_wrapped(c, sig, 40, y, W - 26 - 40, 10.5, BLACK, font="Helvetica",
                               leading=14)
        y = y_after + 5
    c.showPage()

    # ───────────────────────── SLIDE 2 ─────────────────────────
    chrome(c, "THE OPPORTUNITY", data["company_name"], date_str, prepared_line)
    txt(c, "Three commercial observations from ProfitPulse", 26, 80, 12, HexColor("#444444"),
        font="Helvetica-Oblique")

    col_w = (W - 26 * 2 - 9 * 2) / 3
    col_top = 115
    col_h = 240
    fills = [TEAL, BLACK, GOLD]
    idx_colours = [BLACK, AMBER_B, BLACK]
    text_colours = [WHITE, WHITE, BLACK]
    x = 26
    for i, obs in enumerate(data["observations"]):
        observation_column(c, x, col_top, col_w, col_h, fills[i], idx_colours[i],
                            text_colours[i], obs["index"], obs["header"], obs["body"])
        x += col_w + 9

    txt_wrapped(c, data["warm_line"], 26, 388, W - 52, 10.5, HexColor("#444444"),
                font="Helvetica-Oblique", leading=14)
    c.showPage()

    # ───────────────────────── SLIDE 3 ─────────────────────────
    chrome(c, "THE RECOMMENDATION", data["company_name"], date_str, prepared_line)
    left_x, left_w = 26, 545

    c.setFont("Times-Bold", 18)
    c.setFillColor(BLACK)
    hlines = wrap_lines(c, data["service_name"], left_w, 18, "Times-Bold")
    hy = 86
    for line in hlines:
        c.drawString(left_x, rl_y(hy + 18), line)
        hy += 21
    txt(c, data["price_line"], left_x, hy + 6, 13, TEAL, font="Helvetica-Bold")
    txt_wrapped(c, data["service_blurb"], left_x, hy + 30, left_w, 10, BLACK,
                font="Helvetica", leading=13.5)

    step_top = hy + 92
    stroke_rect(c, left_x, step_top, left_w, 108, TEAL, line_width=1, fill_colour=LIGHT_PANEL)
    txt(c, "Step one, answer a few quick questions", left_x + 14, step_top + 12, 11.5, BLACK,
        font="Helvetica-Bold")
    txt(c, "See the solutions matched to your size and industry.", left_x + 14, step_top + 33,
        10.5, DARK_GREY, font="Helvetica")
    q_y = step_top + 58
    txt(c, "profit-pulse.com.au/full-suite-of-products", left_x + 14, q_y, 11.5, TEAL,
        font="Helvetica-Bold")
    link_rect(c, left_x + 14, q_y, 280, 16, "https://profit-pulse.com.au/full-suite-of-products")

    cta_y = step_top + 130
    txt(c, "Purchase the suggested product now to get started", left_x, cta_y, 12.5, AMBER_D,
        font="Helvetica-Bold")
    cta_w = c.stringWidth("Purchase the suggested product now to get started",
                           "Helvetica-Bold", 12.5)
    link_rect(c, left_x, cta_y, cta_w, 16, data["stripe_link"])

    conv_y = cta_y + 40
    txt(c, "Prefer a conversation first?", left_x, conv_y, 10.5, BLACK, font="Helvetica")
    txt(c, "Book a complimentary discovery call", left_x, conv_y + 22, 10.5, TEAL,
        font="Helvetica-Bold")
    bw = c.stringWidth("Book a complimentary discovery call", "Helvetica-Bold", 10.5)
    link_rect(c, left_x, conv_y + 22, bw, 14, data["booking_link"])

    right_x, right_w = 595, 339
    fill_rect(c, right_x, 83, right_w, 305, BLACK)
    txt(c, "Nitesh Roopa", right_x + 18, 102, 16, AMBER_B, font="Times-Bold")
    txt(c, "CA, Managing Partner, ProfitPulse", right_x + 18, 128, 10.5, WHITE, font="Helvetica")

    cy = 162
    for cred in data["credibility_points"]:
        c.setFillColor(OFF_WHITE)
        c.setFont("Helvetica-Bold", 10.5)
        c.drawString(right_x + 18, rl_y(cy + 10.5), "•")
        cy_after = txt_wrapped(c, cred, right_x + 32, cy, right_w - 50, 10.5, OFF_WHITE,
                                font="Helvetica", leading=14)
        cy = cy_after + 6

    c.setStrokeColor(TEAL)
    c.setLineWidth(1)
    ry = rl_y(cy + 14)
    c.line(right_x + 18, ry, right_x + right_w - 18, ry)

    contact_y = cy + 30
    contact_lines = [
        ("Profit-Pulse.com.au", OFF_WHITE),
        ("Nitesh@Profit-Pulse.com.au", TEAL),
        ("+61 411 876 267", OFF_WHITE),
        ("linkedin.com/in/nitesh-roopa-77594163", TEAL),
    ]
    for line, colour in contact_lines:
        txt(c, line, right_x + 18, contact_y, 10, colour, font="Helvetica")
        contact_y += 19

    c.showPage()
    c.save()
    print(f"PDF saved: {out_path}")


if __name__ == "__main__":
    print("Import this module and call build(data, out_path) with real data.")
