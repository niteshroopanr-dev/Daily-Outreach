"""
ProfitPulse Brief PDF Builder
Target: Orbitkey | Date: 05 Aug 2026
Mirrors Brief_Orbitkey_05Aug2026.pptx exactly. Brand colours only, zero dashes.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi)
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
import textwrap

W = 960
H = 540
IN = 72  # points per inch

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WHITE= HexColor("#E6E5DE")
C_TXT1     = HexColor("#222222")
C_TXT2     = HexColor("#333333")
C_MUTED    = HexColor("#666666")
C_PANEL    = HexColor("#F2F2F0")
C_SRC      = HexColor("#AAAAAA")

DATE_STR = "05 Aug 2026"

c = canvas.Canvas("/home/user/Daily-Outreach/Out-reach efforts/Brief_Orbitkey_05Aug2026.pdf",
                   pagesize=(W, H))


def y_(top_in):
    """Convert a top offset in inches to reportlab bottom origin y in points."""
    return H - top_in * IN


def rect(x_in, top_in, w_in, h_in, colour, stroke=None):
    c.setFillColor(colour)
    x = x_in * IN
    y = y_(top_in) - h_in * IN
    w = w_in * IN
    h = h_in * IN
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(1)
        c.rect(x, y, w, h, fill=1, stroke=1)
    else:
        c.rect(x, y, w, h, fill=1, stroke=0)


def text(s, x_in, top_in, size, colour, bold=False, font_base="Helvetica",
         align="left", max_w_in=None, leading=None):
    font = font_base + ("-Bold" if bold else "")
    c.setFillColor(colour)
    c.setFont(font, size)
    x = x_in * IN
    baseline = y_(top_in) - size
    if align == "center" and max_w_in:
        tw = c.stringWidth(s, font, size)
        x = x_in * IN + (max_w_in * IN - tw) / 2
    elif align == "right" and max_w_in:
        tw = c.stringWidth(s, font, size)
        x = x_in * IN + max_w_in * IN - tw
    c.drawString(x, baseline, s)


def text_wrapped(s, x_in, top_in, max_w_in, size, colour, bold=False,
                  font_base="Helvetica", leading_mult=1.15, char_w=None):
    font = font_base + ("-Bold" if bold else "")
    c.setFont(font, size)
    if char_w is None:
        avg_char_w = c.stringWidth("n", font, size)
        wrap_chars = max(10, int((max_w_in * IN) / avg_char_w))
    else:
        wrap_chars = char_w
    lines = []
    for para in s.split("\n\n"):
        wrapped = textwrap.wrap(para, wrap_chars) if para else [""]
        lines.extend(wrapped)
        lines.append("")
    if lines and lines[-1] == "":
        lines.pop()
    y = top_in
    for ln in lines:
        text(ln, x_in, y, size, colour, bold=bold, font_base=font_base)
        y += (size * leading_mult) / IN
    return y


def chrome(eyebrow, header_right="Orbitkey"):
    rect(0, 0, 13.333, 7.5, C_WHITE)
    rect(0, 0, 0.1, 7.5, C_AMBER_D)
    rect(0, 0, 13.333, 1.0, C_BLACK)
    text(eyebrow, 0.3, 0.65, 12, C_OFF_WHITE, bold=True)
    text(header_right, 8.5, 0.62, 15, C_TEAL, bold=True, font_base="Times", align="right", max_w_in=4.53)
    text("Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         0.3, 7.28, 8, C_MUTED)
    text(DATE_STR, 11.5, 7.28, 8, C_MUTED, align="right", max_w_in=1.53)


# ============================================================
# SLIDE 1
# ============================================================
chrome("COMMERCIAL INTELLIGENCE BRIEF")
text("Orbitkey", 0.3, 1.25, 40, C_BLACK, bold=True, font_base="Times")
text("Design led everyday carry accessories brand, Richmond, Melbourne VIC",
     0.3, 2.05, 14, C_TXT2)

cards = [
    ("$24.6M", "Annual revenue", "Smart50 2025 award"),
    ("44", "Team members", "Smart50 2025 award"),
    ("40%", "3 year revenue growth", "Export Awards 2025"),
    ("100+", "Countries exported to", "Export Awards 2025"),
    ("2013", "Founded", "SmartCompany history"),
]
card_w = 2.3
gap = 0.15
top = 2.4
card_h = 1.6
x = 0.3
for num, lbl, src in cards:
    rect(x, top, card_w, card_h, C_BLACK)
    rect(x, top, card_w, 4 / IN, C_TEAL)
    text(num, x + 0.15, top + 0.5, 26, C_AMBER_B, bold=True, font_base="Times")
    text(lbl, x + 0.15, top + 0.95, 11, C_WHITE)
    text(src, x + 0.15, top + 1.32, 7, C_SRC)
    x += card_w + gap

text("KEY COMMERCIAL SIGNALS", 0.3, 4.55, 12, C_TEAL, bold=True)
signals = [
    "Ranked 50th nationally, 2025 Smart50 Awards, fastest growing small businesses (SmartCompany).",
    "Winner, Creative Industries, Australian Export Awards 2025, Victorian state final (Export Awards).",
    "Ships from five third party logistics warehouses to over 100 export countries (Export Awards 2025).",
    "Raised over USD 3.25 million across six Kickstarter campaigns from 47,068 backers (SmartCompany).",
    "Registered as Orbitkey Pty Ltd, ABN 15 634 506 059, Richmond VIC (ASIC public register).",
]
y = 4.9
for sig in signals:
    rect(0.3, y + 0.02, 0.06, 0.06, C_AMBER_D)
    text(sig, 0.52, y + 0.09, 11, C_TXT1)
    y += 0.32
c.showPage()

# ============================================================
# SLIDE 2
# ============================================================
chrome("THE OPPORTUNITY")
text("Orbitkey. Three commercial observations from ProfitPulse.",
     0.3, 1.5, 16, C_BLACK, bold=True, font_base="Times")

col_w = 4.15
col_gap = 0.08
col_top = 1.75
col_h = 5.0
cols = [
    (C_TEAL, C_WHITE, "01", "Serial crowdfunding signals a cash gap",
     "Orbitkey's own founders say Kickstarter helps fund new products and manage cash flow around production. That is a direct signal that cash tied up in inventory, debtors and supplier terms across five global warehouses deserves a closer look.\n\n"
     "Working Capital Unlock maps exactly this: cash trapped in stock, work in progress and terms, with a prioritised release plan. Comparable clients typically release 8 to 15 percent of revenue in freed up cash."),
    (C_BLACK, C_OFF_WHITE, "02", "Three product lines, one margin picture",
     "Orbitkey sells across Everyday Carry, Travel and Bag Organisation, and Work Organisation, through direct to consumer, Kickstarter and wholesale to over a thousand European stores. Three channels and three product families rarely carry identical margin.\n\n"
     "A line by line profitability view would show which products and channels are genuinely funding the growth, and which are being carried by the others."),
    (C_GOLD, C_BLACK, "03", "Fast international growth needs a costed plan",
     "Revenue has grown 40 percent over three years and Orbitkey now exports to more than 100 countries from five logistics warehouses. That pace of expansion usually outruns the financial plan supporting it.\n\n"
     "A Strategic Growth Diagnostic would map revenue, capacity and margin headroom into a costed 12 month plan, so the next stage of growth is funded on purpose."),
]
x = 0.3
for fill, txtcol, idx, head, body in cols:
    rect(x, col_top, col_w, col_h, fill)
    text(idx, x + 0.25, col_top + 0.55, 26, txtcol, bold=True, font_base="Times")
    text_wrapped(head, x + 0.25, col_top + 0.95, col_w - 0.5, 13, txtcol, bold=True, leading_mult=1.15)
    text_wrapped(body, x + 0.25, col_top + 1.55, col_w - 0.5, 10, txtcol, leading_mult=1.25)
    x += col_w + col_gap

text_wrapped("These are observations offered in good faith. Orbitkey has built something genuinely impressive. The question is simply whether the financial architecture matches the pace of the growth.",
             0.3, 6.85, 12.7, 10, C_MUTED)
c.showPage()

# ============================================================
# SLIDE 3
# ============================================================
chrome("THE RECOMMENDATION")
left_w = 7.6
text("Working Capital Unlock", 0.3, 1.55, 24, C_BLACK, bold=True, font_base="Times")
text("$6,000 one off", 0.3, 1.95, 15, C_AMBER_D, bold=True)
text_wrapped("A four week project mapping cash trapped in debtors, inventory, work in progress, supplier terms and bank facilities across Orbitkey's global operation, with a prioritised action list to release cash back into the business.",
             0.3, 2.3, left_w, 11.5, C_TXT1, leading_mult=1.2)

rect(0.3, 3.3, left_w, 1.55, C_PANEL, stroke=C_AMBER_D)
text("Step one, answer a few quick questions", 0.5, 3.65, 13, C_BLACK, bold=True)
text("See the solutions matched to your size and industry.", 0.5, 3.95, 11, C_TXT2)
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 12)
c.drawString(0.5 * IN, y_(4.4), "profit-pulse.com.au/services/find-your-fit")
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          (0.5 * IN, y_(4.4) - 3, 0.5 * IN + c.stringWidth("profit-pulse.com.au/services/find-your-fit", "Helvetica-Bold", 12), y_(4.4) + 12))

rect(0.3, 5.05, left_w, 0.55, C_TEAL)
c.setFillColor(C_WHITE)
c.setFont("Helvetica-Bold", 12)
cta_txt = "Purchase the suggested product now to get started"
tw = c.stringWidth(cta_txt, "Helvetica-Bold", 12)
cta_x = 0.3 * IN + (left_w * IN - tw) / 2
c.drawString(cta_x, y_(5.38), cta_txt)
c.linkURL("https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z",
          (0.3 * IN, y_(5.05) - 0.55 * IN, 0.3 * IN + left_w * IN, y_(5.05)))

text("Prefer a conversation first? Book a complimentary discovery call.", 0.3, 5.95, 10.5, C_TXT2)
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 10.5)
c.drawString(0.3 * IN, y_(6.25), "Book a complimentary discovery call")
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (0.3 * IN, y_(6.25) - 3, 0.3 * IN + c.stringWidth("Book a complimentary discovery call", "Helvetica-Bold", 10.5), y_(6.25) + 12))

rx = 8.1
rw = 4.93
rect(rx, 1.2, rw, 5.55, C_BLACK)
text("Nitesh Roopa", rx + 0.25, 1.7, 17, C_AMBER_B, bold=True, font_base="Times")
text("CA, Managing Partner, ProfitPulse", rx + 0.25, 2.1, 11.5, C_OFF_WHITE)

pts = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
y = 2.7
for p in pts:
    rect(rx + 0.25, y + 0.02, 0.06, 0.06, C_TEAL)
    text_wrapped(p, rx + 0.45, y + 0.09, rw - 0.7, 10.5, C_WHITE, leading_mult=1.15)
    y += 0.5

rect(rx + 0.25, 4.7, rw - 0.5, 1.5 / IN, C_TEAL)
contact = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
y = 5.0
for line in contact:
    text(line, rx + 0.25, y, 11, C_OFF_WHITE)
    y += 0.4

c.showPage()
c.save()
print("PDF saved")
