"""
ProfitPulse Brief PDF Builder
Target: Paire | Date: 04 Aug 2026
Direct PDF generation, 3 slides, matches build_brief.py geometry and content exactly.
Page size: 960pt x 540pt (widescreen, matches PPTX 13.333in x 7.5in at 72dpi)
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

W = 960.0
H = 540.0
PT_PER_IN = 72.0
DATE_STR = "04 Aug 2026"

C_BLACK    = HexColor("#000000")
C_TEAL     = HexColor("#01A296")
C_AMBER_B  = HexColor("#F8C806")
C_AMBER_D  = HexColor("#F6A102")
C_GOLD     = HexColor("#E3A712")
C_WHITE    = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")
C_GREY_LINE = HexColor("#CCCCCC")
C_GREY_MED  = HexColor("#666666")
C_GREY_DK   = HexColor("#333333")
C_GREY_SRC  = HexColor("#9A9A9A")
C_PANEL_BG  = HexColor("#F2F2F0")


def IN(v):
    return v * PT_PER_IN


def rl_y(top_pt):
    """Convert a top-origin y (points from top) to ReportLab bottom-origin y."""
    return H - top_pt


def fill_rect(c, x, top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, rl_y(top + h), w, h, fill=1, stroke=0)


def stroke_rect(c, x, top, w, h, colour, width=1):
    c.setStrokeColor(colour)
    c.setLineWidth(width)
    c.rect(x, rl_y(top + h), w, h, fill=0, stroke=1)


def hline(c, x, top, w, colour, width=1):
    c.setStrokeColor(colour)
    c.setLineWidth(width)
    y = rl_y(top)
    c.line(x, y, x + w, y)


def text(c, s, x, top, size, colour, font="Helvetica", align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl_y(top + size * 0.9)
    if align == "right" and max_w:
        tw = c.stringWidth(s, font, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(s, font, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, baseline, s)


def wrap_lines(c, s, font, size, max_w):
    words = s.split()
    lines = []
    cur = ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def text_block(c, s, x, top, max_w, size, colour, font="Helvetica", leading=None, align="left"):
    leading = leading or size * 1.25
    lines = wrap_lines(c, s, font, size, max_w)
    y = top
    for ln in lines:
        text(c, ln, x, y, size, colour, font=font, align=align, max_w=max_w)
        y += leading
    return y


def chrome(c, eyebrow, right_label="PROFITPULSE"):
    fill_rect(c, 0, 0, IN(0.1), H, C_AMBER_D)
    fill_rect(c, IN(0.1), 0, W - IN(0.1), IN(1.0), C_BLACK)
    text(c, eyebrow, IN(0.35), IN(0.30), 12, C_OFF_WHITE, font="Helvetica-Bold")
    text(c, right_label, IN(9.0), IN(0.30), 12, C_AMBER_B, font="Helvetica-Bold",
         align="right", max_w=IN(4.0))
    hline(c, IN(0.35), IN(7.05), IN(12.6), C_GREY_LINE, width=1)
    text(c, "Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
         IN(0.35), IN(7.10), 8, C_GREY_MED)
    text(c, DATE_STR, IN(10.5), IN(7.10), 8, C_GREY_MED, align="right", max_w=IN(2.4))


c = canvas.Canvas("/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_04Aug2026.pdf",
                   pagesize=(W, H))

# ══════════════════════════════════════════════════════════════════════
# SLIDE 1
# ══════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
chrome(c, "COMMERCIAL INTELLIGENCE BRIEF")

text(c, "Paire", IN(0.35), IN(1.15), 40, C_BLACK, font="Helvetica-Bold")
text(c, "Direct to consumer apparel brand, South Melbourne VIC", IN(0.35), IN(1.90), 13, C_GREY_DK)

stat_cards = [
    ("$9.4M", "FY2025 revenue", "SmartCompany Smart50 2025"),
    ("64%", "FY2025 revenue growth", "SmartCompany Smart50 2025"),
    ("#15", "Smart50 2025 national rank", "SmartCompany, 2025"),
    ("2020", "Year founded", "FashionUnited 2026"),
    ("40", "People on the team", "FashionUnited 2026"),
]
card_w = IN(2.3)
card_h = IN(1.6)
card_gap = IN(0.15)
card_y = IN(2.4)
x = IN(0.35)
for num, label, source in stat_cards:
    fill_rect(c, x, card_y, card_w, card_h, C_BLACK)
    fill_rect(c, x, card_y, card_w, 4, C_TEAL)
    text(c, num, x + IN(0.15), card_y + IN(0.20), 28, C_AMBER_B, font="Helvetica-Bold")
    text_block(c, label, x + IN(0.15), card_y + IN(0.80), card_w - IN(0.3), 11, C_OFF_WHITE, leading=14)
    text(c, source, x + IN(0.15), card_y + IN(1.32), 7, C_GREY_SRC, font="Helvetica-Oblique")
    x += card_w + card_gap

text(c, "REVENUE TRAJECTORY, SMART50 AWARD CITATIONS", IN(0.35), IN(4.30), 11, C_TEAL, font="Helvetica-Bold")
chart_data = [("Smart50 2023", 2.15), ("Smart50 2024", 5.9), ("Smart50 2025", 9.4)]
max_val = 9.4
bar_w = IN(1.5)
bar_gap = IN(0.5)
base_top = IN(6.55)
max_bar_h = IN(1.5)
bx = IN(0.35)
for label, val in chart_data:
    bar_h = max_bar_h * (val / max_val)
    fill_rect(c, bx, base_top - bar_h, bar_w, bar_h, C_TEAL)
    text(c, f"${val}M", bx, base_top - bar_h - IN(0.32), 11, C_BLACK, font="Helvetica-Bold",
         align="center", max_w=bar_w)
    text(c, label, bx, base_top + IN(0.05), 8, C_GREY_MED, align="center", max_w=bar_w)
    bx += bar_w + bar_gap

text(c, "KEY COMMERCIAL SIGNALS", IN(6.85), IN(4.30), 11, C_TEAL, font="Helvetica-Bold")
signals = [
    "Smart50 2025 rank 15, revenue grew 64% to $9.4M (SmartCompany, 2025)",
    "Smart50 2024 double winner, Rising Star and Retail Award, 88% growth (SmartCompany)",
    "Pitched on Shark Tank Australia for a $500,000 equity investment (SmartCompany)",
    "Singapore distributor Seager Inc signed for Southeast Asia wholesale, 2023 (SmartCompany)",
    "Converted QV Melbourne pop up into a permanent flagship store (SmartCompany, 2025)",
    "Launched direct to consumer in the United States, 2026 range (FashionUnited)",
]
y = IN(4.65)
for s in signals:
    lines = wrap_lines(c, "• " + s, "Helvetica", 10.5, IN(6.1))
    for ln in lines:
        text(c, ln, IN(6.85), y, 10.5, C_BLACK)
        y += 13.2
    y += 4

c.showPage()

# ══════════════════════════════════════════════════════════════════════
# SLIDE 2
# ══════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
chrome(c, "THE OPPORTUNITY")
text(c, "Paire: three commercial observations from ProfitPulse", IN(0.35), IN(1.10), 17, C_BLACK, font="Helvetica-Bold")

col_y = IN(1.70)
col_h = IN(4.9)
col_w = IN(4.05)
col_gap = IN(0.09)
cols = [
    (C_TEAL, C_WHITE, "01", "Three growth engines, one capital plan",
     "Paire is scaling revenue at 64% a year while running three capital hungry expansions at once: "
     "a permanent flagship plus planned Sydney stores, a new Singapore wholesale relationship, and a "
     "2026 direct to consumer launch in the United States. Each draws on the same pool of cash and "
     "management time. No public evidence points to a single, costed plan that sequences these three "
     "bets against the capacity of the business to fund them."),
    (C_BLACK, C_WHITE, "02", "A funding question still open",
     "The Shark Tank Australia pitch shows the founders have already tested appetite for outside equity, "
     "reportedly seeking $500,000, but no source confirms a completed deal. Meanwhile headcount has grown "
     "from roughly ten to a stated forty. Whether the next stage of growth is funded from trading cash flow, "
     "debt, or equity is an open question that shapes every other decision this year."),
    (C_GOLD, C_BLACK, "03", "Structure has not caught up with scale",
     "Paire now sells across Australia, Singapore, and the United States, from a single Melbourne base whose "
     "exact corporate structure for the US market is not publicly disclosed. As revenue heads toward eight "
     "figures, the reporting rhythm and capital allocation discipline that suited a two million dollar startup "
     "may not suit a multi country retail and wholesale business."),
]
cx = IN(0.35)
for fill, txt_colour, idx, header, body in cols:
    fill_rect(c, cx, col_y, col_w, col_h, fill)
    text(c, idx, cx + IN(0.25), col_y + IN(0.25), 30, txt_colour, font="Helvetica-Bold")
    hy = col_y + IN(0.95)
    hy = text_block(c, header, cx + IN(0.25), hy, col_w - IN(0.5), 14, txt_colour, font="Helvetica-Bold", leading=17)
    by = col_y + IN(1.75)
    text_block(c, body, cx + IN(0.25), by, col_w - IN(0.5), 10.5, txt_colour, leading=13.2)
    cx += col_w + col_gap

text_block(c,
    "These are observations offered in good faith. Paire has built something impressive. "
    "The question is simply whether the financial architecture matches the ambition.",
    IN(0.35), IN(6.66), IN(12.6), 10.5, C_GREY_DK, font="Helvetica-Oblique", leading=13)

c.showPage()

# ══════════════════════════════════════════════════════════════════════
# SLIDE 3
# ══════════════════════════════════════════════════════════════════════
fill_rect(c, 0, 0, W, H, C_WHITE)
chrome(c, "THE RECOMMENDATION")

lx = IN(0.35)
lw = IN(7.4)
text(c, "Strategic Growth Diagnostic", lx, IN(1.15), 22, C_BLACK, font="Helvetica-Bold")
text(c, "$5,000 one off", lx, IN(1.68), 15, C_TEAL, font="Helvetica-Bold")
text_block(c, "A six week engagement that maps revenue, capacity, and margin headroom across the "
              "retail, wholesale, and United States channels, then produces a 12 month growth plan "
              "with funding and capital allocation steps spelled out for each.",
           lx, IN(2.10), lw, 11.5, C_BLACK, leading=14.5)

stroke_rect(c, lx, IN(3.25), lw, IN(1.35), C_TEAL, width=1)
fill_rect(c, lx, IN(3.25), lw, IN(1.35), C_PANEL_BG)
stroke_rect(c, lx, IN(3.25), lw, IN(1.35), C_TEAL, width=1)
text(c, "Step one, answer a few quick questions", lx + IN(0.2), IN(3.35), 12.5, C_TEAL, font="Helvetica-Bold")
text(c, "See the solutions matched to your size and industry.", lx + IN(0.2), IN(3.70), 10.5, C_BLACK)
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          (lx + IN(0.2), rl_y(IN(4.00) + 16), lx + IN(0.2) + IN(4.5), rl_y(IN(4.00))),
          relative=0)
text(c, "profit-pulse.com.au/services/find-your-fit", lx + IN(0.2), IN(4.00), 11.5, C_TEAL, font="Helvetica-Bold")

fill_rect(c, lx, IN(4.80), lw, IN(0.65), C_AMBER_B)
c.linkURL("https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
          (lx, rl_y(IN(4.80) + IN(0.65)), lx + lw, rl_y(IN(4.80))), relative=0)
text(c, "Purchase the suggested product now to get started", lx, IN(4.97), 13, C_BLACK,
     font="Helvetica-Bold", align="center", max_w=lw)

text(c, "Prefer a conversation first? Book a complimentary discovery call.", lx, IN(5.65), 10.5, C_BLACK)
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (lx, rl_y(IN(5.98) + 16), lx + IN(4.5), rl_y(IN(5.98))), relative=0)
text(c, "bookings.cloud.microsoft/book/ProfitPulse1", lx, IN(5.98), 10.5, C_TEAL, font="Helvetica-Bold")

rx = IN(8.05)
rw = IN(4.9)
fill_rect(c, rx, IN(1.15), rw, IN(5.75), C_BLACK)
fill_rect(c, rx, IN(1.15), rw, 4, C_TEAL)
text(c, "Nitesh Roopa", rx + IN(0.25), IN(1.40), 18, C_AMBER_B, font="Helvetica-Bold")
text(c, "CA, Managing Partner, ProfitPulse", rx + IN(0.25), IN(1.85), 11.5, C_OFF_WHITE)
hline(c, rx + IN(0.25), IN(2.30), rw - IN(0.5), C_TEAL, width=1)

credibility = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest single deal, USD 1.3 billion, Cahora Bassa",
    "Total GRBT project value over AUD 10 billion",
]
y = IN(2.65)
for item in credibility:
    text(c, item, rx + IN(0.25), y, 11, C_OFF_WHITE)
    y += 34

hline(c, rx + IN(0.25), y + IN(0.05), rw - IN(0.5), C_TEAL, width=1)
y += IN(0.35)
contact_lines = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for item in contact_lines:
    text(c, item, rx + IN(0.25), y, 11, C_WHITE)
    y += 34

c.showPage()
c.save()
print("PDF saved")
