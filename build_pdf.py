"""
ProfitPulse Brief PDF Builder
Target: Paire | Date: 22 Jul 2026
Mirrors build_brief.py (PPTX) layout, house style v3.3. Brand colours only. Zero dashes.
Page size: 960pt x 540pt (13.333in x 7.5in at 72dpi, matches PPTX exactly).
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, Color

IN = 72.0  # points per inch

W = 13.333 * IN
H = 7.5 * IN

C_BLACK     = HexColor("#000000")
C_TEAL      = HexColor("#01A296")
C_AMBER_B   = HexColor("#F8C806")
C_AMBER_D   = HexColor("#F6A102")
C_GOLD      = HexColor("#E3A712")
C_WHITE     = HexColor("#FFFFFF")
C_OFF_WHITE = HexColor("#E6E5DE")

COMPANY = "Paire"
DATE_STR = "22 Jul 2026"

FONT = "Helvetica"
FONT_B = "Helvetica-Bold"
FONT_I = "Helvetica-Oblique"
SERIF = "Times-Roman"
SERIF_B = "Times-Bold"

out_path = "/home/user/Daily-Outreach/Out-reach efforts/Brief_Paire_22Jul2026.pdf"
c = canvas.Canvas(out_path, pagesize=(W, H))
c.setTitle("Paire | ProfitPulse Brief | 22 Jul 2026")
c.setAuthor("ProfitPulse")
c.setSubject("Strategic Growth Diagnostic")


def rl_y(screen_y):
    return H - screen_y


def fill_rect(x, y_top, w, h, colour):
    c.setFillColor(colour)
    c.rect(x, rl_y(y_top + h), w, h, fill=1, stroke=0)


def stroke_rect(x, y_top, w, h, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    c.setFillColor(Color(0, 0, 0, alpha=0))
    c.rect(x, rl_y(y_top + h), w, h, fill=0, stroke=1)


def hline(x, y_top, w, colour, lw=1):
    c.setStrokeColor(colour)
    c.setLineWidth(lw)
    ry = rl_y(y_top)
    c.line(x, ry, x + w, ry)


def txt(s, x, y_top, size, colour, font=FONT, align="left", max_w=None):
    c.setFillColor(colour)
    c.setFont(font, size)
    baseline = rl_y(y_top + size * 0.92)
    if align == "right" and max_w:
        tw = c.stringWidth(s, font, size)
        x = x + max_w - tw
    c.drawString(x, baseline, s)


def txt_wrapped(s, x, y_top, max_w, size, colour, font=FONT, leading=None):
    if leading is None:
        leading = size * 1.3
    c.setFillColor(colour)
    c.setFont(font, size)
    words = s.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if c.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    y = y_top
    for line in lines:
        c.drawString(x, rl_y(y + size * 0.92), line)
        y += leading
    return y, len(lines)


def chrome(eyebrow, right_label):
    fill_rect(0, 0, W, H, C_WHITE)
    fill_rect(0, 0, 0.1 * IN, H, C_AMBER_D)
    fill_rect(0.1 * IN, 0, W - 0.1 * IN, 1.0 * IN, C_BLACK)
    txt(eyebrow, 0.35 * IN, 0.42 * IN, 12, C_OFF_WHITE, font=FONT_B)
    txt(right_label, 9.0 * IN, 0.42 * IN, 12, C_AMBER_B, font=FONT_B, align="right", max_w=4.1 * IN)
    hline(0.35 * IN, 7.03 * IN, 12.6 * IN, C_OFF_WHITE, 0.75)
    txt("Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au",
        0.35 * IN, 7.22 * IN, 8, C_BLACK, font=FONT)
    txt(DATE_STR, 10.0 * IN, 7.22 * IN, 8, C_BLACK, font=FONT, align="right", max_w=2.95 * IN)


def stat_card(x, y_top, w, h, number, label_lines, source):
    fill_rect(x, y_top, w, h, C_BLACK)
    fill_rect(x, y_top, w, 4, C_TEAL)
    txt(number, x + 0.14 * IN, y_top + 0.14 * IN, 28, C_AMBER_B, font=SERIF_B)
    ly = y_top + 0.7 * IN
    for line in label_lines:
        txt(line, x + 0.14 * IN, ly, 10.5, C_OFF_WHITE, font=FONT)
        ly += 13.5
    txt(source, x + 0.14 * IN, y_top + h - 0.28 * IN, 7.5, C_OFF_WHITE, font=FONT_I)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1
# ════════════════════════════════════════════════════════════════════════════
chrome("COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE")

txt(COMPANY, 0.35 * IN, 1.28 * IN, 40, C_BLACK, font=SERIF_B)
txt("Direct to consumer apparel brand, South Melbourne VIC",
    0.35 * IN, 1.98 * IN, 12, C_BLACK, font=FONT)

card_top = 2.32 * IN
card_h = 1.6 * IN
card_w = 1.98 * IN
gap = 0.15 * IN
x0 = 0.35 * IN
cards = [
    ("$9.4M", ["FY25 revenue,", "Smart50 rank"], "SmartCompany, Nov 2025"),
    ("64%",   ["FY25 revenue", "growth"],         "SmartCompany, Nov 2025"),
    ("88%",   ["FY24 revenue", "growth"],          "SmartCompany, 2024"),
    ("#15",   ["2025 Smart50", "national rank"],   "SmartCompany, Nov 2025"),
    ("2021",  ["Year founded,", "per Smart50"],    "SmartCompany, 2023"),
    ("16",    ["Full time staff,", "FY25"],         "SmartCompany, Nov 2025"),
]
for i, (num, lbl, src) in enumerate(cards):
    stat_card(x0 + i * (card_w + gap), card_top, card_w, card_h, num, lbl, src)

txt("KEY COMMERCIAL SIGNALS", 0.35 * IN, 4.42 * IN, 12, C_TEAL, font=FONT_B)

signals = [
    "Ranked 15th nationally on the 2025 Smart50 fastest growing companies list. Source: SmartCompany, Nov 2025",
    "Won the Smart50 2024 Rising Star Award and the Smart50 2024 Retail Award. Source: SmartCompany, 2024",
    "Opened first bricks and mortar flagship store at QV Melbourne, January 2025. Source: Inside Retail, Jan 2025",
    "Now stocked through wholesale partners in Singapore, 12 stores, and Malaysia, 3 stores. Source: Business News Australia",
    "Featured on Shark Tank Australia, national television exposure for the brand. Source: SmartCompany",
    "Revenue grew from about $1 million in 2021 to $9.4 million in FY25. Source: SmartCompany Smart50 profiles",
]
sy = 4.78 * IN
for s in signals:
    ny, nlines = txt_wrapped("•  " + s, 0.35 * IN, sy, 12.6 * IN, 11.5, C_BLACK, font=FONT, leading=15.5)
    sy = ny + 5.5

c.showPage()

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2
# ════════════════════════════════════════════════════════════════════════════
chrome("THE OPPORTUNITY", COMPANY.upper())
txt("Three commercial observations from ProfitPulse",
    0.35 * IN, 1.38 * IN, 14, C_BLACK, font=FONT_I)

col_top = 1.62 * IN
col_h = 4.9 * IN
col_w = 4.31 * IN
col_x0 = 0.35 * IN

observations = [
    (C_TEAL, C_BLACK, C_BLACK, "01", "Three growth fronts, one balance sheet",
     "In under two years Paire has added two Melbourne retail stores to its "
     "online business and opened its first international wholesale accounts "
     "in Singapore and Malaysia. Each channel carries its own cost structure: "
     "retail fit out and staffing, wholesale margin and freight, online "
     "customer acquisition. Without a channel by channel view, it is easy to "
     "keep funding the loudest growth story rather than the most profitable one."),
    (C_BLACK, C_TEAL, C_OFF_WHITE, "02", "Inventory is the quiet cash risk",
     "Apparel businesses scaling into new stores and new export markets at "
     "the same time typically see cash tied up in stock rise faster than "
     "revenue, since every new store and every new wholesale partner needs "
     "its own opening stock position. A Working Capital Unlock exercise "
     "would show exactly how much cash sits in stock today and where it can "
     "be released without starving the next store opening."),
    (C_GOLD, C_BLACK, C_BLACK, "03", "Nine times revenue growth deserves a costed plan",
     "Revenue has grown roughly ninefold since 2021, reaching $9.4 million "
     "in FY25 on 64 percent growth, following 88 percent growth the year "
     "before. That is the profile of a business ready for a formal twelve "
     "month growth plan, ranking retail expansion, wholesale scaling, and "
     "new product lines by expected return for the team and for any future "
     "funding partner."),
]

for i, (fill, numc, textc, idx, header, body) in enumerate(observations):
    left = col_x0 + i * col_w
    fill_rect(left, col_top, col_w, col_h, fill)
    txt(idx, left + 0.22 * IN, col_top + 0.5 * IN, 34, numc, font=SERIF_B)
    hy, _ = txt_wrapped(header, left + 0.22 * IN, col_top + 0.98 * IN, col_w - 0.44 * IN, 15, textc, font=SERIF_B, leading=18)
    txt_wrapped(body, left + 0.22 * IN, col_top + 1.86 * IN, col_w - 0.44 * IN, 11, textc, font=FONT, leading=15.2)

txt_wrapped(("These observations are offered in good faith. Paire has built something "
             "genuinely impressive in four years. The question is simply whether the "
             "operating and capital plan is keeping pace with three simultaneous growth fronts."),
            0.35 * IN, 6.66 * IN, 12.6 * IN, 10.5, C_BLACK, font=FONT_I, leading=13.5)

c.showPage()

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3
# ════════════════════════════════════════════════════════════════════════════
chrome("THE RECOMMENDATION", COMPANY.upper())

left_x = 0.35 * IN
left_w = 7.35 * IN
right_x = 7.95 * IN
right_w = 5.0 * IN

txt("Strategic Growth Diagnostic", left_x, 1.5 * IN, 24, C_BLACK, font=SERIF_B)
txt("$5,000 one off, ProfitPulse verified price", left_x, 1.9 * IN, 13, C_TEAL, font=FONT_B)

txt_wrapped(("Maps revenue, capacity, and margin headroom across retail, wholesale, "
             "and online, then produces a twelve month growth plan with funding and "
             "capital allocation steps for each channel."),
            left_x, 2.2 * IN, left_w, 11.5, C_BLACK, font=FONT, leading=15)

stroke_rect(left_x, 2.95 * IN, left_w, 1.35 * IN, C_TEAL, 1)
fill_rect(left_x, 2.95 * IN, left_w, 1.35 * IN, C_OFF_WHITE)
stroke_rect(left_x, 2.95 * IN, left_w, 1.35 * IN, C_TEAL, 1)
txt("Step one, answer a few quick questions", left_x + 0.2 * IN, 3.2 * IN, 12.5, C_BLACK, font=FONT_B)
txt("See the solutions matched to your size and industry.", left_x + 0.2 * IN, 3.54 * IN, 11, C_BLACK, font=FONT)
c.setFillColor(C_TEAL)
c.linkURL("https://profit-pulse.com.au/services/find-your-fit/",
          (left_x + 0.2 * IN, rl_y(4.0 * IN) - 4, left_x + 4.0 * IN, rl_y(3.86 * IN)), relative=0)
txt("profit-pulse.com.au/services/find-your-fit", left_x + 0.2 * IN, 3.88 * IN, 12.5, C_TEAL, font=FONT_B)

fill_rect(left_x, 4.5 * IN, 5.1 * IN, 0.55 * IN, C_AMBER_D)
txt("Purchase the suggested product now to get started", left_x + 0.18 * IN, 4.68 * IN, 12, C_BLACK, font=FONT_B)
c.linkURL("https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h",
          (left_x, rl_y(5.05 * IN), left_x + 5.1 * IN, rl_y(4.5 * IN)), relative=0)

txt("Prefer a conversation first?", left_x, 5.42 * IN, 11.5, C_BLACK, font=FONT)
txt("Book a complimentary discovery call", left_x, 5.74 * IN, 11.5, C_TEAL, font=FONT_B)
c.linkURL("https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true",
          (left_x, rl_y(5.9 * IN), left_x + 3.5 * IN, rl_y(5.6 * IN)), relative=0)

txt_wrapped("Wedge: rapid multi channel expansion without a costed capital plan across retail, wholesale, and online.",
            left_x, 6.25 * IN, left_w, 9.5, C_BLACK, font=FONT_I, leading=12.5)

# Right column panel
fill_rect(right_x, 1.15 * IN, right_w, 5.55 * IN, C_BLACK)
txt("Nitesh Roopa", right_x + 0.22 * IN, 1.6 * IN, 18, C_AMBER_B, font=SERIF_B)
txt("CA, Managing Partner, ProfitPulse", right_x + 0.22 * IN, 1.95 * IN, 11.5, C_OFF_WHITE, font=FONT)
hline(right_x + 0.22 * IN, 2.12 * IN, right_w - 0.44 * IN, C_TEAL, 1)

cred_lines = [
    "16 years across 4 countries",
    "52 deals executed and managed",
    "Largest deal USD 1.3 billion, Cahora Bassa Hydro",
    "Queensland GRBT project value over AUD 10 billion",
]
cy = 2.42 * IN
for cl in cred_lines:
    txt("•  " + cl, right_x + 0.22 * IN, cy, 10.5, C_OFF_WHITE, font=FONT)
    cy += 17.5

hline(right_x + 0.22 * IN, 3.85 * IN, right_w - 0.44 * IN, C_TEAL, 1)

contact_lines = [
    ("Profit-Pulse.com.au", C_OFF_WHITE),
    ("Nitesh@Profit-Pulse.com.au", C_TEAL),
    ("+61 411 876 267", C_OFF_WHITE),
    ("linkedin.com/in/nitesh-roopa-77594163", C_TEAL),
]
cy = 4.18 * IN
for ct, col in contact_lines:
    txt(ct, right_x + 0.22 * IN, cy, 10.5, col, font=FONT)
    cy += 18.5

txt("Brisbane, Australia", right_x + 0.22 * IN, 6.35 * IN, 10, C_OFF_WHITE, font=FONT)

c.showPage()
c.save()
print(f"PDF saved: {out_path}")
