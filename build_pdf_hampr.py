#!/usr/bin/env python3
"""Build Brief_Hampr_09Jun2026.pdf using fpdf2 — mirrors the PPTX three-slide deck."""

from fpdf import FPDF
import os

OUT  = "/home/user/Daily-Outreach/Out-reach efforts"
DATE = "09 Jun 2026"

# Brand colours as R,G,B tuples
BLK  = (0,   0,   0  )
TEAL = (1,   162, 150)
AMB  = (248, 200, 6  )
AMBD = (246, 161, 2  )
GOLD = (227, 167, 18 )
WHT  = (255, 255, 255)
OWHT = (230, 229, 222)
LGRY = (245, 245, 245)

# Slide canvas: 254mm x 142.875mm (10in x 5.625in at 25.4mm/in)
W = 254
H = 142.875

STRIPE_URL = "https://buy.stripe.com/dRmbJ22uA8ekdlCgYT3ks0r"
QUEST_URL  = "https://profit-pulse.com.au/full-suite-of-products"
BOOKING_URL= ("https://bookings.cloud.microsoft/book/"
              "ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")


class Brief(FPDF):
    def __init__(self):
        super().__init__(orientation="L", unit="mm", format=(H, W))
        self.set_auto_page_break(False)
        self.set_margins(0, 0, 0)

    # ── Primitives ────────────────────────────────────────────────────────────

    def fill(self, x, y, w, h, rgb):
        self.set_fill_color(*rgb)
        self.rect(x, y, w, h, style="F")

    def put_text(self, x, y, w, h, text, sz=10, bold=False, italic=False,
                 clr=BLK, align="L", font="Helvetica", multi=False):
        if bold and italic:
            style = "BI"
        elif bold:
            style = "B"
        elif italic:
            style = "I"
        else:
            style = ""
        self.set_xy(x, y)
        self.set_font(font, style, sz)
        self.set_text_color(*clr)
        if multi:
            self.multi_cell(w, h * 0.38, text, align=align, ln=1)
        else:
            self.cell(w, h, text, align=align)

    def put_multi(self, x, y, w, line_h, text, sz=10, bold=False, italic=False,
                  clr=BLK, align="L"):
        style = ("B" if bold else "") + ("I" if italic else "")
        self.set_xy(x, y)
        self.set_font("Helvetica", style, sz)
        self.set_text_color(*clr)
        self.multi_cell(w, line_h, text, align=align)

    # ── Common frame ──────────────────────────────────────────────────────────

    def frame(self, eyebrow):
        # Amber left stripe
        self.fill(0, 0, 2.0, H, AMB)
        # Black header band
        self.fill(2.0, 0, W - 2.0, 21.6, BLK)
        # Eyebrow
        self.put_text(4, 2.5, 120, 8, eyebrow, sz=9, bold=True, clr=OWHT)
        # Right label
        self.put_text(180, 2.5, 70, 8, "PROFITPULSE", sz=9, bold=True, clr=OWHT, align="R")
        # Footer rule
        self.fill(2.0, H - 3.8, W - 2.0, 0.35, OWHT)
        footer = (f"Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  "
                  f"Profit-Pulse.com.au  |  {DATE}")
        self.put_text(4, H - 3.2, W - 6, 3.0, footer, sz=6.5, italic=True, clr=BLK)

    # ── Stat card ─────────────────────────────────────────────────────────────

    def stat_card(self, x, y, w, h, num, lab1, lab2, src):
        self.fill(x, y, w, h, BLK)
        self.fill(x, y, w, 1.1, TEAL)
        # Number
        self.set_xy(x + 1.2, y + 1.5)
        self.set_font("Helvetica", "B", 17)
        self.set_text_color(*AMB)
        self.cell(w - 2, 7, num)
        # Labels
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*WHT)
        self.set_xy(x + 1.2, y + h * 0.52)
        self.cell(w - 2, 4, lab1)
        self.set_text_color(*OWHT)
        self.set_xy(x + 1.2, y + h * 0.67)
        self.cell(w - 2, 4, lab2)
        # Source
        self.set_font("Helvetica", "I", 5.5)
        self.set_text_color(*TEAL)
        self.set_xy(x + 1.2, y + h * 0.83)
        self.cell(w - 2, 3.5, src)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Commercial Intelligence Brief
# ═══════════════════════════════════════════════════════════════════════════════

pdf = Brief()
pdf.add_page()
pdf.frame("COMMERCIAL INTELLIGENCE BRIEF")

# Company name
pdf.set_xy(4, 22.5)
pdf.set_font("Helvetica", "B", 28)
pdf.set_text_color(*BLK)
pdf.cell(170, 12, "Hampr")

# Descriptor
pdf.put_text(4, 34, 248, 5, "B2B workplace hospitality and corporate food technology marketplace   |   Sydney, NSW", sz=9, clr=BLK)

# 6 stat cards  y=40, h=27
CW = 39.6; CH = 27; CY = 40; GAP = 1.65
XS = [4 + i * (CW + GAP) for i in range(6)]

CARDS = [
    ("AUD $12.7M",  "Revenue",               "FY2024",               "Smart50 2024, SmartCompany Nov 2024"),
    ("153%",        "Revenue growth",         "year on year FY2024",  "Smart50 2024, SmartCompany Nov 2024"),
    ("#8",          "Smart50 ranking",        "nationally, 2024",     "SmartCompany Smart50 Nov 2024"),
    ("37",          "Employees",              "full time equivalents","Smart50 2024, SmartCompany Nov 2024"),
    ("2018",        "Year founded",           "Sydney, NSW",          "Smart50 2024 profile"),
    ("$11.1M",      "Total funding raised",   "across 4 rounds",      "Crunchbase / Tracxn, May 2025"),
]
for i, (num, l1, l2, src) in enumerate(CARDS):
    pdf.stat_card(XS[i], CY, CW, CH, num, l1, l2, src)

# Key commercial signals
pdf.put_text(4, 70, 80, 5, "KEY COMMERCIAL SIGNALS", sz=8.5, bold=True, clr=TEAL)

SIGNALS = [
    "Singapore market entry FY2026: hampr.sg website live, international expansion stated publicly.  Source: Smart50 2024 profile, Nov 2024",
    "Two major enterprise tenders targeted within 12 months of Smart50 submission.  Source: Smart50 2024 profile, Nov 2024",
    "Enterprise client publicly named: Toyota Finance Australia referenced in company media.  Source: Momentum91 podcast, May 2025",
    "Investor backing includes Blackbird Ventures, Techstars, Startmate, Lee Kim Tah Group.  Source: Crunchbase / Tracxn, 2025",
    "National operations active across Sydney, Melbourne, Perth, Brisbane, and Canberra.  Source: Hampr website, 2025",
    "No publicly named CFO or Head of Finance in any company profile or media.  Source: Multiple profiles, 2025",
]
SY = 75
for sig in SIGNALS:
    pdf.put_multi(4, SY, 248, 4.5, sig, sz=8, clr=BLK)
    SY += 9


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — The Opportunity
# ═══════════════════════════════════════════════════════════════════════════════

pdf.add_page()
pdf.frame("THE OPPORTUNITY")

pdf.put_text(4, 22.5, 248, 7, "Hampr   |   Three commercial observations from ProfitPulse",
             sz=11, bold=True, clr=BLK)

COLS = [
    (4,    84, TEAL, WHT, WHT, BLK),
    (89,   84, BLK,  WHT, WHT, AMB),
    (174,  80, GOLD, BLK, BLK, BLK),
]
IDXS  = ["01", "02", "03"]
HEADS = [
    "The marketplace float is a cash flow event waiting to happen",
    "The Singapore launch needs a financial model, not just a website",
    "Enterprise contracts will change the financial DNA of the business",
]
BODIES = [
    ("At $12.7M in revenue growing at 153% year on year, Hampr processes payments between "
     "corporate clients and a network of food and event suppliers. As transaction volume grows "
     "so does the cash in transit between client receipts and supplier payments. Without a "
     "structured 13 week cash flow forecast a fast growing marketplace can run short of cash "
     "even while reporting a healthy profit. ProfitPulse fractional CFO engagements routinely "
     "identify six to twelve percent of annual revenue in manageable float."),

    ("Hampr has moved into Singapore with a live website and a publicly stated FY2026 target. "
     "International expansion requires upfront investment in team, regulatory compliance, and "
     "market development well ahead of local revenue. With the most recent funding round at "
     "$1.08M in May 2025, a rigorous financial model for the Singapore launch is needed: one "
     "that maps cash burn, the revenue ramp, and the break even timeline. A Fractional CFO "
     "Partnership is designed to carry exactly this work."),

    ("Two enterprise tenders in active pursuit means a step change from a diversified SME "
     "client base to a smaller number of much larger contracts. Enterprise clients carry 60 to "
     "90 day payment terms and detailed financial reporting expectations. This concentrates "
     "revenue risk, extends the cash conversion cycle, and demands board grade reporting that "
     "a founder led business at this scale may not yet have in place. Getting ahead of this "
     "transition now protects both margin and negotiating position."),
]

CT = 30; CH2 = 100
for (cx, cw, bg, tc, hc, nc), idx, head, body in zip(COLS, IDXS, HEADS, BODIES):
    pdf.fill(cx, CT, cw, CH2, bg)
    # Index
    pdf.set_xy(cx + 2.5, CT + 2.5)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*nc)
    pdf.cell(cw - 5, 10, idx)
    # Heading
    pdf.set_xy(cx + 2.5, CT + 15)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(*hc)
    pdf.multi_cell(cw - 5, 4.5, head)
    # Body
    pdf.set_xy(cx + 2.5, CT + 33)
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(*tc)
    pdf.multi_cell(cw - 5, 4, body)

# Warm close
pdf.put_multi(4, 134, 248, 4.5,
    ("These observations are offered in good faith. Hampr has built something genuinely "
     "impressive. The question is simply whether the financial architecture is ready for the next chapter."),
    sz=7.5, italic=True, clr=BLK)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — The Recommendation
# ═══════════════════════════════════════════════════════════════════════════════

pdf.add_page()
pdf.frame("THE RECOMMENDATION")

# ── Left column ───────────────────────────────────────────────────────────────
LX = 4; LW = 140

pdf.set_xy(LX, 23)
pdf.set_font("Helvetica", "B", 19)
pdf.set_text_color(*BLK)
pdf.cell(LW, 10, "Fractional CFO Partnership")

pdf.put_text(LX, 34, LW, 7, "$7,500 per month", sz=12, bold=True, clr=TEAL)

pdf.put_multi(LX, 42, LW, 4.2,
    ("A senior financial partner at the table on a monthly cadence. Monthly management pack, "
     "quarterly board grade review, ad hoc decision support, and a single annual deep dive. "
     "For Hampr: a financial rhythm for the Singapore launch, enterprise contract cash flow "
     "modelling, and the reporting infrastructure to match the ambition."),
    sz=9, clr=BLK)

# Step one block
pdf.fill(LX, 63, 1.0, 35, TEAL)
pdf.put_text(LX + 3, 64, LW - 3, 5.5, "Step one: answer a few quick questions", sz=9.5, bold=True, clr=BLK)
pdf.put_text(LX + 3, 70, LW - 3, 5, "See the solutions matched to your size and industry.", sz=9, clr=BLK)
# Link text (clean, no tracking tags on brief)
pdf.set_xy(LX + 3, 76)
pdf.set_font("Helvetica", "B", 9)
pdf.set_text_color(*TEAL)
pdf.cell(LW - 3, 5, "profit-pulse.com.au/full-suite-of-products",
         link=QUEST_URL)

# Direct purchase CTA
pdf.fill(LX, 85, 1.0, 13, AMBD)
pdf.put_text(LX + 3, 86, LW - 3, 5, "Already know this is the priority?", sz=9.5, bold=True, clr=BLK)
pdf.set_xy(LX + 3, 92)
pdf.set_font("Helvetica", "B", 9)
pdf.set_text_color(*AMBD)
pdf.cell(LW - 3, 5, "Purchase the suggested product now to get started",
         link=STRIPE_URL)

# Discovery call
pdf.fill(LX, 101, 1.0, 13, AMB)
pdf.put_text(LX + 3, 102, LW - 3, 5, "Prefer a conversation first?", sz=9.5, bold=True, clr=BLK)
pdf.set_xy(LX + 3, 108)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(*TEAL)
pdf.cell(LW - 3, 5, "Book a complimentary discovery call with Nitesh Roopa",
         link=BOOKING_URL)

# ── Vertical divider ─────────────────────────────────────────────────────────
pdf.fill(147, 22, 0.75, H - 28, TEAL)

# ── Right column — credibility ────────────────────────────────────────────────
RX = 150; RW = 100
pdf.fill(RX, 22, RW + 2, H - 28, BLK)

pdf.set_xy(RX + 3, 24)
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(*AMB)
pdf.cell(RW - 3, 9, "Nitesh Roopa")

pdf.put_text(RX + 3, 34, RW - 3, 5.5, "CA, Managing Partner   |   ProfitPulse", sz=9.5, clr=OWHT)
pdf.fill(RX + 3, 41, RW - 10, 0.5, TEAL)

CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique Government, Hydro",
    "Total GRBT project value in Queensland: over AUD 10 billion",
]
CY3 = 43
for c in CREDS:
    pdf.put_multi(RX + 3, CY3, RW - 5, 4, c, sz=8, clr=OWHT)
    CY3 += 8

pdf.fill(RX + 3, CY3 + 1, RW - 10, 0.5, TEAL)
CY3 += 4

CONTACTS = [
    "Profit-Pulse.com.au",
    "Nitesh@Profit-Pulse.com.au",
    "+61 411 876 267",
    "linkedin.com/in/nitesh-roopa-77594163",
]
for ctxt in CONTACTS:
    pdf.put_text(RX + 3, CY3, RW - 5, 5.5, ctxt, sz=8, clr=TEAL)
    CY3 += 6.5

# ── Save ─────────────────────────────────────────────────────────────────────
out_path = os.path.join(OUT, "Brief_Hampr_09Jun2026.pdf")
pdf.output(out_path)
print(f"Saved: {out_path}")
