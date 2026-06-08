#!/usr/bin/env python3
"""Build Brief_Hampr_09Jun2026.pdf via reportlab — mirrors the three-slide PPTX deck."""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch, mm
from reportlab.lib.colors import HexColor, white, black
import os

OUT  = "/home/user/Daily-Outreach/Out-reach efforts"
DATE = "09 Jun 2026"

# Brand hex colours
C_BLK  = HexColor("#000000")
C_TEAL = HexColor("#01A296")
C_AMB  = HexColor("#F8C806")
C_AMBD = HexColor("#F6A102")
C_GOLD = HexColor("#E3A712")
C_WHT  = HexColor("#FFFFFF")
C_OWHT = HexColor("#E6E5DE")

# Slide: 10in x 5.625in  (landscape)
W = 10 * inch
H = 5.625 * inch

STRIPE_URL = "https://buy.stripe.com/dRmbJ22uA8ekdlCgYT3ks0r"
QUEST_URL  = "https://profit-pulse.com.au/full-suite-of-products"
BOOKING_URL= ("https://bookings.cloud.microsoft/book/"
              "ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true")

out_path = os.path.join(OUT, "Brief_Hampr_09Jun2026.pdf")
c = canvas.Canvas(out_path, pagesize=(W, H))


def in_(n): return n * inch


# ── Helpers ───────────────────────────────────────────────────────────────────

def filled_rect(x, y, w, h, clr):
    """y is from bottom in reportlab."""
    c.setFillColor(clr)
    c.rect(x, y, w, h, fill=1, stroke=0)


def text_at(x, y, txt, size=10, bold=False, italic=False, clr=C_BLK,
            align="left", max_w=None):
    """y from bottom. Single line."""
    c.setFillColor(clr)
    fname = "Helvetica-Bold" if bold else "Helvetica"
    if italic and bold:
        fname = "Helvetica-BoldOblique"
    elif italic:
        fname = "Helvetica-Oblique"
    c.setFont(fname, size)
    if align == "right" and max_w:
        tw = c.stringWidth(txt, fname, size)
        x = x + max_w - tw
    elif align == "center" and max_w:
        tw = c.stringWidth(txt, fname, size)
        x = x + (max_w - tw) / 2
    c.drawString(x, y, txt)


def wrap_text(x, y, txt, size, clr, max_w, line_h, bold=False, italic=False):
    """Wrap text into lines, drawing from y downward (returns final y)."""
    fname = "Helvetica-Bold" if bold else "Helvetica"
    if italic and bold:
        fname = "Helvetica-BoldOblique"
    elif italic:
        fname = "Helvetica-Oblique"
    c.setFont(fname, size)
    c.setFillColor(clr)
    words = txt.split()
    line = ""
    cur_y = y
    for word in words:
        test = (line + " " + word).strip()
        if c.stringWidth(test, fname, size) <= max_w:
            line = test
        else:
            if line:
                c.drawString(x, cur_y, line)
                cur_y -= line_h
            line = word
    if line:
        c.drawString(x, cur_y, line)
        cur_y -= line_h
    return cur_y


# ── Common frame ──────────────────────────────────────────────────────────────

def frame(eyebrow):
    # Amber left stripe
    filled_rect(0, 0, in_(0.08), H, C_AMB)
    # Black header band
    filled_rect(in_(0.08), H - in_(0.85), W - in_(0.08), in_(0.85), C_BLK)
    # Eyebrow
    text_at(in_(0.2), H - in_(0.37), eyebrow, size=9, bold=True, clr=C_OWHT)
    # Right label
    text_at(in_(0.2), H - in_(0.37), "PROFITPULSE", size=9, bold=True, clr=C_OWHT,
            align="right", max_w=W - in_(0.25))
    # Footer rule
    filled_rect(in_(0.08), in_(0.07), W - in_(0.08), in_(0.015), C_OWHT)
    footer = (f"Prepared by Nitesh Roopa CA, Managing Partner, ProfitPulse  |  "
              f"Profit-Pulse.com.au  |  {DATE}")
    text_at(in_(0.2), in_(0.09), footer, size=6.5, italic=True, clr=C_BLK)


# ── Stat card ─────────────────────────────────────────────────────────────────

def stat_card(x, y_top, w, h, num, lab1, lab2, src):
    """y_top is distance from top of slide (converted internally)."""
    y = H - y_top - h
    filled_rect(x, y, w, h, C_BLK)
    filled_rect(x, y + h - in_(0.045), w, in_(0.045), C_TEAL)
    text_at(x + in_(0.05), y + h - in_(0.11), num, size=18, bold=True, clr=C_AMB)
    text_at(x + in_(0.05), y + h * 0.48, lab1, size=7.5, clr=C_WHT)
    text_at(x + in_(0.05), y + h * 0.32, lab2, size=7.5, clr=C_OWHT)
    text_at(x + in_(0.05), y + h * 0.14, src, size=5.5, italic=True, clr=C_TEAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Commercial Intelligence Brief
# ══════════════════════════════════════════════════════════════════════════════

frame("COMMERCIAL INTELLIGENCE BRIEF")

# Company name
c.setFillColor(C_BLK)
c.setFont("Helvetica-Bold", 30)
c.drawString(in_(0.2), H - in_(1.3), "Hampr")

# Descriptor
text_at(in_(0.2), H - in_(1.55),
        "B2B workplace hospitality and corporate food technology marketplace   |   Sydney, NSW",
        size=9, clr=C_BLK)

# 6 stat cards
CW = in_(1.565); CH = in_(1.07); CY_TOP = in_(1.62); GAP = in_(0.064)
XS = [in_(0.12) + i * (CW + GAP) for i in range(6)]

CARDS = [
    ("AUD $12.7M",  "Revenue",               "FY2024",               "Smart50 2024, SmartCompany Nov 2024"),
    ("153%",        "Revenue growth",         "year on year FY2024",  "Smart50 2024, SmartCompany Nov 2024"),
    ("#8",          "Smart50 ranking",        "nationally, 2024",     "SmartCompany Smart50 Nov 2024"),
    ("37",          "Employees",              "full time equiv.",     "Smart50 2024, SmartCompany Nov 2024"),
    ("2018",        "Year founded",           "Sydney, NSW",          "Smart50 2024 profile"),
    ("$11.1M",      "Total funding raised",   "across 4 rounds",      "Crunchbase / Tracxn, May 2025"),
]
for i, (num, l1, l2, src) in enumerate(CARDS):
    stat_card(XS[i], CY_TOP, CW, CH, num, l1, l2, src)

# Key signals label
text_at(in_(0.2), H - in_(2.82), "KEY COMMERCIAL SIGNALS",
        size=8.5, bold=True, clr=C_TEAL)

SIGNALS = [
    "Singapore market entry FY2026: hampr.sg website live, international expansion stated publicly.  Source: Smart50 2024 profile, Nov 2024",
    "Two major enterprise tenders targeted within 12 months of Smart50 submission.  Source: Smart50 2024 profile, Nov 2024",
    "Enterprise client publicly named: Toyota Finance Australia referenced in company media.  Source: Momentum91 podcast, May 2025",
    "Investor backing includes Blackbird Ventures, Techstars, Startmate, Lee Kim Tah Group.  Source: Crunchbase / Tracxn, 2025",
    "National operations active across Sydney, Melbourne, Perth, Brisbane, and Canberra.  Source: Hampr website, 2025",
    "No publicly named CFO or Head of Finance in any company profile or media.  Source: Multiple profiles, 2025",
]
SY_FROM_TOP = in_(2.97)
for sig in SIGNALS:
    wrap_text(in_(0.2), H - SY_FROM_TOP, sig, 8, C_BLK, W - in_(0.4), in_(0.33))
    SY_FROM_TOP += in_(0.345)

c.showPage()


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — The Opportunity
# ══════════════════════════════════════════════════════════════════════════════

frame("THE OPPORTUNITY")

c.setFillColor(C_BLK)
c.setFont("Helvetica-Bold", 11)
c.drawString(in_(0.2), H - in_(1.06),
             "Hampr   |   Three commercial observations from ProfitPulse")

# Three columns
COL_X   = [in_(0.10), in_(3.43), in_(6.77)]
COL_W   = [in_(3.22), in_(3.22), in_(3.10)]
COL_BG  = [C_TEAL, C_BLK, C_GOLD]
COL_TC  = [C_WHT, C_WHT, C_BLK]
COL_HC  = [C_WHT, C_WHT, C_BLK]
COL_NC  = [C_BLK, C_AMB, C_BLK]
CT_TOP  = in_(1.22)
CH2     = in_(3.97)

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
     "even while reporting a healthy profit. ProfitPulse engagements routinely identify six to "
     "twelve percent of annual revenue in manageable float."),

    ("Hampr has moved into Singapore with a live website and a publicly stated FY2026 target. "
     "International expansion requires upfront investment in team, compliance, and market "
     "development well ahead of local revenue. With the most recent funding round at $1.08M "
     "in May 2025, a rigorous financial model for the Singapore launch is needed: one that "
     "maps cash burn, the revenue ramp, and the break even timeline. A Fractional CFO "
     "Partnership is designed to carry exactly this work."),

    ("Two enterprise tenders in active pursuit means a step change from a diversified SME "
     "client base to a smaller number of much larger contracts. Enterprise clients carry 60 to "
     "90 day payment terms and detailed financial reporting expectations. This concentrates "
     "revenue risk, extends the cash conversion cycle, and demands board grade reporting that "
     "a founder led business at this scale may not yet have in place. Getting ahead now "
     "protects margin and negotiating position."),
]

for i in range(3):
    cx, cw = COL_X[i], COL_W[i]
    cy = H - CT_TOP - CH2
    filled_rect(cx, cy, cw, CH2, COL_BG[i])
    # Index
    c.setFillColor(COL_NC[i])
    c.setFont("Helvetica-Bold", 22)
    c.drawString(cx + in_(0.12), H - CT_TOP - in_(0.5), IDXS[i])
    # Heading
    wrap_text(cx + in_(0.12), H - CT_TOP - in_(0.65),
              HEADS[i], 9, COL_HC[i], cw - in_(0.24), in_(0.19), bold=True)
    # Body
    wrap_text(cx + in_(0.12), H - CT_TOP - in_(1.35),
              BODIES[i], 8.5, COL_TC[i], cw - in_(0.24), in_(0.155))

# Warm close
wrap_text(in_(0.15), H - in_(5.27),
          ("These observations are offered in good faith. Hampr has built something genuinely "
           "impressive. The question is simply whether the financial architecture is ready for the next chapter."),
          8, C_BLK, W - in_(0.3), in_(0.14), italic=True)

c.showPage()


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — The Recommendation
# ══════════════════════════════════════════════════════════════════════════════

frame("THE RECOMMENDATION")

LX = in_(0.12); LW = in_(5.55)

# Service name
c.setFillColor(C_BLK)
c.setFont("Helvetica-Bold", 19)
c.drawString(LX + in_(0.05), H - in_(1.35), "Fractional CFO Partnership")

# Price
text_at(LX + in_(0.05), H - in_(1.65), "$7,500 per month",
        size=12, bold=True, clr=C_TEAL)

# Description
wrap_text(LX + in_(0.05), H - in_(1.9),
    ("A senior financial partner at the table on a monthly cadence. Monthly management pack, "
     "quarterly board grade review, ad hoc decision support, and a single annual deep dive. "
     "For Hampr: a financial rhythm for the Singapore launch, enterprise contract cash flow "
     "modelling, and the reporting infrastructure to match the ambition."),
    9, C_BLK, LW - in_(0.05), in_(0.16))

# Rule
filled_rect(LX + in_(0.05), H - in_(2.5), LW - in_(0.1), in_(0.015), C_TEAL)

# Step one accent bar
filled_rect(LX, H - in_(2.55) - in_(1.4), in_(0.04), in_(1.4), C_TEAL)
text_at(LX + in_(0.12), H - in_(2.57), "Step one: answer a few quick questions",
        size=9.5, bold=True, clr=C_BLK)
text_at(LX + in_(0.12), H - in_(2.78), "See the solutions matched to your size and industry.",
        size=9, clr=C_BLK)
# Questionnaire link (clean — no tracking on brief)
c.setFillColor(C_TEAL)
c.setFont("Helvetica-Bold", 9)
c.drawString(LX + in_(0.12), H - in_(2.98), "profit-pulse.com.au/full-suite-of-products")
c.linkURL(QUEST_URL,
          (LX + in_(0.12), H - in_(2.98) - 2,
           LX + in_(0.12) + c.stringWidth("profit-pulse.com.au/full-suite-of-products",
                                           "Helvetica-Bold", 9),
           H - in_(2.98) + 9))

# Rule
filled_rect(LX + in_(0.05), H - in_(3.25), LW - in_(0.1), in_(0.015), C_AMBD)

# Direct purchase CTA
filled_rect(LX, H - in_(3.28) - in_(0.54), in_(0.04), in_(0.54), C_AMBD)
text_at(LX + in_(0.12), H - in_(3.3), "Already know this is the priority?",
        size=9.5, bold=True, clr=C_BLK)
c.setFillColor(C_AMBD)
c.setFont("Helvetica-Bold", 9)
c.drawString(LX + in_(0.12), H - in_(3.52),
             "Purchase the suggested product now to get started")
c.linkURL(STRIPE_URL,
          (LX + in_(0.12), H - in_(3.52) - 2,
           LX + in_(0.12) + c.stringWidth("Purchase the suggested product now to get started",
                                          "Helvetica-Bold", 9),
           H - in_(3.52) + 9))

# Rule
filled_rect(LX + in_(0.05), H - in_(3.85), LW - in_(0.1), in_(0.015), C_TEAL)

# Discovery call
filled_rect(LX, H - in_(3.88) - in_(0.56), in_(0.04), in_(0.56), C_AMB)
text_at(LX + in_(0.12), H - in_(3.9), "Prefer a conversation first?",
        size=9.5, bold=True, clr=C_BLK)
c.setFillColor(C_TEAL)
c.setFont("Helvetica", 9)
c.drawString(LX + in_(0.12), H - in_(4.12),
             "Book a complimentary discovery call with Nitesh Roopa")
c.linkURL(BOOKING_URL,
          (LX + in_(0.12), H - in_(4.12) - 2,
           LX + in_(0.12) + c.stringWidth("Book a complimentary discovery call with Nitesh Roopa",
                                          "Helvetica", 9),
           H - in_(4.12) + 9))

# ── Vertical divider ─────────────────────────────────────────────────────────
filled_rect(in_(5.76), H - in_(5.45), in_(0.03), in_(4.57), C_TEAL)

# ── Right column — credibility ────────────────────────────────────────────────
RX = in_(5.84); RW = in_(4.05)
filled_rect(RX, H - in_(5.45), RW, in_(4.57), C_BLK)

c.setFillColor(C_AMB)
c.setFont("Helvetica-Bold", 16)
c.drawString(RX + in_(0.12), H - in_(1.32), "Nitesh Roopa")

text_at(RX + in_(0.12), H - in_(1.6), "CA, Managing Partner   |   ProfitPulse",
        size=9.5, clr=C_OWHT)
filled_rect(RX + in_(0.12), H - in_(1.72), RW - in_(0.25), in_(0.02), C_TEAL)

CREDS = [
    "16 years across 4 countries",
    "52 deals executed and managed across the career",
    "Largest single deal: USD 1.3 billion, Cahora Bassa, Mozambique Government, Hydro",
    "Total GRBT project value in Queensland: over AUD 10 billion",
]
CY_RL = in_(1.82)
for cr in CREDS:
    wrap_text(RX + in_(0.12), H - CY_RL, cr, 8, C_OWHT, RW - in_(0.25), in_(0.14))
    CY_RL += in_(0.32)

filled_rect(RX + in_(0.12), H - CY_RL - in_(0.02), RW - in_(0.25), in_(0.02), C_TEAL)
CY_RL += in_(0.1)

for ctxt in ["Profit-Pulse.com.au", "Nitesh@Profit-Pulse.com.au",
             "+61 411 876 267", "linkedin.com/in/nitesh-roopa-77594163"]:
    text_at(RX + in_(0.12), H - CY_RL, ctxt, size=8, clr=C_TEAL)
    CY_RL += in_(0.27)

c.showPage()

# ── Save ─────────────────────────────────────────────────────────────────────
c.save()
print(f"Saved: {out_path}")
