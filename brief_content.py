"""
Shared content and layout constants for the Equality Media + Marketing brief.
Date for: 17 Jul 2026. Run date: 16 Jul 2026 23:05 Brisbane.
Imported by build_pdf.py and build_brief.py so the PDF and PPTX stay in sync.
"""

COMPANY = "Equality Media + Marketing"
COMPANY_SHORT = "Equality Media"
LOCATION_LINE = "Independent media and marketing agency, Richmond, Melbourne VIC"
DATE_FOR = "17 Jul 2026"
DATE_SLUG = "17Jul2026"
FILE_SLUG = "EqualityMedia"
UTM_SLUG = "equality_media_marketing"

QUESTIONNAIRE_CLEAN = "profit-pulse.com.au/services/find-your-fit"
QUESTIONNAIRE_URL_BASE = "https://profit-pulse.com.au/services/find-your-fit/"
QUESTIONNAIRE_URL_HTML = (
    QUESTIONNAIRE_URL_BASE
    + f"?utm_source=outreach&utm_medium=html&utm_campaign=nightly_outreach&utm_content={UTM_SLUG}"
)
QUESTIONNAIRE_URL_EMAIL = (
    QUESTIONNAIRE_URL_BASE
    + f"?utm_source=outreach&utm_medium=email&utm_campaign=nightly_outreach&utm_content={UTM_SLUG}"
)
BOOKING_LINK = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

# Matched service (primary): A4 Strategic Growth Diagnostic, Command tier
PRIMARY_SERVICE_NAME = "Strategic Growth Diagnostic"
PRIMARY_SERVICE_CODE = "A4"
PRIMARY_SERVICE_PRICE = "$5,000"
PRIMARY_SERVICE_CADENCE = "one off"
PRIMARY_SERVICE_STRIPE = "https://buy.stripe.com/eVqdRad9e66cftK23Z3ks0h"
PRIMARY_SERVICE_DESC = (
    "A six week engagement that maps revenue, capacity and margin headroom, then "
    "produces a 12 month growth plan with funding and capital allocation steps and "
    "three scenarios, so next year's growth is engineered rather than extrapolated."
)
SUPPORTING_SERVICES = [
    "Customer Concentration and Profitability Map",
    "Fractional CFO Partnership",
]

# Tier engine (Section 3C), fully worked
TIER_INPUTS = [
    ("Annual revenue", "$10M to $20M ($13M, FY23 to FY24)", 6),
    ("Gross margin", "Not publicly disclosed, neutral default 40% to 55%", 3),
    ("Team size", "16 to 40 (15 to 17 verified, reclassified to a larger award band in 2026)", 3),
    ("Operating locations", "Single location (Richmond, Melbourne)", 1),
    ("Business structure", "Single entity", 1),
    ("Revenue model", "Mix of recurring and one off (retainer plus campaign)", 2),
    ("Customer concentration", "Some reliance on key accounts, neutral default", 2),
    ("Funding and debt position", "Not publicly disclosed, neutral default modest debt", 2),
    ("Forward growth", "High growth, over 30% (49% revenue, 65% team growth cited)", 6),
    ("Recent trajectory", "Scaling rapidly", 6),
    ("Commercial event", "None publicly signalled", 0),
]
TIER_RAW_SCORE = sum(x[2] for x in TIER_INPUTS)  # 32
TIER_INDUSTRY = "Marketing, creative and media agencies"
TIER_BIAS = 2
TIER_ESSENTIAL_CEILING = 27 - (TIER_BIAS - 3) * 2  # 29
TIER_COMMAND_CEILING = 37 - (TIER_BIAS - 3) * 2  # 39
TIER_REASON = (
    f"Raw score {TIER_RAW_SCORE} against bias {TIER_BIAS} ceilings "
    f"(Essential {TIER_ESSENTIAL_CEILING}, Command {TIER_COMMAND_CEILING}) gives base tier Command; "
    "no lift (margin band above 2) or drop (margin band not 5) applies. Final tier Command."
)
TIER_NAME_INTERNAL_ONLY = "Command"  # NEVER shown on brief, email, or dossier body

# Stat cards for slide 1 (6 cards). number <=6 chars, label 2 lines ~30 chars, source ~28 chars
STAT_CARDS = [
    {"number": "$13M", "label": "Revenue, FY23 to FY24", "source": "SmartCompany Smart50 2025"},
    {"number": "#24", "label": "Smart50 2025 national rank", "source": "SmartCompany, 2025"},
    {"number": "49%", "label": "Revenue growth, year on year", "source": "SmartCompany Smart50 2025"},
    {"number": "65%", "label": "Team growth, past year", "source": "B&T Fast 10, 2025"},
    {"number": "2018", "label": "Founded, Marilla Akkermans", "source": "Company site, LinkedIn"},
    {"number": "#55", "label": "AFR Fast 100 2025 rank", "source": "B&T, 2025"},
]

# Verified revenue trajectory, 3 real points, for the slide 1 chart
CHART_DATA = [
    ("2022", 3.2),
    ("FY22 to 23", 7.3),
    ("FY23 to 24", 13.0),
]
CHART_SOURCE = "Source: Future of Australia Podcast Ep53 (2022); SmartCompany Smart50 2024 and 2025 citations"

KEY_SIGNALS = [
    "Revenue compounded from $3.2M (2022) to $13M (FY23 to FY24), per Smart50 citations",
    "Climbed from Smart50 rank 45 in 2024 to rank 24 in 2025, SmartCompany reporting",
    "Team grew 65% in the past year while revenue grew 49%, per B&T Fast 10 profile",
    "Landed the BIODERMA skincare media account, a named new client win, per B&T",
    "Gold, Best Small Organisation, AFR BOSS Best Places to Work 2026, per Mi3 and AdNews",
]

# Slide 2, three observations. First (teal) aligns with the primary recommended service.
OBSERVATIONS = [
    {
        "index": "01",
        "fill": "TEAL",
        "header": "Growth is compounding faster than a plan can absorb it",
        "body": (
            "Revenue has moved from $3.2 million to $13 million in roughly two years, a "
            "trajectory that rarely continues without strain appearing somewhere, in "
            "delivery capacity, in margin, or in cash timing. The Smart50 citations "
            "documenting this growth show it accelerating, not levelling off. The "
            "practical question for the next 12 months is not whether Equality Media can "
            "keep growing, the evidence says it can, but which capacity, hiring and "
            "capital steps actually support another year at this pace."
        ),
    },
    {
        "index": "02",
        "fill": "BLACK",
        "header": "One vertical carries an outsized share of the story",
        "body": (
            "Trade press coverage repeatedly ties this growth to the property sector "
            "specifically, citing close to a billion dollars in property sales supported "
            "by its campaigns in a single year. That is a genuine strength and a "
            "concentration signal at once. When one vertical carries this much of the "
            "narrative, a slowdown in that sector's marketing spend, not any failure of "
            "the agency's own work, becomes the largest single risk to next year's "
            "revenue line."
        ),
    },
    {
        "index": "03",
        "fill": "GOLD",
        "header": "Headcount is scaling as fast as revenue, which changes the question",
        "body": (
            "A reported 65% increase in team size in a single year, alongside "
            "consecutive Best Places to Work recognitions, shows real investment in "
            "people. This is exactly the phase where reporting and cash forecasting "
            "built for a $3 million agency stop scaling cleanly to a $13 million one "
            "with several times the staff. A monthly finance rhythm is the natural next "
            "step once the growth plan itself is set."
        ),
    },
]
CLOSING_LINE = (
    "These are observations offered in good faith. Equality Media has built something "
    "genuinely admired in its market. The question is simply whether the financial "
    "architecture is scaling at the same pace as the team."
)

WEDGE_TEXT = (
    "Equality Media has compounded from $3.2 million to $13 million in revenue across "
    "roughly two years, moved from Smart50 rank 45 to rank 24, and grown its team by "
    "65 percent in the past year alone, all while being recognised nationally as a best "
    "place to work. That is an unusually clean growth story.\n\n"
    "The question a Chartered Accountant would ask next is whether the next 12 months "
    "of growth is engineered or simply extrapolated: which capacity, hiring and capital "
    "allocation steps actually support another year at this pace, and where the plan "
    "would break first if growth kept compounding at even half this rate.\n\n"
    "The Strategic Growth Diagnostic answers that in six weeks: revenue, capacity and "
    "margin headroom mapped, then a 12 month growth plan with funding and capital "
    "allocation steps spelled out across three scenarios."
)

# Verified firmographics for the HTML dossier (Section 5.2)
VERIFIED_PROFILE = [
    ("Company", "Equality Media + Marketing (Equality Media Pty Ltd)"),
    ("Location", "Richmond, Melbourne VIC"),
    ("Industry", "Marketing, creative and media agencies"),
    ("Founded", "2018, by Marilla Akkermans. Source: SmartCompany, B&T, LinkedIn"),
    ("Team size", "11 to 50 (LinkedIn); 17 (Smart50 2024); 15 (Smart50 2025); reclassified into a 20 to 99 employee award band, AFR BOSS 2026"),
    ("Revenue", "$3.2M (2022, Future of Australia Podcast Ep53); $7.3M (FY22 to 23, Smart50 2024); $13M (FY23 to 24, Smart50 2025)"),
    ("Growth", "49% year on year (Smart50 2025); team grew 65% in the past year (B&T Fast 10, 2025)"),
    ("Rankings", "Smart50 2025 rank 24 of 50; AFR Fast 100 2025 rank 55; Gold, Best Small Organisation, AFR BOSS Best Places to Work 2026"),
    ("Gross margin", "Not publicly disclosed"),
    ("Funding or debt position", "Not publicly disclosed, no public signal of a funding round or debt raise"),
    ("Decision maker", "Marilla Akkermans, Founder and Managing Director. Source: company site, LinkedIn, multiple trade press profiles"),
]

# Section 5.3C, likely contact emails
CONTACT_DOMAIN = "equalitymedia.co"
CONTACT_CANDIDATES = [
    {"email": "hello@equalitymedia.co", "tag": "Verified", "note": "Published on the company's own Privacy Policy page as its contact address"},
    {"email": "marilla@equalitymedia.co", "tag": "Inferred", "note": "Most common indexed format for this domain per third party aggregation, unverified"},
    {"email": "marilla.akkermans@equalitymedia.co", "tag": "Inferred", "note": "Common first.last format, unverified"},
    {"email": "marillaakkermans@equalitymedia.co", "tag": "Inferred", "note": "Common firstlast format, unverified"},
    {"email": "makkermans@equalitymedia.co", "tag": "Inferred", "note": "Common first initial plus last format, unverified"},
    {"email": "marilla_akkermans@equalitymedia.co", "tag": "Inferred", "note": "Underscore format, unverified"},
]
EMAIL_TO_PLACEHOLDER = "marilla@equalitymedia.co"  # most likely guess, placeholder only, NOT verified

DECISION_MAKER_NAME = "Marilla Akkermans"
DECISION_MAKER_TITLE = "Founder and Managing Director"

EMAIL_SUBJECT = "Planning Equality Media's next 12 months of growth"
