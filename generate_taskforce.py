# -*- coding: utf-8 -*-
"""
ProfitPulse nightly outreach generator for Taskforce Australia, 01 Jul 2026.
Builds the HTML dossier and the three slide PPTX brief from verified data.
"""
import build_dossier
import build_brief_v3
import build_pdf_v3

DATE_STR = "01 Jul 2026"
COMPANY = "Taskforce Australia"
SLUG = "taskforce_australia"

QUESTIONNAIRE_BASE = "https://profit-pulse.com.au/full-suite-of-products"
QUESTIONNAIRE_HTML = (QUESTIONNAIRE_BASE +
    f"?utm_source=outreach&utm_medium=html&utm_campaign=nightly_outreach&utm_content={SLUG}")
QUESTIONNAIRE_EMAIL = (QUESTIONNAIRE_BASE +
    f"?utm_source=outreach&utm_medium=email&utm_campaign=nightly_outreach&utm_content={SLUG}")
STRIPE_LINK = "https://buy.stripe.com/28EbJ2edi0LSbdugYT3ks1D"  # G3 Command, $3,950 one off
BOOKING_LINK = "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true"

WEDGE_TEXT = (
    "Taskforce has taken its revenue from $7.43 million to $12.8 million across three "
    "consecutive Smart50 listings while holding headcount close to flat at 19 to 20 people. "
    "That is the signature of a business stacking distinct revenue lines: the original job "
    "allocation marketplace, the RentSafe warranty and installation SaaS platform now serving "
    "20 major consumer brands across 300 real estate offices, and a newer housing division that "
    "the company's own 2025 Smart50 listing names as the strongest and most profitable area of "
    "the business. Naming one line as the most profitable in a public award citation is itself a "
    "signal. It suggests the comparison has already been made informally, but the formal, "
    "defensible, line by line version, gross margin, contribution margin and operational drag, "
    "has likely not been built. The next phase of investment, including the in house software "
    "capability the company says it is building specifically to find more cost effective ways "
    "to deliver, is exactly the kind of capital allocation decision that should follow a rigorous "
    "profitability map rather than precede it."
)

DOSSIER_DATA = {
    "company_name": COMPANY,
    "date_str": DATE_STR,
    "framing": (
        "Selected for three consecutive years of Smart50 recognised growth on a flat headcount, "
        "with the company's own latest listing flagging uneven profitability across its three "
        "lines of business."
    ),
    "profile_rows": [
        {"label": "Legal and trading name", "value": "Taskforce Australia Pty Ltd",
         "source": "SmartCompany Smart50 listings, 2023 to 2025"},
        {"label": "Website", "value": "taskforce.com.au", "source": "Company website"},
        {"label": "Location", "value": "Burnley, Melbourne VIC",
         "source": "SmartCompany Smart50 2023 and 2025 listings"},
        {"label": "Industry", "value":
         "Trade services and SaaS warranty, installation and compliance platform, classified by "
         "SmartCompany under real estate and construction",
         "source": "SmartCompany Smart50 listings"},
        {"label": "Founded", "value": "2014", "source": "SmartCompany Smart50 2023 listing"},
        {"label": "Decision maker", "value": "Jason Bright, Founder and Managing Director",
         "source": "Elite Agent, Telstra Best of Business Awards coverage"},
        {"label": "Revenue FY2025", "value": "$12.8 million, 31 percent growth, Smart50 rank 37",
         "source": "SmartCompany Smart50 Awards 2025"},
        {"label": "Revenue FY2024", "value": "$9.7 million, 26 percent growth, Smart50 rank 46",
         "source": "SmartCompany Smart50 Awards 2024"},
        {"label": "Revenue FY2023", "value": "$7.43 million, 108.61 percent growth",
         "source": "SmartCompany Smart50 Awards 2023"},
        {"label": "Team size", "value": "19 employees, 2025 and 2024 listings; 20 in 2023",
         "source": "SmartCompany Smart50 Awards, 2023 to 2025"},
        {"label": "Tradie network", "value":
         "Approximately 5,500 tradespeople in 2025, grown from approximately 1,000 in 2019",
         "source": "SmartCompany Smart50 2025; SmartCompany StartupSmart, 23 August 2019"},
        {"label": "RentSafe platform", "value":
         "Launched 2021, over 140,000 jobs delivered across 20 major consumer brands and 300 "
         "real estate offices", "source": "SmartCompany Smart50 2023 listing"},
        {"label": "Capital raised", "value": "$1.5 million from XSallarate and Futurist Capital",
         "source": "SmartCompany StartupSmart, 23 August 2019"},
        {"label": "Recent award", "value":
         "Victorian State Winner, Outstanding Growth, Telstra Best of Business Awards 2024",
         "source": "Elite Agent, national gala 28 November 2024"},
        {"label": "Employee band fit", "value": "Confirmed within 10 to 50",
         "source": "Derived from Smart50 listings above"},
    ],
    "wedge_html": WEDGE_TEXT,
    "service_name": "Product and Service Line Profitability",
    "service_description": (
        "A three week project ranking marketplace jobs, the RentSafe SaaS platform, and the "
        "housing division by gross margin, contribution margin and operational drag."
    ),
    "price_line": "$3,950 one off",
    "price_only": "$3,950 one off",
    "supporting_services": [
        "Operational Intelligence Review",
        "KPI Dashboard Build and Run",
    ],
    "questionnaire_link_tagged": QUESTIONNAIRE_HTML,
    "stripe_link": STRIPE_LINK,
    "signals": [
        "Revenue compounded from $7.43 million to $12.8 million across three Smart50 listings, "
        "2023 to 2025 (SmartCompany Smart50 Awards 2023, 2024, 2025).",
        "Headcount held close to flat at 19 to 20 people through the full three year run "
        "(SmartCompany Smart50 listings).",
        "RentSafe SaaS platform has processed over 140,000 jobs across 20 major consumer brands "
        "(SmartCompany Smart50 2023 listing).",
        "The company's own 2025 listing names its housing division as the most profitable line "
        "in the business (SmartCompany Smart50 2025).",
        "Named Victorian State Winner, Telstra Best of Business Awards, Outstanding Growth "
        "category, 2024 (Elite Agent, national gala 28 November 2024).",
        "Raised $1.5 million from XSallarate and Futurist Capital to expand sales and the "
        "platform (SmartCompany StartupSmart, 23 August 2019).",
    ],
    "email_html": f"""
    <div class="field">Subject</div>
    <div class="subject">Taskforce's three lines and the margin question</div>
    <p>Hi Jason,</p>
    <p>I noticed Taskforce's run on the Smart50 list this year, three years running now, with
    revenue moving from $7.43 million to $12.8 million while your team has stayed close to
    twenty people. That is a genuinely hard thing to pull off, and the housing division note in
    your 2025 listing caught my eye.</p>
    <p>I'm Nitesh Roopa, a Chartered Accountant and the Managing Partner of ProfitPulse, a
    Brisbane based fractional CFO practice. Across sixteen years and four countries I've worked
    on more than fifty deals, including financial leadership on a ten billion dollar Queensland
    infrastructure business case, so most of my work sits at exactly this kind of growth and
    complexity point.</p>
    <p>With three distinct lines now running under one roof, the original job marketplace,
    RentSafe, and housing, the natural next question is which of them actually carries the
    margin once the full cost to serve is allocated. Our Product and Service Line Profitability
    engagement answers that in three weeks. I've put together a short brief on Taskforce
    specifically, attached as a PDF.</p>
    <p>If it would help, the quickest way to see what fits is to answer a few short questions at
    <a href="{QUESTIONNAIRE_EMAIL}">profit-pulse.com.au/full-suite-of-products</a>, which matches
    solutions to your size and industry.</p>
    <p>If you would rather talk first, I'm always happy to do a complimentary discovery call, you
    can find a time directly at the link below. Either way, congratulations on the run you're
    on.</p>
    <div class="sig">
    Nitesh Roopa<br>
    CA, Managing Partner<br>
    ProfitPulse<br>
    <a href="https://profit-pulse.com.au">Profit-Pulse.com.au</a><br>
    <a href="mailto:Nitesh@Profit-Pulse.com.au">Nitesh@Profit-Pulse.com.au</a><br>
    +61 411 876 267<br>
    <a href="{BOOKING_LINK}">Book a complimentary discovery call</a>
    </div>
    """,
    "system_notes": [
        "Selected target: Taskforce Australia Pty Ltd, Burnley VIC. Confirmed not present in "
        "Outreach_Register.csv prior to this run.",
        "Verified sources: SmartCompany Smart50 Awards 2023, 2024 and 2025 listings for revenue, "
        "growth, employees and founding year; SmartCompany StartupSmart, 23 August 2019, for the "
        "funding raise; Elite Agent, Telstra Best of Business Awards coverage, for the decision "
        "maker title and the 2024 Victorian State Winner result.",
        "Sourcing method note: direct WebFetch to SmartCompany and related primary domains "
        "returned HTTP 403 this session due to an organisation level egress policy, confirmed via "
        "the proxy status endpoint, not a per site block. Every cited figure was triangulated "
        "through multiple independently phrased WebSearch queries returning consistent, "
        "attributed text from the named, dated source pages. No figure is included that was not "
        "corroborated by a specific, named public source.",
        "Inferred commercial goal, marked as an inference: with three distinct revenue lines now "
        "running concurrently and the business's own most recent award listing naming one line as "
        "most profitable, Taskforce most likely needs a formal, line by line profitability view "
        "to guide where the next phase of software and headcount investment should go.",
        "Matched primary service code and name: G3, Product and Service Line Profitability.",
        "Supporting services named: G1 Operational Intelligence Review; D3 KPI Dashboard Build "
        "and Run.",
        "Inferred tier: Command. Reason: Section 3C tier engine raw score 31 against an industry "
        "bias of 3 (Technology and SaaS), giving an Essential ceiling of 27 and a Command ceiling "
        "of 37; no lift or drop adjustment applied since the assumed gross margin band (mid, 40 "
        "to 55 percent) does not meet either adjustment threshold.",
        "Assumptions recorded per Section 3C default guidance for inputs that are not public: "
        "gross margin assumed mid band, 40 to 55 percent; operating locations assumed single "
        "site; business structure assumed company plus trust or two entities; revenue model "
        "assumed a mix of recurring and one off, consistent with the marketplace plus SaaS mix "
        "actually observed; customer concentration assumed some reliance on key accounts; "
        "funding and debt position assumed modest debt with no covenants; forward growth assumed "
        "growing, 15 to 30 percent; recent trajectory assessed as growing based on the verified "
        "31 percent FY2025 growth rate; commercial event assumed expansion, based on the verified "
        "housing division and software capability growth signals.",
        f"Questionnaire URL used in the dossier: {QUESTIONNAIRE_HTML}",
        f"Questionnaire URL used in the email draft: {QUESTIONNAIRE_EMAIL}",
        f"Direct Stripe checkout link recorded: {STRIPE_LINK}, Product and Service Line "
        "Profitability, Command tier, $3,950 one off.",
        "Email did not include the direct checkout link, consistent with "
        "EMAIL_INCLUDE_DIRECT_CHECKOUT set to false in Section 1.",
        "PDF build method note: headless LibreOffice conversion from the PPTX is unavailable in "
        "this environment, so the PDF was generated directly from the same content data using "
        "ReportLab, matching the PPTX text, structure and links. Each PDF page was rendered to "
        "an image and visually inspected against the Section 6.6 checklist before saving; two "
        "issues found on the first pass, dead space under the slide two observation columns and "
        "an oversized slide three credibility panel, were corrected and re inspected.",
        "Confirmation: nothing has been sent. This dossier, the brief and the email draft are "
        "for Nitesh Roopa's review only.",
    ],
}

BRIEF_DATA = {
    "date_str": DATE_STR,
    "company_name": COMPANY,
    "descriptor": "Trade services and SaaS warranty platform, Burnley, Melbourne VIC",
    "stat_cards": [
        {"number": "$12.8M", "label": "FY2025 revenue, Smart50 rank 37",
         "source": "SmartCompany Smart50 2025"},
        {"number": "31%", "label": "FY2025 revenue growth rate",
         "source": "SmartCompany Smart50 2025"},
        {"number": "19", "label": "Employees, Smart50 listing",
         "source": "SmartCompany Smart50 2025"},
        {"number": "2014", "label": "Year founded, Burnley VIC",
         "source": "SmartCompany Smart50 2023"},
        {"number": "5,500", "label": "Tradespeople in active network",
         "source": "SmartCompany Smart50 2025"},
        {"number": "$1.5M", "label": "Capital raised, XSallarate, Futurist",
         "source": "SmartCompany, 23 Aug 2019"},
    ],
    "signals": [
        "Revenue compounded from $7.43 million to $12.8 million across three Smart50 listings, "
        "2023 to 2025.",
        "Headcount held close to flat at 19 to 20 people through the full three year run.",
        "RentSafe SaaS platform has processed over 140,000 jobs across 20 major consumer brands.",
        "Company's own 2025 listing names its housing division as the most profitable line.",
        "Named Victorian State Winner, Telstra Best of Business Awards, Outstanding Growth, 2024.",
        "Raised $1.5 million from XSallarate and Futurist Capital to expand sales and platform.",
    ],
    "observations": [
        {"index": "01", "header": "Three lines, unproven margins", "body":
         "Taskforce runs three lines: the original tradie marketplace, the RentSafe SaaS "
         "platform serving 20 major brands, and a housing division the company itself calls its "
         "most profitable line. A claim like that, made without published unit economics, "
         "usually means the comparison was done by feel, not a formal margin model."},
        {"index": "02", "header": "Revenue up, headcount flat", "body":
         "Revenue moved from $7.43 million to $12.8 million across three Smart50 cycles while "
         "the team held at 19 to 20 people. That is strong apparent productivity, but more lines "
         "and bigger brand contracts are now being managed without a matching increase in "
         "finance or reporting capacity."},
        {"index": "03", "header": "A cost signal without a profit map", "body":
         "The 2025 listing notes Taskforce is expanding in house software specifically to find "
         "more cost effective delivery. That is a profit instinct without yet the profit map. "
         "Knowing which line actually carries the margin would show the software team exactly "
         "where that investment pays back fastest."},
    ],
    "warm_line": (
        "These are observations offered in good faith. Taskforce has built a genuinely "
        "differentiated platform business. The question is simply whether the margin map keeps "
        "pace with the lines it now runs."
    ),
    "service_name": "Product and Service Line Profitability",
    "price_line": "$3,950 one off",
    "service_blurb": (
        "A three week project ranking marketplace jobs, RentSafe SaaS, and housing by gross "
        "margin, contribution margin and operational drag."
    ),
    "stripe_link": STRIPE_LINK,
    "booking_link": BOOKING_LINK,
    "credibility_points": [
        "16 years across 4 countries",
        "52 deals executed and managed",
        "Largest single deal USD 1.3 billion, Cahora Bassa, Mozambique",
        "Total GRBT project value over AUD 10 billion",
    ],
}

if __name__ == "__main__":
    out_html = "/home/user/Daily-Outreach/Out-reach efforts/Dossier_TaskforceAustralia_01Jul2026.html"
    out_pptx = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Jul2026.pptx"
    out_pdf = "/home/user/Daily-Outreach/Out-reach efforts/Brief_TaskforceAustralia_01Jul2026.pdf"
    build_dossier.build(DOSSIER_DATA, out_html)
    build_brief_v3.build(BRIEF_DATA, out_pptx)
    build_pdf_v3.build(BRIEF_DATA, out_pdf)
