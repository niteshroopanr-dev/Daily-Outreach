"""
ProfitPulse HTML Dossier Builder, Version 3.
Self contained, brand styled, internal review file for Nitesh only.
Fill in DATA below and run: python3 build_dossier.py
"""

import html as _html


def esc(s):
    return _html.escape(str(s))


def build(data, out_path):
    profile_rows = "".join(
        f"""<tr><td class="k">{esc(r['label'])}</td><td class="v">{esc(r['value'])}</td>
        <td class="s">{esc(r['source'])}</td></tr>"""
        for r in data["profile_rows"]
    )

    signals = "".join(f"<li>{esc(s)}</li>" for s in data["signals"])

    supporting = ""
    if data.get("supporting_services"):
        supporting = "<p class='supporting'><strong>Supporting services named:</strong> " + \
            ", ".join(esc(s) for s in data["supporting_services"]) + "</p>"

    email_body_html = data["email_html"]

    notes = "".join(f"<li>{esc(n)}</li>" for n in data["system_notes"])

    contact_rows = "".join(
        f"""<tr><td class="ck">{esc(c['address'])}</td><td class="cv">{esc(c['status'])}</td>
        <td class="cs">{esc(c['basis'])}</td></tr>"""
        for c in data["contact_emails"]
    )
    contact_email_html = f"""<table class="contact-table">
      <tr><td class="ck">Address</td><td class="cv">Status</td><td class="cs">Basis</td></tr>
      {contact_rows}
    </table>
    <p class="contact-note">{esc(data['contact_email_note'])}</p>"""

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ProfitPulse Dossier, {esc(data['company_name'])}, {esc(data['date_str'])}</title>
<style>
  :root {{
    --black:#000000; --teal:#01A296; --amber-b:#F8C806; --amber-d:#F6A102;
    --gold:#E3A712; --white:#FFFFFF; --off-white:#E6E5DE;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin:0; padding:0; background:var(--white); color:var(--black);
    font-family: Calibri, Arial, sans-serif; line-height:1.5;
  }}
  .wrap {{ max-width: 880px; margin: 0 auto; padding: 0 20px 60px; }}
  header.top {{
    background: var(--black); color: var(--white); padding: 28px 20px;
    margin-bottom: 28px;
  }}
  header.top .wrap2 {{ max-width: 880px; margin:0 auto; }}
  header.top h1 {{
    font-family: Cambria, Georgia, serif; color: var(--amber-b);
    font-size: 32px; margin: 0 0 6px;
  }}
  header.top .meta {{ color: var(--off-white); font-size: 14px; }}
  header.top .framing {{ color: var(--white); font-size: 15px; margin-top: 10px; max-width: 700px;}}
  h2.section {{
    font-family: Cambria, Georgia, serif; color: var(--black);
    border-left: 6px solid var(--teal); padding-left: 12px;
    font-size: 21px; margin-top: 40px;
  }}
  table.profile {{ width:100%; border-collapse: collapse; margin-top: 14px; }}
  table.profile td {{ padding: 8px 10px; border-bottom: 1px solid #e2e2e2; vertical-align: top; font-size: 13.5px; }}
  table.profile td.k {{ font-weight: 700; width: 22%; color:#222; }}
  table.profile td.v {{ width: 38%; }}
  table.profile td.s {{ width: 40%; color:#555; font-size: 12px; font-style: italic; }}
  .wedge-box {{ background:#f6f6f4; border-left: 4px solid var(--gold); padding: 18px 20px; margin-top: 16px; font-size: 14.5px; }}
  .service-block {{ background: var(--black); color: var(--white); padding: 22px 24px; margin-top: 20px; border-top: 4px solid var(--teal); }}
  .service-block h3 {{ color: var(--amber-b); font-family: Cambria, Georgia, serif; margin: 0 0 8px; font-size: 19px; }}
  .service-block p {{ margin: 6px 0; font-size: 14px; color: var(--off-white); }}
  .service-block .price {{ color: var(--teal); font-weight: 700; font-size: 15px; }}
  .supporting {{ font-size: 13px; color:#444; margin-top: 10px; }}
  .cta-primary {{ margin-top: 22px; }}
  .cta-primary a.btn {{
    display:inline-block; background:var(--teal); color:var(--white);
    font-weight:700; text-decoration:none; padding:16px 28px; border-radius:30px;
    font-size:16px;
  }}
  .cta-primary .sub {{ color:var(--black); font-size:14px; margin-top:8px; }}
  .cta-secondary {{ margin-top: 18px; padding: 14px 18px; border: 1px solid var(--amber-d); border-radius: 8px; }}
  .cta-secondary a.btn2 {{
    display:inline-block; border:1px solid var(--amber-d); color:var(--amber-d);
    font-weight:700; text-decoration:none; padding:12px 22px; border-radius:26px;
    font-size:15px; margin-top:8px;
  }}
  ul.signals {{ font-size: 14px; padding-left: 22px; }}
  ul.signals li {{ margin-bottom: 8px; }}
  .review-banner {{
    background:#B00020; color:var(--white); font-weight:700; padding:16px 20px;
    border-radius: 6px; font-size: 14.5px; margin-bottom: 18px;
  }}
  .email-box {{ background:#fafaf8; border:1px solid #ddd; padding: 24px; margin-top: 10px; }}
  .email-box .field {{ font-size: 13px; color:#555; margin-bottom: 4px; }}
  .email-box .subject {{ font-weight:700; font-size: 15px; margin: 6px 0 16px; }}
  .email-box p {{ font-size: 14px; margin: 0 0 14px; }}
  .email-box .sig {{ font-size: 14px; margin-top: 10px; }}
  .email-box a {{ color: var(--teal); }}
  ul.notes {{ font-size: 12.5px; color:#444; padding-left: 20px; }}
  ul.notes li {{ margin-bottom: 6px; }}
  table.contact-table {{ width:100%; border-collapse: collapse; margin-top: 14px; }}
  table.contact-table td {{ padding: 7px 10px; border-bottom: 1px solid #e2e2e2; vertical-align: top; font-size: 13px; }}
  table.contact-table td.ck {{ font-weight: 700; width: 34%; color:#222; }}
  table.contact-table td.cv {{ width: 18%; }}
  table.contact-table td.cs {{ width: 48%; color:#555; font-size: 12px; font-style: italic; }}
  .contact-note {{ font-size: 12.5px; color:#555; margin-top: 10px; }}
  footer.bottom {{ margin-top: 50px; font-size: 11px; color:#888; text-align:center; }}
</style>
</head>
<body>
<header class="top">
  <div class="wrap2">
    <h1>{esc(data['company_name'])}</h1>
    <div class="meta">{esc(data['date_str'])}</div>
    <div class="framing">{esc(data['framing'])}</div>
  </div>
</header>
<div class="wrap">

  <h2 class="section">Verified profile</h2>
  <table class="profile">
    <tr><td class="k">Field</td><td class="v">Value</td><td class="s">Public basis</td></tr>
    {profile_rows}
  </table>

  <h2 class="section">The wedge and the matched service</h2>
  <div class="wedge-box">{data['wedge_html']}</div>

  <div class="service-block">
    <h3>{esc(data['service_name'])}</h3>
    <p>{esc(data['service_description'])}</p>
    <p class="price">{esc(data['price_line'])}. This figure is a verified ProfitPulse price; the questionnaire confirms the exact fit for {esc(data['company_name'])}.</p>
  </div>
  {supporting}

  <div class="cta-primary">
    <a class="btn" href="{esc(data['questionnaire_link_tagged'])}">Find your fit in two minutes</a>
    <div class="sub">Answer a few quick questions and see the solutions matched to your size and industry.</div>
    <div class="sub" style="margin-top:4px;color:#555;">profit-pulse.com.au/services/find-your-fit</div>
  </div>

  <div class="cta-secondary">
    Already know this is the priority? You can begin with {esc(data['service_name'])}, {esc(data['price_only'])}.<br>
    <a class="btn2" href="{esc(data['stripe_link'])}">Purchase the suggested product now to get started</a>
  </div>

  <h2 class="section">Key commercial signals</h2>
  <ul class="signals">{signals}</ul>

  <h2 class="section">Likely contact email addresses, internal use only</h2>
  {contact_email_html}

  <h2 class="section">The email draft</h2>
  <div class="review-banner">REVIEW GATE: This email is a draft only. No communication has been sent. Only Nitesh Roopa may authorise and send this message.</div>
  <div class="email-box">
    <div class="field">To: {esc(data['email_to_placeholder'])} (placeholder, confirm before any send)</div>
    {email_body_html}
  </div>

  <h2 class="section">System notes</h2>
  <ul class="notes">{notes}</ul>

  <footer class="bottom">ProfitPulse internal dossier. For Nitesh Roopa review only. Not for external distribution.</footer>
</div>
</body>
</html>
"""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_doc)
    print(f"HTML saved: {out_path}")


if __name__ == "__main__":
    print("Import this module and call build(data, out_path) with real data.")
