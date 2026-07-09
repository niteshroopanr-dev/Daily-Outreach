"""
ProfitPulse nightly outreach HTML dossier builder, V3.3.
Single self contained HTML file for internal review only. Brand colours only,
zero dashes. Data driven: pass a dict matching the DATA shape at the bottom.
"""

BLACK = "#000000"
TEAL = "#01A296"
AMBER_B = "#F8C806"
AMBER_D = "#F6A102"
GOLD = "#E3A712"
WHITE = "#FFFFFF"
OFF_WHITE = "#E6E5DE"


def render(data):
    stat_rows = "".join(
        f"""<tr><td style="padding:8px 12px;border-bottom:1px solid #eee;color:{BLACK};">{f['label']}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;color:{BLACK};font-weight:700;">{f['value']}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;color:{TEAL};font-size:12px;">{f['source']}</td></tr>"""
        for f in data["verified_facts"]
    )

    signal_items = "".join(f"<li style='margin-bottom:8px;'>{s}</li>" for s in data["signals"])

    support_services = "".join(f"<li>{s}</li>" for s in data["supporting_services"])

    contact_rows = "".join(
        f"""<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">{c['address']}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;">{c['status']}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;font-size:12px;color:#333;">{c['basis']}</td></tr>"""
        for c in data["contact_candidates"]
    )

    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ProfitPulse Outreach Dossier, {data['company']}, {data['date_str']}</title>
</head>
<body style="margin:0;padding:0;background:{WHITE};font-family:Calibri,Arial,sans-serif;color:{BLACK};">

<div style="background:{BLACK};color:{WHITE};padding:28px 32px;">
  <div style="color:{TEAL};font-size:12px;font-weight:700;letter-spacing:1px;">PROFITPULSE INTERNAL REVIEW DOSSIER</div>
  <h1 style="margin:6px 0 4px 0;font-size:32px;color:{AMBER_B};">{data['company']}</h1>
  <div style="font-size:14px;color:{OFF_WHITE};">{data['date_str']}</div>
  <div style="font-size:14px;color:{OFF_WHITE};margin-top:10px;max-width:800px;">{data['selection_reason']}</div>
</div>

<div style="max-width:960px;margin:0 auto;padding:28px 24px;">

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;">Verified profile</h2>
  <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:24px;">
    <tr style="background:{OFF_WHITE};"><th style="text-align:left;padding:8px 12px;">Field</th><th style="text-align:left;padding:8px 12px;">Value</th><th style="text-align:left;padding:8px 12px;">Basis</th></tr>
    {stat_rows}
  </table>

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;">Key commercial signals</h2>
  <ul style="font-size:13px;line-height:1.5;">{signal_items}</ul>

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;">The wedge and the matched service</h2>
  <p style="font-size:14px;line-height:1.6;">{data['wedge']}</p>

  <div style="background:{BLACK};color:{WHITE};padding:20px 24px;border-radius:6px;margin:16px 0;">
    <div style="color:{AMBER_B};font-weight:700;font-size:18px;">{data['service_name']}</div>
    <div style="color:{OFF_WHITE};font-size:13px;margin:6px 0;">{data['service_desc']}</div>
    <div style="color:{TEAL};font-weight:700;font-size:16px;">{data['service_price_line']}, a figure the questionnaire confirms for your business</div>
    <div style="color:{OFF_WHITE};font-size:12px;margin-top:10px;">Supporting services considered: <ul style="margin:6px 0 0 0;">{support_services}</ul></div>
  </div>

  <div style="text-align:center;margin:28px 0;">
    <a href="{data['questionnaire_url_tagged']}" style="display:inline-block;background:{TEAL};color:{WHITE};font-weight:700;text-decoration:none;padding:16px 28px;border-radius:30px;font-family:sans-serif;font-size:16px;">Find your fit in two minutes</a>
    <div style="color:{BLACK};font-size:14px;margin-top:8px;">Answer a few quick questions and see the solutions matched to your size and industry.</div>
    <div style="color:{TEAL};font-size:13px;margin-top:4px;">profit-pulse.com.au/services/find-your-fit</div>
  </div>

  <div style="text-align:center;margin:20px 0 32px 0;">
    <div style="font-size:13px;color:{BLACK};margin-bottom:8px;">Already know this is the priority? You can begin with {data['service_name']}, {data['service_price_line']}.</div>
    <a href="{data['stripe_link']}" style="display:inline-block;border:1px solid {AMBER_D};color:{AMBER_D};font-weight:700;text-decoration:none;padding:12px 22px;border-radius:26px;font-family:sans-serif;font-size:15px;">Purchase the suggested product now to get started</a>
  </div>

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;">Likely contact email addresses, internal use only</h2>
  <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:24px;">
    <tr style="background:{OFF_WHITE};"><th style="text-align:left;padding:6px 10px;">Address</th><th style="text-align:left;padding:6px 10px;">Status</th><th style="text-align:left;padding:6px 10px;">Basis</th></tr>
    {contact_rows}
  </table>

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;">Email draft</h2>
  <div style="background:{AMBER_D};color:{BLACK};padding:14px 18px;font-weight:700;border-radius:4px;margin-bottom:16px;">
    REVIEW GATE: This email is a draft only. No communication has been sent. Only Nitesh Roopa may authorise and send this message.
  </div>
  <div style="border:1px solid #ddd;padding:20px;border-radius:6px;">
    <div style="font-size:13px;color:#555;margin-bottom:10px;">To: {data['email_to_placeholder']} (placeholder, to be confirmed by Nitesh before any send)</div>
    <div style="font-size:13px;color:#555;margin-bottom:14px;">Subject: {data['email_subject']}</div>
    <div style="font-size:14px;line-height:1.7;white-space:pre-wrap;">{data['email_body']}</div>
  </div>

  <h2 style="color:{TEAL};font-size:16px;text-transform:uppercase;letter-spacing:1px;margin-top:28px;">System notes</h2>
  <div style="font-size:12px;line-height:1.7;background:{OFF_WHITE};padding:16px;border-radius:6px;white-space:pre-wrap;">{data['system_notes']}</div>

</div>
</body>
</html>
"""
    return html
