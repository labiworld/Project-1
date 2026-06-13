#!/usr/bin/env python3
"""
LandCulture Investment Ltd — Contract Web Form
===============================================
Run:
    python3 web_app.py

Then open http://localhost:5000 in your browser.

Email setup:
    Copy .env.example to .env and fill in your Gmail credentials.
"""

import os
import smtplib
import io
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from flask import Flask, render_template_string, request, send_file, flash, redirect, url_for

from generate_contract import generate_contract, PLACEHOLDERS

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "landculture-secret-2026")

# ---------------------------------------------------------------------------
# HTML template (single-file, no separate templates folder needed)
# ---------------------------------------------------------------------------

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>LandCulture — Contract Generator</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #f4f6f9;
      color: #222;
      min-height: 100vh;
    }
    header {
      background: #1a4b2e;
      color: #fff;
      padding: 18px 32px;
      display: flex;
      align-items: center;
      gap: 14px;
    }
    header h1 { font-size: 1.3rem; font-weight: 700; letter-spacing: .5px; }
    header p  { font-size: .85rem; opacity: .8; margin-top: 2px; }
    .container { max-width: 760px; margin: 36px auto; padding: 0 16px 60px; }
    .card {
      background: #fff;
      border-radius: 10px;
      box-shadow: 0 2px 12px rgba(0,0,0,.08);
      padding: 32px 36px;
    }
    h2 { font-size: 1.05rem; color: #1a4b2e; margin-bottom: 22px; border-bottom: 2px solid #e8f0eb; padding-bottom: 10px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px 24px; }
    .field { display: flex; flex-direction: column; gap: 5px; }
    .field.full { grid-column: 1 / -1; }
    label { font-size: .82rem; font-weight: 600; color: #444; text-transform: uppercase; letter-spacing: .4px; }
    input[type=text], input[type=email] {
      border: 1px solid #cdd5e0;
      border-radius: 6px;
      padding: 9px 12px;
      font-size: .95rem;
      outline: none;
      transition: border-color .2s;
      width: 100%;
    }
    input:focus { border-color: #1a4b2e; box-shadow: 0 0 0 3px rgba(26,75,46,.1); }
    .hint { font-size: .75rem; color: #888; margin-top: 2px; }
    .section-title {
      grid-column: 1 / -1;
      font-size: .8rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .6px;
      color: #1a4b2e;
      border-top: 1px solid #e8f0eb;
      padding-top: 18px;
      margin-top: 6px;
    }
    .actions { margin-top: 28px; display: flex; gap: 12px; flex-wrap: wrap; }
    .btn {
      padding: 11px 28px;
      border: none;
      border-radius: 7px;
      font-size: .95rem;
      font-weight: 600;
      cursor: pointer;
      transition: background .2s, transform .1s;
    }
    .btn:active { transform: scale(.98); }
    .btn-primary { background: #1a4b2e; color: #fff; }
    .btn-primary:hover { background: #14602e; }
    .btn-email { background: #2563eb; color: #fff; }
    .btn-email:hover { background: #1d4ed8; }
    .flash { padding: 12px 18px; border-radius: 7px; margin-bottom: 20px; font-size: .92rem; }
    .flash.success { background: #d1fae5; color: #065f46; border: 1px solid #6ee7b7; }
    .flash.error   { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
    .email-section { margin-top: 28px; padding-top: 22px; border-top: 2px solid #e8f0eb; }
    .email-section h2 { margin-bottom: 14px; }
    .email-row { display: flex; gap: 14px; align-items: flex-end; flex-wrap: wrap; }
    .email-row .field { flex: 1; min-width: 220px; }
    @media (max-width: 600px) { .grid { grid-template-columns: 1fr; } .field.full { grid-column: 1; } }
  </style>
</head>
<body>
<header>
  <div>
    <h1>LandCulture Investment Ltd</h1>
    <p>Contract of Sale Generator</p>
  </div>
</header>

<div class="container">
  {% for msg, category in messages %}
  <div class="flash {{ category }}">{{ msg }}</div>
  {% endfor %}

  <div class="card">
    <h2>Customer Details</h2>
    <form method="POST" action="/generate" id="contractForm">

      <div class="grid">

        <div class="section-title">Contract Date</div>

        <div class="field">
          <label>Day</label>
          <input type="text" name="CONTRACT_DAY" placeholder="e.g. 5th" value="{{ vals.get('CONTRACT_DAY','') }}" required />
        </div>
        <div class="field">
          <label>Month</label>
          <input type="text" name="CONTRACT_MONTH" placeholder="e.g. June" value="{{ vals.get('CONTRACT_MONTH','') }}" required />
        </div>
        <div class="field">
          <label>Year</label>
          <input type="text" name="CONTRACT_YEAR" placeholder="e.g. 2026" value="{{ vals.get('CONTRACT_YEAR','2026') }}" required />
        </div>

        <div class="section-title">Customer Information</div>

        <div class="field full">
          <label>Full Name (in CAPS)</label>
          <input type="text" name="CUSTOMER_NAME" placeholder="e.g. MR. JOHN DOE" value="{{ vals.get('CUSTOMER_NAME','') }}" required />
        </div>
        <div class="field full">
          <label>Address</label>
          <input type="text" name="CUSTOMER_ADDRESS" placeholder="e.g. 5, ALLEN AVENUE, LAGOS" value="{{ vals.get('CUSTOMER_ADDRESS','') }}" required />
        </div>

        <div class="section-title">Plot Details</div>

        <div class="field">
          <label>Number of Plots (words+digits)</label>
          <input type="text" name="NUM_PLOTS_WORDS" placeholder="e.g. Two (2)" value="{{ vals.get('NUM_PLOTS_WORDS','') }}" required />
          <span class="hint">This appears in the body text</span>
        </div>
        <div class="field">
          <label>Number of Plots (CAPS)</label>
          <input type="text" name="NUM_PLOTS_WORDS_UPPER" placeholder="e.g. TWO (2)" value="{{ vals.get('NUM_PLOTS_WORDS_UPPER','') }}" required />
          <span class="hint">Used in clauses (uppercase)</span>
        </div>
        <div class="field">
          <label>Plot Size (sqm)</label>
          <input type="text" name="PLOT_SIZE_SQM" placeholder="e.g. 900" value="{{ vals.get('PLOT_SIZE_SQM','') }}" required />
        </div>

        <div class="section-title">Payment Details</div>

        <div class="field">
          <label>Total Price (digits)</label>
          <input type="text" name="TOTAL_PRICE_DIGITS" placeholder="e.g. 3,000,000" value="{{ vals.get('TOTAL_PRICE_DIGITS','') }}" required />
        </div>
        <div class="field">
          <label>Total Price (CAPS words)</label>
          <input type="text" name="TOTAL_PRICE_WORDS" placeholder="e.g. THREE MILLION NAIRA ONLY" value="{{ vals.get('TOTAL_PRICE_WORDS','') }}" required />
        </div>
        <div class="field">
          <label>Deposit Paid (digits)</label>
          <input type="text" name="DEPOSIT_DIGITS" placeholder="e.g. 800,000.00" value="{{ vals.get('DEPOSIT_DIGITS','') }}" required />
        </div>
        <div class="field">
          <label>Deposit Paid (CAPS words)</label>
          <input type="text" name="DEPOSIT_WORDS" placeholder="e.g. EIGHT HUNDRED THOUSAND NAIRA ONLY" value="{{ vals.get('DEPOSIT_WORDS','') }}" required />
        </div>
        <div class="field">
          <label>Balance (digits)</label>
          <input type="text" name="BALANCE_DIGITS" placeholder="e.g. 2,200,000.00" value="{{ vals.get('BALANCE_DIGITS','') }}" required />
        </div>
        <div class="field">
          <label>Balance (CAPS words)</label>
          <input type="text" name="BALANCE_WORDS" placeholder="e.g. TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY" value="{{ vals.get('BALANCE_WORDS','') }}" required />
        </div>

        <div class="section-title">Payment Schedule</div>

        <div class="field">
          <label>Payment Start Date</label>
          <input type="text" name="PAYMENT_START_DATE" placeholder="e.g. 26th March, 2026" value="{{ vals.get('PAYMENT_START_DATE','') }}" required />
        </div>
        <div class="field">
          <label>Duration (months)</label>
          <input type="text" name="PAYMENT_DURATION_MONTHS" placeholder="e.g. 12" value="{{ vals.get('PAYMENT_DURATION_MONTHS','') }}" required />
        </div>
        <div class="field">
          <label>Payment End Date</label>
          <input type="text" name="PAYMENT_END_DATE" placeholder="e.g. 26th Day of March, 2027" value="{{ vals.get('PAYMENT_END_DATE','') }}" required />
        </div>
        <div class="field">
          <label>Deadline (CAPS, for clauses)</label>
          <input type="text" name="PAYMENT_START_DAY_FULL" placeholder="e.g. 26TH MARCH, 2027" value="{{ vals.get('PAYMENT_START_DAY_FULL','') }}" required />
        </div>

      </div><!-- .grid -->

      <div class="actions">
        <button type="submit" name="action" value="download" class="btn btn-primary">Download Contract (.docx)</button>
      </div>

      <!-- Email section -->
      <div class="email-section">
        <h2>Send Contract by Email (optional)</h2>
        <div class="email-row">
          <div class="field">
            <label>Customer Email</label>
            <input type="email" name="customer_email" placeholder="customer@example.com" value="{{ vals.get('customer_email','') }}" />
          </div>
          <div>
            <button type="submit" name="action" value="email" class="btn btn-email">Generate &amp; Send Email</button>
          </div>
        </div>
        {% if not email_configured %}
        <p class="hint" style="margin-top:10px; color:#b45309;">
          ⚠ Email not configured. Set GMAIL_USER and GMAIL_APP_PASSWORD in your .env file to enable sending.
        </p>
        {% endif %}
      </div>

    </form>
  </div>
</div>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _email_configured():
    return bool(os.environ.get("GMAIL_USER") and os.environ.get("GMAIL_APP_PASSWORD"))


def _send_email(to_address: str, customer_name: str, docx_bytes: bytes, filename: str):
    gmail_user = os.environ["GMAIL_USER"]
    gmail_password = os.environ["GMAIL_APP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = gmail_user
    msg["To"] = to_address
    msg["Subject"] = f"Your Contract of Sale — LandCulture Investment Ltd"

    body = f"""Dear {customer_name.title()},

Please find attached your Contract of Sale document from LandCulture Investment Ltd.

If you have any questions, please do not hesitate to contact us.

Best regards,
LandCulture Investment Ltd
Plot 6, Old Jebba Road, beside IPMAN Building, Sango, Ilorin, Kwara State
"""
    msg.attach(MIMEText(body, "plain"))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(docx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, to_address, msg.as_string())


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template_string(
        HTML,
        messages=get_flashed_messages_with_categories(),
        vals={},
        email_configured=_email_configured(),
    )


def get_flashed_messages_with_categories():
    from flask import get_flashed_messages
    return get_flashed_messages(with_categories=True)


@app.route("/generate", methods=["POST"])
def generate():
    action = request.form.get("action", "download")
    vals = {k: request.form.get(k, "").strip() for k in request.form}

    # Build data dict for generate_contract()
    field_keys = [f[0] for f in PLACEHOLDERS] + ["NUM_PLOTS_WORDS_UPPER"]
    data = {k: vals.get(k, "") for k in field_keys}

    # Generate contract and read the output bytes
    try:
        out_path = generate_contract(data)
        filename = os.path.basename(out_path)
        with open(out_path, "rb") as f:
            docx_bytes = f.read()
    except Exception as e:
        flash(f"Error generating contract: {e}", "error")
        return render_template_string(HTML, messages=get_flashed_messages_with_categories(), vals=vals, email_configured=_email_configured())

    if action == "email":
        customer_email = vals.get("customer_email", "").strip()
        if not customer_email:
            flash("Please enter a customer email address to send the contract.", "error")
            return render_template_string(HTML, messages=get_flashed_messages_with_categories(), vals=vals, email_configured=_email_configured())
        if not _email_configured():
            flash("Email is not configured. Please set GMAIL_USER and GMAIL_APP_PASSWORD in your .env file.", "error")
            return render_template_string(HTML, messages=get_flashed_messages_with_categories(), vals=vals, email_configured=_email_configured())
        try:
            _send_email(customer_email, data.get("CUSTOMER_NAME", "Customer"), docx_bytes, filename)
            flash(f"Contract successfully sent to {customer_email}!", "success")
        except Exception as e:
            flash(f"Email failed: {e}", "error")
        return render_template_string(HTML, messages=get_flashed_messages_with_categories(), vals=vals, email_configured=_email_configured())

    # Default: download
    return send_file(
        io.BytesIO(docx_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------------------------------------------------------
# Load .env if present
# ---------------------------------------------------------------------------

def _load_dotenv():
    env_file = os.path.join(os.path.dirname(__file__), ".env")
    if not os.path.exists(env_file):
        return
    with open(env_file) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip())


if __name__ == "__main__":
    _load_dotenv()
    print("\n=== LandCulture Contract Web Form ===")
    print("Open your browser at: http://localhost:5000\n")
    app.run(debug=False, host="0.0.0.0", port=5000)
