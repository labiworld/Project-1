#!/usr/bin/env python3
"""
LandCulture Investment Ltd — HubSpot-Connected Contract Dashboard
=================================================================
Run:
    python3 hubspot_app.py

Then open http://localhost:5001 in your browser.

Setup:
    Copy .env.example to .env and set HUBSPOT_TOKEN
"""

import os, io, json
from datetime import datetime

import requests
from flask import Flask, render_template_string, request, send_file, jsonify

from generate_contract import generate_contract

app = Flask(__name__)

# ---------------------------------------------------------------------------
# HubSpot helpers
# ---------------------------------------------------------------------------

def hs_headers():
    token = os.environ.get("HUBSPOT_TOKEN", "")
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_contacts(limit=30):
    """Fetch the 30 most recently created contacts from HubSpot."""
    props = ["firstname", "lastname", "email", "phone", "address", "city", "state"]
    payload = {
        "limit": limit,
        "properties": props,
        "sorts": [{"propertyName": "createdate", "direction": "DESCENDING"}],
    }
    r = requests.post(
        "https://api.hubapi.com/crm/v3/objects/contacts/search",
        headers=hs_headers(),
        json=payload,
    )
    r.raise_for_status()
    return r.json().get("results", [])


def get_contact(contact_id):
    """Fetch a single contact with all properties."""
    r = requests.get(
        f"https://api.hubapi.com/crm/v3/objects/contacts/{contact_id}?properties=firstname,lastname,email,phone,address,city,state",
        headers=hs_headers()
    )
    r.raise_for_status()
    return r.json()


def get_deals_for_contact(contact_id):
    """Fetch deals associated with a contact."""
    r = requests.get(
        f"https://api.hubapi.com/crm/v3/objects/contacts/{contact_id}/associations/deals",
        headers=hs_headers()
    )
    if r.status_code != 200:
        return []
    deal_ids = [a["id"] for a in r.json().get("results", [])]
    deals = []
    props = ",".join([
        "dealname", "closedate",
        # Standard HubSpot field
        "amount",
        # Custom fields (exact internal names from HubSpot)
        "total_balance_remaining", "balance_digits", "balance_words",
        "plot_size", "plot_size_sqm",
        "number_of_plot", "num_plots_words", "num_plots_upper",
        "installmental_plan", "payment_duration_months",
        "total_price_digits", "total_price_words",
        "deposit_digits", "deposit_words",
        "payment_start_date", "payment_end_date", "payment_start_day_full",
    ])
    for did in deal_ids[:5]:
        dr = requests.get(
            f"https://api.hubapi.com/crm/v3/objects/deals/{did}?properties={props}",
            headers=hs_headers()
        )
        if dr.status_code == 200:
            deals.append(dr.json())
    return deals


def get_all_deal_properties():
    """Fetch all available deal properties from HubSpot (for field mapping)."""
    r = requests.get(
        "https://api.hubapi.com/crm/v3/properties/deals",
        headers=hs_headers()
    )
    if r.status_code == 200:
        return {p["name"]: p["label"] for p in r.json().get("results", [])}
    return {}


# ---------------------------------------------------------------------------
# Map HubSpot fields → contract fields
# ---------------------------------------------------------------------------

def build_contract_data(contact, deal=None):
    """Build the contract data dict from HubSpot contact + deal."""
    cp = contact.get("properties", {})
    dp = deal.get("properties", {}) if deal else {}

    def g(d, *keys, default=""):
        for k in keys:
            v = d.get(k)
            if v: return str(v).upper()
        return default

    # Customer name
    first = g(cp, "firstname")
    last  = g(cp, "lastname")
    full_name = f"{first} {last}".strip().upper()

    # Address
    address_parts = [cp.get("address",""), cp.get("city",""), cp.get("state","")]
    address = ", ".join(p for p in address_parts if p).upper()

    num_plots = g(dp, "number_of_plot", "num_plots_words")

    return {
        "CUSTOMER_NAME":           g(dp, "customer_name_full") or full_name,
        "CUSTOMER_ADDRESS":        g(dp, "customer_address") or address,
        # "Number of Plot" in HubSpot → also auto-derive UPPER version
        "NUM_PLOTS_WORDS":         num_plots,
        "NUM_PLOTS_WORDS_UPPER":   num_plots.upper() if num_plots else "",
        "NUM_PLOTS_DIGITS":        num_plots,
        # "Plot Size" in HubSpot
        "PLOT_SIZE_SQM":           g(dp, "plot_size", "plot_size_sqm"),
        # "Amount" in HubSpot = Total Price
        "TOTAL_PRICE_DIGITS":      g(dp, "amount", "total_price_digits"),
        "TOTAL_PRICE_WORDS":       g(dp, "total_price_words"),
        "DEPOSIT_DIGITS":          g(dp, "deposit_digits"),
        "DEPOSIT_WORDS":           g(dp, "deposit_words"),
        # Balance — not yet in HubSpot, will show as missing (yellow)
        "BALANCE_DIGITS":          g(dp, "total_balance_remaining", "balance_digits", "balance"),
        "BALANCE_WORDS":           g(dp, "balance_words"),
        "PAYMENT_START_DATE":      g(dp, "payment_start_date"),
        # Installment plan — not yet in HubSpot, will show as missing (yellow)
        "PAYMENT_DURATION_MONTHS": g(dp, "installmental_plan", "installment_plan", "payment_duration_months"),
        "PAYMENT_END_DATE":        g(dp, "payment_end_date"),
        "PAYMENT_START_DAY_FULL":  g(dp, "payment_start_day_full"),
    }


# ---------------------------------------------------------------------------
# HTML Dashboard
# ---------------------------------------------------------------------------

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>LandCulture — Legal Dashboard</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; color: #222; }
    header { background: #1a4b2e; color: #fff; padding: 16px 32px; display: flex; align-items: center; justify-content: space-between; }
    header h1 { font-size: 1.2rem; font-weight: 700; }
    header p  { font-size: .82rem; opacity: .75; }
    .container { max-width: 1100px; margin: 32px auto; padding: 0 16px 60px; }
    .top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
    .top-bar h2 { font-size: 1rem; color: #1a4b2e; }
    .btn-refresh { background: #1a4b2e; color: #fff; border: none; padding: 9px 20px; border-radius: 6px; cursor: pointer; font-size: .88rem; font-weight: 600; }
    .btn-refresh:hover { background: #145c2e; }
    .card { background: #fff; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.07); overflow: hidden; }
    table { width: 100%; border-collapse: collapse; }
    th { background: #f0f4f1; text-align: left; padding: 12px 16px; font-size: .78rem; text-transform: uppercase; letter-spacing: .5px; color: #555; border-bottom: 1px solid #e0e8e2; }
    td { padding: 12px 16px; font-size: .9rem; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #f8fdf9; }
    .btn-generate { background: #1a4b2e; color: #fff; border: none; padding: 7px 16px; border-radius: 6px; cursor: pointer; font-size: .82rem; font-weight: 600; text-decoration: none; }
    .btn-generate:hover { background: #145c2e; }
    .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: .75rem; font-weight: 600; background: #d1fae5; color: #065f46; }
    .badge.no-deal { background: #fee2e2; color: #991b1b; }
    .empty { text-align: center; padding: 60px; color: #888; }
    .error-box { background: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; padding: 14px 20px; border-radius: 8px; margin-bottom: 20px; }
    /* Mobile cards */
    .contact-cards { display: none; }
    .contact-card { background: #fff; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.07); padding: 16px; margin-bottom: 12px; }
    .contact-card .name { font-weight: 700; font-size: 1rem; margin-bottom: 4px; }
    .contact-card .meta { font-size: .85rem; color: #555; margin-bottom: 12px; }
    .contact-card .btn-generate { display: block; width: 100%; text-align: center; padding: 12px; font-size: 1rem; }
    @media (max-width: 600px) {
      header { padding: 14px 16px; }
      .card { display: none; }
      .contact-cards { display: block; }
      .modal { width: 100%; max-width: 100%; max-height: 100vh; border-radius: 0; }
      .field-row { grid-template-columns: 1fr; }
    }

    /* Modal */
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-overlay.active { display: flex; }
    .modal { background: #fff; border-radius: 12px; width: 90%; max-width: 620px; max-height: 90vh; overflow-y: auto; box-shadow: 0 20px 60px rgba(0,0,0,.3); }
    .modal-header { background: #1a4b2e; color: #fff; padding: 18px 24px; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center; }
    .modal-header h3 { font-size: 1rem; }
    .modal-close { background: none; border: none; color: #fff; font-size: 1.4rem; cursor: pointer; line-height: 1; }
    .modal-body { padding: 24px; }
    .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px; }
    .field { display: flex; flex-direction: column; gap: 4px; }
    .field.full { grid-column: 1 / -1; }
    .field label { font-size: .75rem; font-weight: 700; text-transform: uppercase; color: #555; }
    .field input[type=text] { padding: 8px 10px; background: #fff; border: 1px solid #cdd5e0; border-radius: 6px; font-size: .88rem; color: #222; width: 100%; outline: none; transition: border-color .2s; }
    .field input[type=text]:focus { border-color: #1a4b2e; box-shadow: 0 0 0 3px rgba(26,75,46,.1); }
    .field input[type=text].missing { background: #fff3cd; border-color: #ffc107; }
    .field input[type=text].missing:focus { border-color: #1a4b2e; background: #fff; }
    .modal-footer { padding: 16px 24px; border-top: 1px solid #eee; display: flex; gap: 12px; justify-content: flex-end; }
    .btn-cancel { background: #e5e7eb; color: #444; border: none; padding: 10px 22px; border-radius: 7px; cursor: pointer; font-weight: 600; }
    .btn-dl { background: #1a4b2e; color: #fff; border: none; padding: 10px 22px; border-radius: 7px; cursor: pointer; font-weight: 600; font-size: .95rem; }
    .btn-dl:hover { background: #145c2e; }
    .section-label { font-size: .75rem; font-weight: 700; text-transform: uppercase; color: #1a4b2e; border-bottom: 1px solid #e8f0eb; padding-bottom: 6px; margin-bottom: 12px; margin-top: 4px; grid-column: 1 / -1; }
  </style>
</head>
<body>
<header>
  <div>
    <h1>LandCulture Investment Ltd — Legal Dashboard</h1>
    <p>HubSpot-connected contract generator</p>
  </div>
</header>

<div class="container">
  {% if error %}
  <div class="error-box">⚠ HubSpot connection error: {{ error }}</div>
  {% endif %}

  <div class="top-bar">
    <h2>Recent Customers ({{ contacts|length }})</h2>
    <button class="btn-refresh" onclick="location.reload()">↻ Refresh</button>
  </div>

  <!-- Desktop table -->
  <div class="card">
    {% if contacts %}
    <table>
      <thead>
        <tr>
          <th>Customer Name</th>
          <th>Email</th>
          <th>Phone</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
        {% for c in contacts %}
        <tr>
          <td><strong>{{ c.name }}</strong></td>
          <td>{{ c.email or '—' }}</td>
          <td>{{ c.phone or '—' }}</td>
          <td>
            <button class="btn-generate" onclick="openModal('{{ c.id }}', '{{ c.name }}')">
              Generate COS
            </button>
          </td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
    {% else %}
    <div class="empty">No contacts found in HubSpot.</div>
    {% endif %}
  </div>

  <!-- Mobile cards -->
  <div class="contact-cards">
    {% if contacts %}
      {% for c in contacts %}
      <div class="contact-card">
        <div class="name">{{ c.name }}</div>
        <div class="meta">{{ c.email or '—' }} &nbsp;·&nbsp; {{ c.phone or '—' }}</div>
        <button class="btn-generate" onclick="openModal('{{ c.id }}', '{{ c.name }}')">
          Generate COS
        </button>
      </div>
      {% endfor %}
    {% else %}
    <div class="empty">No contacts found in HubSpot.</div>
    {% endif %}
  </div>
</div>

<!-- Modal -->
<div class="modal-overlay" id="modal">
  <div class="modal">
    <div class="modal-header">
      <h3 id="modal-title">Contract of Sale</h3>
      <button class="modal-close" onclick="closeModal()">×</button>
    </div>
    <div class="modal-body" id="modal-body">
      <p>Loading...</p>
    </div>
    <div class="modal-footer">
      <button class="btn-cancel" onclick="closeModal()">Cancel</button>
      <button class="btn-dl" id="btn-download" onclick="downloadContract()">Generate &amp; Download COS</button>
    </div>
  </div>
</div>

<script>
let currentContactId = null;
let contractData = null;

function openModal(contactId, name) {
  currentContactId = contactId;
  document.getElementById('modal-title').textContent = 'Contract of Sale — ' + name;
  document.getElementById('modal-body').innerHTML = '<p style="text-align:center;padding:40px;color:#888;">Loading HubSpot data...</p>';
  document.getElementById('modal').classList.add('active');

  fetch('/api/contact/' + contactId)
    .then(r => r.json())
    .then(data => {
      contractData = data;
      renderModal(data);
    })
    .catch(e => {
      document.getElementById('modal-body').innerHTML = '<p style="color:red">Error loading data: ' + e + '</p>';
    });
}

function renderModal(data) {
  const fields = [
    { key: 'CUSTOMER_NAME',           label: 'Customer Name',           section: 'Customer' },
    { key: 'CUSTOMER_ADDRESS',        label: 'Address',                  section: null },
    { key: 'NUM_PLOTS_WORDS',         label: 'No. of Plots (words)',     section: 'Plot Details' },
    { key: 'NUM_PLOTS_WORDS_UPPER',   label: 'No. of Plots (CAPS)',      section: null },
    { key: 'PLOT_SIZE_SQM',           label: 'Plot Size (sqm)',          section: null },
    { key: 'TOTAL_PRICE_DIGITS',      label: 'Total Price (digits)',     section: 'Payment' },
    { key: 'TOTAL_PRICE_WORDS',       label: 'Total Price (words)',      section: null },
    { key: 'DEPOSIT_DIGITS',          label: 'Deposit (digits)',         section: null },
    { key: 'DEPOSIT_WORDS',           label: 'Deposit (words)',          section: null },
    { key: 'BALANCE_DIGITS',          label: 'Balance (digits)',         section: null },
    { key: 'BALANCE_WORDS',           label: 'Balance (words)',          section: null },
    { key: 'PAYMENT_START_DATE',      label: 'Payment Start Date',       section: 'Schedule' },
    { key: 'PAYMENT_DURATION_MONTHS', label: 'Duration (months)',        section: null },
    { key: 'PAYMENT_END_DATE',        label: 'Payment End Date',         section: null },
    { key: 'PAYMENT_START_DAY_FULL',  label: 'Deadline (CAPS)',          section: null },
  ];

  let html = '<div class="field-row">';
  fields.forEach(f => {
    const val = data[f.key] || '';
    const missing = !val;
    if (f.section) {
      html += `<div class="section-label">${f.section}</div>`;
    }
    html += `<div class="field">
      <label>${f.label}${missing ? ' <span style="color:#b45309;font-weight:400;text-transform:none;">⚠ missing</span>' : ''}</label>
      <input type="text" id="field_${f.key}" name="${f.key}" value="${val}" placeholder="Enter ${f.label.toLowerCase()}" class="${missing ? 'missing' : ''}" oninput="this.classList.remove('missing')" />
    </div>`;
  });
  html += '</div>';
  document.getElementById('modal-body').innerHTML = html;
}

function closeModal() {
  document.getElementById('modal').classList.remove('active');
  currentContactId = null;
  contractData = null;
}

function downloadContract() {
  if (!contractData) return;

  // Read current values from the editable inputs
  const payload = {};
  document.querySelectorAll('#modal-body input[type=text]').forEach(inp => {
    payload[inp.name] = inp.value.trim();
  });

  const btn = document.getElementById('btn-download');
  btn.textContent = 'Generating...';
  btn.disabled = true;

  fetch('/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  })
  .then(r => {
    if (!r.ok) return r.text().then(t => { throw new Error(t); });
    return r.blob();
  })
  .then(blob => {
    const name = (payload.CUSTOMER_NAME || 'CUSTOMER').replace(/[\s]+/g, '_').replace(/[.]/g, '');
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `CONTRACT_${name}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    btn.textContent = 'Generate & Download COS';
    btn.disabled = false;
    closeModal();
  })
  .catch(e => {
    alert('Error: ' + e.message);
    btn.textContent = 'Generate & Download COS';
    btn.disabled = false;
  });
}

// Close modal on overlay click
document.getElementById('modal').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    contacts = []
    error = None
    try:
        raw = get_contacts(30)
        for c in raw:
            p = c.get("properties", {})
            first = p.get("firstname") or ""
            last  = p.get("lastname") or ""
            contacts.append({
                "id":    c["id"],
                "name":  f"{first} {last}".strip() or f"Contact {c['id']}",
                "email": p.get("email", ""),
                "phone": p.get("phone", ""),
            })
    except Exception as e:
        error = str(e)

    return render_template_string(DASHBOARD_HTML, contacts=contacts, error=error)


@app.route("/api/contact/<contact_id>")
def api_contact(contact_id):
    try:
        contact = get_contact(contact_id)
        deals   = get_deals_for_contact(contact_id)
        deal    = deals[0] if deals else None
        data    = build_contract_data(contact, deal)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/debug-deal/<contact_id>")
def api_debug_deal(contact_id):
    """Return raw HubSpot deal properties so we can see exact internal names."""
    try:
        deals = get_deals_for_contact(contact_id)
        if not deals:
            return jsonify({"error": "No deals found for this contact"})
        raw_props = deals[0].get("properties", {})
        # Filter out empty/null values and HubSpot system fields
        filtered = {k: v for k, v in raw_props.items() if v and not k.startswith("hs_")}
        return jsonify(filtered)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def api_generate():
    try:
        data = request.get_json()
        out_path = generate_contract(data)
        with open(out_path, "rb") as f:
            docx_bytes = f.read()
        customer_safe = (data.get("CUSTOMER_NAME", "CUSTOMER")
                         .replace(" ", "_").replace(".", "").replace("/", "_"))
        filename = f"CONTRACT_{customer_safe}.docx"
        return send_file(
            io.BytesIO(docx_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return str(e), 500


# ---------------------------------------------------------------------------
# Load .env and run
# ---------------------------------------------------------------------------

def _load_dotenv():
    env_file = os.path.join(os.path.dirname(__file__), ".env")
    if not os.path.exists(env_file): return
    with open(env_file) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                os.environ.setdefault(k.strip(), v.strip())


if __name__ == "__main__":
    _load_dotenv()
    if not os.environ.get("HUBSPOT_TOKEN"):
        print("ERROR: Set HUBSPOT_TOKEN in your .env file")
    else:
        port = int(os.environ.get("PORT", 5001))
        print(f"\n=== LandCulture Legal Dashboard ===")
        print(f"Open your browser at: http://localhost:{port}\n")
        app.run(debug=False, host="0.0.0.0", port=port)
