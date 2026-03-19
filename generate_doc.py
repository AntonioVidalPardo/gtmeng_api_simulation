"""Generate API documentation as a .docx file."""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# -- Styles --
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)

# -- Title --
title = doc.add_heading("GTM Engineering API – Documentation", level=0)
title.runs[0].font.size = Pt(24)

doc.add_paragraph(
    "Static JSON API for email outreach account and campaign analytics.\n"
    "Served via GitHub Pages — no authentication required."
)

# -- Base URL --
p = doc.add_paragraph()
run = p.add_run("Base URL:  ")
run.bold = True
run = p.add_run("https://antoniovidalpardo.github.io/gtmeng_api_simulation/")
run.font.color.rgb = RGBColor(0, 102, 204)

p = doc.add_paragraph()
run = p.add_run("Note: ")
run.bold = True
p.add_run(
    "This is a static API. All endpoints return the full dataset as a JSON array. "
    "Filtering, sorting, and pagination should be done client-side after fetching."
)

# ---------- helper ----------
def add_endpoint_section(heading, path, description, fields, example):
    doc.add_heading(heading, level=1)

    p = doc.add_paragraph()
    run = p.add_run("GET  ")
    run.bold = True
    run.font.color.rgb = RGBColor(22, 163, 74)
    run = p.add_run(path)
    run.bold = True
    run.font.name = "Consolas"
    run.font.size = Pt(11)

    doc.add_paragraph(description)

    # Schema table
    doc.add_heading("Response Schema", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    hdr = table.rows[0].cells
    for i, text in enumerate(["Field", "Type", "Description"]):
        hdr[i].text = text
        for par in hdr[i].paragraphs:
            for run in par.runs:
                run.bold = True

    for field_name, field_type, field_desc in fields:
        row = table.add_row().cells
        row[0].text = field_name
        for par in row[0].paragraphs:
            for run in par.runs:
                run.font.name = "Consolas"
                run.font.size = Pt(10)
        row[1].text = field_type
        row[2].text = field_desc

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(0.8)
        row.cells[2].width = Inches(3.7)

    # Example response
    doc.add_heading("Example Response", level=2)
    p = doc.add_paragraph()
    run = p.add_run(example)
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    p.paragraph_format.space_before = Pt(4)


# ---------- Endpoint 1: Accounts ----------
add_endpoint_section(
    heading="1. Account Summaries",
    path="/api/accounts.json",
    description=(
        "Returns aggregated performance summaries for all 12 email sending accounts, "
        "including lifetime totals and calculated engagement rates."
    ),
    fields=[
        ("account", "string", "Email address of the sending account"),
        ("owner", "string", "Name of the account owner"),
        ("tags", "string[]", 'Tags assigned to the account (e.g. "antonio; batch_1")'),
        ("total_emails_sent", "integer", "Total emails sent across all dates"),
        ("total_emails_opened", "integer", "Total emails opened"),
        ("total_replies", "integer", "Total replies received"),
        ("first_date", "string", "First active date (M/D/YYYY format)"),
        ("last_date", "string", "Most recent active date (M/D/YYYY format)"),
        ("days_active", "integer", "Number of days with sending activity"),
        ("open_rate", "number", "Open rate as a percentage (e.g. 41.2 = 41.2%)"),
        ("reply_rate", "number", "Reply rate as a percentage (e.g. 14.7 = 14.7%)"),
    ],
    example="""\
[
  {
    "account": "antonio@getgtmeng.com",
    "owner": "antonio",
    "tags": ["antonio; batch_1"],
    "total_emails_sent": 68,
    "total_emails_opened": 28,
    "total_replies": 10,
    "first_date": "2/20/2026",
    "last_date": "2/24/2026",
    "days_active": 10,
    "open_rate": 41.2,
    "reply_rate": 14.7
  },
  ...
]""",
)

# ---------- Endpoint 2: Account Analytics ----------
add_endpoint_section(
    heading="2. Daily Account Analytics",
    path="/api/account-analytics.json",
    description=(
        "Returns daily email performance metrics broken down by individual sending account. "
        "Contains 184 records across 12 accounts, covering Feb 18 – Mar 9, 2026."
    ),
    fields=[
        ("account", "string", "Email address of the sending account"),
        ("tag", "string", 'Tag associated with this account (e.g. "antonio; batch_1")'),
        ("date", "string", "Date of the record (M/D/YYYY format)"),
        ("emails_sent", "integer", "Number of emails sent on this date"),
        ("emails_opened", "integer", "Number of emails opened on this date"),
        ("replies", "integer", "Number of replies received on this date"),
    ],
    example="""\
[
  {
    "account": "antonio@getgtmengcourse.com",
    "tag": "antonio; batch_1",
    "date": "2/18/2026",
    "emails_sent": 5,
    "emails_opened": 1,
    "replies": 0
  },
  ...
]""",
)

# ---------- Endpoint 3: Campaign Analytics ----------
add_endpoint_section(
    heading="3. Daily Campaign Analytics",
    path="/api/campaign-analytics.json",
    description=(
        "Returns daily performance metrics for email campaigns. "
        "Contains 140 records across 3 campaigns (campaign_1, campaign_2, campaign_3), "
        "covering Feb 18 – Mar 9, 2026."
    ),
    fields=[
        ("campaign_name", "string", "Name of the campaign (campaign_1, campaign_2, or campaign_3)"),
        ("date", "string", "Date of the record (M/D/YYYY format)"),
        ("emails_sent", "integer", "Number of emails sent on this date"),
        ("emails_opened", "integer", "Number of emails opened on this date"),
        ("positive_replies", "integer", "Number of positive replies received on this date"),
    ],
    example="""\
[
  {
    "campaign_name": "campaign_1",
    "date": "2/18/2026",
    "emails_sent": 100,
    "emails_opened": 80,
    "positive_replies": 5
  },
  ...
]""",
)

# ---------- Usage Examples ----------
doc.add_heading("Usage Examples", level=1)

doc.add_heading("JavaScript (fetch)", level=2)
p = doc.add_paragraph()
run = p.add_run("""\
const BASE = "https://antoniovidalpardo.github.io/gtmeng_api_simulation";

// Get all account summaries
const accounts = await fetch(`${BASE}/api/accounts.json`).then(r => r.json());

// Get daily analytics for a specific account
const analytics = await fetch(`${BASE}/api/account-analytics.json`).then(r => r.json());
const antonioData = analytics.filter(r => r.account === "antonio@getgtmeng.com");

// Get campaign analytics for a specific campaign
const campaigns = await fetch(`${BASE}/api/campaign-analytics.json`).then(r => r.json());
const campaign1 = campaigns.filter(r => r.campaign_name === "campaign_1");""")
run.font.name = "Consolas"
run.font.size = Pt(9)

doc.add_heading("Python (requests)", level=2)
p = doc.add_paragraph()
run = p.add_run("""\
import requests

BASE = "https://antoniovidalpardo.github.io/gtmeng_api_simulation"

# Get all account summaries
accounts = requests.get(f"{BASE}/api/accounts.json").json()

# Get daily analytics filtered by account
analytics = requests.get(f"{BASE}/api/account-analytics.json").json()
antonio_data = [r for r in analytics if r["account"] == "antonio@getgtmeng.com"]

# Get campaign analytics filtered by date
campaigns = requests.get(f"{BASE}/api/campaign-analytics.json").json()
march_data = [r for r in campaigns if r["date"].startswith("3/")]""")
run.font.name = "Consolas"
run.font.size = Pt(9)

doc.add_heading("cURL", level=2)
p = doc.add_paragraph()
run = p.add_run("""\
# Fetch account summaries
curl -s https://antoniovidalpardo.github.io/gtmeng_api_simulation/api/accounts.json | jq .

# Fetch daily account analytics
curl -s https://antoniovidalpardo.github.io/gtmeng_api_simulation/api/account-analytics.json | jq .

# Fetch campaign analytics
curl -s https://antoniovidalpardo.github.io/gtmeng_api_simulation/api/campaign-analytics.json | jq .""")
run.font.name = "Consolas"
run.font.size = Pt(9)

# ---------- Save ----------
output_path = "/home/user/gtmeng_api_simulation/GTM_Engineering_API_Documentation.docx"
doc.save(output_path)
print(f"Saved to {output_path}")
