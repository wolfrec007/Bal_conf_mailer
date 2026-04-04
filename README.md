# Balance Confirmation Mailer

A Streamlit application for preparing and sending bulk balance confirmation emails from an Excel file.

The app supports two delivery modes:

- Desktop Outlook app via `win32com` (save as drafts or send immediately)
- SMTP providers (Gmail and Office 365/Outlook SMTP)

It is designed for AR/AP balance confirmation workflows where each recipient receives a personalized email generated from placeholders.

## Features

- Upload an Excel sheet and auto-read party records from `Party List`
- Inline data correction before generating emails
- AR/AP filtering so you can send only selected ledger types
- Template-based subject and body with dynamic placeholders
- Per-email review and manual edits before dispatch
- Multiple sending engines:
	- Desktop Outlook app (draft/save or send)
	- Gmail SMTP
	- Office 365 SMTP
- Dispatch progress, summary metrics, and detailed result log

## Project Structure

- `conf_email.py`: Main Streamlit app
- `README.md`: Documentation

## Requirements

- Python 3.10+
- Windows is required for Desktop Outlook integration (`win32com`)
- Any OS can run SMTP mode (Gmail/Office 365), subject to network and provider access rules

Python packages used by the app:

- `streamlit`
- `openpyxl`
- `pandas`
- `pywin32` (only needed when using Desktop Outlook mode)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/wolfrec007/Bal_conf_mailer.git
cd Bal_conf_mailer
```

2. (Recommended) Create and activate a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

3. Install dependencies:

```bash
pip install streamlit openpyxl pandas pywin32
```

If you do not plan to use Desktop Outlook integration, `pywin32` is optional.

## Run the App

From the project root:

```bash
streamlit run conf_email.py
```

Streamlit will print a local URL (typically `http://localhost:8501`). Open it in your browser.

## Input Excel Format

The app expects an Excel workbook with a worksheet named exactly:

- `Party List`

Header behavior:

- The app reads headers from row 2
- Data rows start from row 3

Minimum required columns per row:

- `Party Name`
- `Email ID`

Recommended columns (used in placeholders and formatting):

- `Contact Person`
- `Type (AR/AP)`
- `Currency`
- `Outstanding Balance`
- `Due Date`
- `Reference / Invoice No.`
- `Remarks`

Rows without `Party Name` or `Email ID` are skipped.

## End-to-End Workflow

### 1) Global Configuration

Configure sender and context details:

- Company name
- Signatory
- Reply-to email
- Confirmation date
- Sending provider
- Optional global CC
- AR/AP type filters

Provider options:

- Desktop Outlook App (`win32com`)
	- Action: Save as Drafts or Send Immediately
- Gmail SMTP
	- Requires sender email and app password
- Office 365 SMTP
	- Requires sender email and password/app password per org policy

### 2) Upload and Edit Data

- Upload your Excel file
- Review table in the data editor
- Correct names/emails/amounts directly in-app before generation

### 3) Build Template

- Edit subject and body templates
- Click Generate Emails
- The app creates one draft per eligible row based on filter + valid email

### 4) Review Drafts

Per recipient, you can:

- Include/exclude from dispatch
- Edit To/CC
- Edit subject/body text

### 5) Dispatch

- Execute send or draft save
- Track progress live
- Review success/failure metrics and dispatch log

## Placeholder Reference

You can use these placeholders in both subject and body templates:

- `{{party_name}}`
- `{{contact}}`
- `{{email}}`
- `{{type}}`
- `{{currency}}`
- `{{amount}}`
- `{{due_date}}`
- `{{reference}}`
- `{{remarks}}`
- `{{conf_date}}`
- `{{company}}`
- `{{signatory}}`
- `{{reply_to}}`
- `{{flow}}`

Notes:

- `{{amount}}` is currency-formatted
- `{{flow}}` resolves based on type:
	- AP -> "payable to your organisation"
	- AR -> "receivable from your organisation"

## Default Subject and Body

The app includes a built-in default template and allows full customization before generation.

Default subject pattern:

- `Balance Confirmation as on {{conf_date}} – {{party_name}} [Ref: {{reference}}]`

## Email Validation and Filtering Rules

- Basic email regex validation is applied before draft generation
- Only rows matching selected AR/AP filter are included
- Invalid email rows are ignored during generation

## Sending Notes

### Desktop Outlook Mode

- Requires Microsoft Outlook desktop app installed and configured
- Uses local Outlook profile/session
- Can save as draft or send immediately
- Most suitable for Windows environments

### SMTP Mode

- Sends immediately (no draft mode)
- Supports Gmail SSL (`smtp.gmail.com:465`)
- Supports Office 365 STARTTLS (`smtp.office365.com:587`)
- Includes short delay between messages to avoid bursts

## Security Guidance

- Do not hardcode passwords in source code
- Use app passwords when provider requires them (especially Gmail)
- Use a dedicated sender mailbox for auditability
- Limit repository sharing if templates/data include sensitive financial context

## Troubleshooting

### "Sheet 'Party List' not found"

- Ensure worksheet name is exactly `Party List`

### SMTP login failure

- Verify credentials
- Confirm app password is used when required
- Check org policies for SMTP AUTH and MFA constraints

### Outlook connection errors

- Ensure Outlook desktop app is installed
- Open Outlook once and verify profile is logged in
- Use Windows host for `win32com` mode

### Emails not generated

Check that:

- Required columns contain values (`Party Name`, `Email ID`)
- Email IDs are valid format
- AR/AP filters include the row type

## Limitations

- No attachment handling in current version
- No HTML email body rendering (plain text body only)
- No retry/backoff queue for failed SMTP sends
- No persistent database log (session-only dispatch log)

## Suggested Enhancements

- Add attachment support (statement PDFs, reconciliation docs)
- Add HTML template mode
- Export dispatch report to CSV/Excel
- Add role-based authentication before sending
- Add dry-run validation summary for skipped rows

## License

No license file is currently included in this repository. Add a `LICENSE` file if you intend to define usage terms.
