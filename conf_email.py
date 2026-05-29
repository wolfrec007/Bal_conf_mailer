"""
Balance Confirmation – Bulk Email App (Streamlit)
==================================================
Run:
    pip install -r requirements.txt
    streamlit run conf_email.py
"""

import streamlit as st
import openpyxl
import smtplib
import ssl
import re
import io
import time
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Balance Confirmation Mailer",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS (Dark Mode Compatible) ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=DM+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid var(--primary-color); }
.stTabs [data-baseweb="tab"] { font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 1.5px; text-transform: uppercase; padding: 10px 20px; border-radius: 4px 4px 0 0; }
.sec-head { font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: 2px; text-transform: uppercase; color: var(--text-color); opacity: 0.7; border-bottom: 1px solid var(--border-color); padding-bottom: 6px; margin: 16px 0 12px; }
.metric-box { background: var(--secondary-background-color); border: 1px solid var(--border-color); border-radius: 6px; padding: 14px 18px; text-align: center; }
.metric-val { font-size: 24px; font-weight: bold; color: var(--primary-color); }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PLACEHOLDERS CONFIG
# Placeholder keys match Excel column names exactly: {{Column Name}}
# cfg_ keys come from the saved configuration panel.
# ══════════════════════════════════════════════════════════════════════════════

PLACEHOLDERS = {
    # ── Excel columns ──────────────────────────────────────────────────────
    "{{Party Name}}":             lambda r, cfg: str(r.get("Party Name", "")),
    "{{Email ID}}":               lambda r, cfg: str(r.get("Email ID", "")),
    "{{Contact Person}}":         lambda r, cfg: str(r.get("Contact Person") or "Sir/Madam"),
    "{{Type (AR/AP)}}":           lambda r, cfg: "AP" if "AP" in str(r.get("Type (AR/AP)", "")).upper() else "AR",
    "{{Currency}}":               lambda r, cfg: str(r.get("Currency") or "INR"),
    "{{Outstanding Balance}}":    lambda r, cfg: fmt_amount(r.get("Outstanding Balance", 0), r.get("Currency") or "INR"),
    "{{Due Date}}":               lambda r, cfg: str(r.get("Due Date", ""))[:10] if r.get("Due Date") else "as agreed",
    "{{Reference / Invoice No.}}": lambda r, cfg: str(r.get("Reference / Invoice No.") or "N/A"),
    "{{Remarks}}":                lambda r, cfg: str(r.get("Remarks") or ""),
    # ── Config panel values ────────────────────────────────────────────────
    "{{Confirmation Date}}":      lambda r, cfg: cfg.get("conf_date", ""),
    "{{Company Name}}":           lambda r, cfg: cfg.get("company", ""),
    "{{Signatory}}":              lambda r, cfg: cfg.get("signatory", ""),
    "{{Reply-To Email}}":         lambda r, cfg: cfg.get("reply_to", ""),
    # ── Derived ───────────────────────────────────────────────────────────
    "{{Flow}}":                   lambda r, cfg: "payable to your organisation" if "AP" in str(r.get("Type (AR/AP)", "")).upper() else "receivable from your organisation",
}

DEFAULT_SUBJECT = "Balance Confirmation as on {{Confirmation Date}} – {{Party Name}} [Ref: {{Reference / Invoice No.}}]"

DEFAULT_BODY = """Dear {{Contact Person}},

Greetings from {{Company Name}}!

As part of our periodic balance confirmation exercise, we kindly request you to confirm the outstanding balance as on {{Confirmation Date}}.

As per our records, the following amount is {{Flow}}:

  Party Name          : {{Party Name}}
  Outstanding Amount  : {{Outstanding Balance}}
  Reference           : {{Reference / Invoice No.}}
  Due Date            : {{Due Date}}
  Remarks             : {{Remarks}}

Kindly confirm the above balance by replying to this email. If there is any discrepancy, please share the relevant details and supporting documents so that we may reconcile at the earliest.

Your prompt response will be greatly appreciated.

For any queries, please write to us at: {{Reply-To Email}}

Warm regards,
{{Signatory}}
{{Company Name}}"""

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def fmt_amount(value, currency="INR"):
    try: return f"{currency} {float(value):,.2f}"
    except (TypeError, ValueError): return f"{currency} 0.00"

def extract_email(raw):
    """Extract plain address from 'Name <email@x.com>' or 'email@x.com'."""
    raw = str(raw or "").strip()
    match = re.search(r"<([^@]+@[^@]+\.[^@]+)>", raw)
    if match:
        return match.group(1).strip()
    return raw

def valid_email(e):
    return bool(e and re.match(r"[^@]+@[^@]+\.[^@]+", extract_email(str(e))))

def resolve(template, row, cfg):
    result = template
    for ph, fn in PLACEHOLDERS.items():
        result = result.replace(ph, fn(row, cfg))
    return result

def merge_cc(row_cc, global_cc):
    """Combine per-row CC and global CC, deduplicating.
    Accepts comma or semicolon separators and 'Name <email>' format."""
    parts = []
    for raw in (row_cc, global_cc):
        for addr in re.split(r"[,;]", str(raw or "")):
            addr = extract_email(addr)
            if addr and addr not in parts:
                parts.append(addr)
    return ", ".join(parts)

def find_header_row(ws):
    """Auto-detect whether headers are in row 1 or row 2."""
    for row_num in (1, 2):
        vals = [str(c.value).strip() if c.value else "" for c in ws[row_num]]
        if "Party Name" in vals and "Email ID" in vals:
            return row_num
    return 2

def read_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Party List" not in wb.sheetnames:
        return None, "Sheet 'Party List' not found. Please use the provided template."
    ws = wb["Party List"]
    header_row = find_header_row(ws)
    headers = [str(c.value).strip() if c.value else "" for c in ws[header_row]]
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(row): continue
        rec = dict(zip(headers, row))
        if rec.get("Party Name") and rec.get("Email ID"):
            rec["Email ID"] = extract_email(rec["Email ID"])
            rows.append(rec)
    return rows, None

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

defaults = {
    "rows": [], "drafts": [], "send_log": [], "file_id": None,
    "subject_tpl": DEFAULT_SUBJECT, "body_tpl": DEFAULT_BODY,
    "cfg": {},  # saved config
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
# HEADER & GLOBAL CONFIGURATION (MAIN AREA)
# ══════════════════════════════════════════════════════════════════════════════

hdr_col, help_col = st.columns([11, 1])
with hdr_col:
    st.markdown("""
<div style="background: var(--secondary-background-color); padding:24px 32px 20px;border-radius:8px;border-bottom:4px solid #c8873a;margin-bottom:24px;">
  <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:2px;opacity:.7;text-transform:uppercase;margin-bottom:4px;">Financial Operations</div>
  <div style="font-size:24px;font-weight:700;">Balance Confirmation Mailer</div>
</div>
""", unsafe_allow_html=True)
with help_col:
    with st.popover("❓", help="Placeholder reference"):
        st.markdown("### Placeholder Reference")
        st.markdown("Use these in your **Subject** and **Body** templates. They are replaced with each party's data when emails are generated.")
        st.markdown("#### From Excel columns")
        st.markdown("""
| Placeholder | Column / Source | Default if blank |
|---|---|---|
| `{{Party Name}}` | Party Name | — |
| `{{Email ID}}` | Email ID | — |
| `{{Contact Person}}` | Contact Person | Sir/Madam |
| `{{Type (AR/AP)}}` | Type (AR/AP) | AR |
| `{{Currency}}` | Currency | INR |
| `{{Outstanding Balance}}` | Outstanding Balance | 0.00 |
| `{{Due Date}}` | Due Date | as agreed |
| `{{Reference / Invoice No.}}` | Reference / Invoice No. | N/A |
| `{{Remarks}}` | Remarks | *(blank)* |
""")
        st.markdown("#### From Configuration panel")
        st.markdown("""
| Placeholder | Source |
|---|---|
| `{{Confirmation Date}}` | Confirmation Date field |
| `{{Company Name}}` | Company Name field |
| `{{Signatory}}` | Signatory & Designation field |
| `{{Reply-To Email}}` | Reply-To Email field |
""")
        st.markdown("#### Derived")
        st.markdown("""
| Placeholder | What it produces |
|---|---|
| `{{Flow}}` | *"receivable from your organisation"* (AR) or *"payable to your organisation"* (AP) |
""")

with st.expander("⚙️ Global Configuration", expanded=True):
    col_cfg1, col_cfg2, col_cfg3 = st.columns(3)

    with col_cfg1:
        st.markdown("**Company Details**")
        company_name   = st.text_input("Company Name",            placeholder="Horizon Industries",        key="w_company")
        signatory      = st.text_input("Signatory & Designation", placeholder="Ramesh Verma, CFO",         key="w_signatory")
        reply_to_email = st.text_input("Reply-To Email",          placeholder="accounts@yourcompany.com",  key="w_reply_to")
        conf_date      = st.date_input("Confirmation Date",       value=datetime.today(),                  key="w_conf_date")
        conf_date_str  = conf_date.strftime("%d %B %Y")

    with col_cfg2:
        st.markdown("**Email Engine & Settings**")
        send_provider = st.selectbox("Select Email Provider", [
            "Desktop Outlook App (win32com)",
            "Gmail (SMTP)",
            "Office 365 / Outlook (SMTP)"
        ], key="w_provider")

        email_action = "Send"
        smtp_user = smtp_pass = from_name = ""

        if "Desktop" in send_provider:
            email_action = st.radio("Action when processed:", ["Save as Drafts", "Send Immediately"], key="w_action")
        else:
            st.info("💡 SMTP Connections send emails immediately.")
            smtp_user = st.text_input("Sender Email Address", placeholder="you@domain.com", key="w_smtp_user")
            smtp_pass = st.text_input("App Password",  type="password",                     key="w_smtp_pass")
            from_name = st.text_input("Display Name",  placeholder="Accounts Team",          key="w_from_name")

    with col_cfg3:
        st.markdown("**Filters & CC**")
        global_cc = st.text_input(
            "Global CC (comma-separated)",
            placeholder="manager@co.com",
            key="w_global_cc",
            help="Added to every email. Per-party CC can also be set in the Excel 'CC' column."
        )
        type_filter = st.multiselect("Include types", ["AR", "AP"], default=["AR", "AP"], key="w_type_filter")

    st.divider()
    if st.button("💾 Save Config", type="primary"):
        st.session_state.cfg = {
            "company":       st.session_state.w_company,
            "signatory":     st.session_state.w_signatory,
            "reply_to":      st.session_state.w_reply_to,
            "conf_date":     conf_date_str,
            "global_cc":     st.session_state.w_global_cc,
            "type_filter":   st.session_state.w_type_filter,
            "send_provider": st.session_state.w_provider,
            "email_action":  email_action,
            "smtp_user":     smtp_user,
            "smtp_pass":     smtp_pass,
            "from_name":     from_name,
        }
        st.success("✅ Configuration saved.")

# Always use current widget values — saved config is only used when it exists and is non-empty.
# This prevents global_cc silently becoming "" when the user types but forgets to Save Config.
saved = st.session_state.cfg
if saved:
    cfg = {
        "company":   saved["company"],
        "signatory": saved["signatory"],
        "reply_to":  saved["reply_to"],
        "conf_date": saved["conf_date"],
    }
    effective_global_cc = saved["global_cc"]
    effective_filter    = saved["type_filter"]
    st.caption(f"🔒 Config saved — **{saved['company'] or '—'}** · {saved['conf_date']} · CC: {saved['global_cc'] or 'none'}")
else:
    cfg = {
        "company":   company_name,
        "signatory": signatory,
        "reply_to":  reply_to_email,
        "conf_date": conf_date_str,
    }
    effective_global_cc = global_cc
    effective_filter    = type_filter

# ══════════════════════════════════════════════════════════════════════════════
# MAIN TABS
# ══════════════════════════════════════════════════════════════════════════════

tab1, tab2, tab3, tab4 = st.tabs(["📂 Step 1 · Upload & Edit", "✍️ Step 2 · Template", "✏️ Step 3 · Review", "📤 Step 4 · Dispatch"])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — UPLOAD & EDIT
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    uploaded = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], help='Must contain sheet "Party List"')

    if uploaded and st.session_state.file_id != uploaded.file_id:
        rows, err = read_excel(uploaded.read())
        if err:
            st.error(f"❌ {err}")
        else:
            st.session_state.rows = rows
            st.session_state.drafts = []
            st.session_state.file_id = uploaded.file_id
            st.success(f"✅ Loaded **{len(rows)}** parties.")

    if st.session_state.rows:
        st.markdown('<div class="sec-head">Review & Edit Data</div>', unsafe_allow_html=True)
        st.caption("💡 Notice a typo? Double-click any cell below to edit the data before generating emails.")

        df = pd.DataFrame(st.session_state.rows)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", key="data_editor")
        st.session_state.rows = edited_df.to_dict("records")

        ar_rows = [r for r in st.session_state.rows if "AP" not in str(r.get("Type (AR/AP)", "")).upper()]
        ap_rows = [r for r in st.session_state.rows if "AP" in str(r.get("Type (AR/AP)", "")).upper()]
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Parties", len(st.session_state.rows))
        c2.metric("AR Entries", len(ar_rows))
        c3.metric("AP Entries", len(ap_rows))

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    with st.expander("📋 Available Placeholders — click to expand", expanded=False):
        st.markdown("Copy any placeholder into your Subject or Body template. It will be replaced with the matching value for each party when emails are generated.")
        st.markdown("#### From Excel columns")
        st.markdown("""
| Placeholder | Reads from column | Default if blank |
|---|---|---|
| `{{Party Name}}` | Party Name | — |
| `{{Email ID}}` | Email ID | — |
| `{{Contact Person}}` | Contact Person | Sir/Madam |
| `{{Type (AR/AP)}}` | Type (AR/AP) | AR |
| `{{Currency}}` | Currency | INR |
| `{{Outstanding Balance}}` | Outstanding Balance — formatted as *Currency X,XXX.XX* | 0.00 |
| `{{Due Date}}` | Due Date | as agreed |
| `{{Reference / Invoice No.}}` | Reference / Invoice No. | N/A |
| `{{Remarks}}` | Remarks | *(blank)* |
""")
        st.markdown("#### From Configuration panel")
        st.markdown("""
| Placeholder | What it inserts |
|---|---|
| `{{Confirmation Date}}` | The date set in the Confirmation Date field e.g. *31 March 2026* |
| `{{Company Name}}` | Your firm/company name |
| `{{Signatory}}` | Name and designation of the email signatory |
| `{{Reply-To Email}}` | The reply-to email address shown in the email footer |
""")
        st.markdown("#### Derived / Computed")
        st.markdown("""
| Placeholder | What it produces |
|---|---|
| `{{Flow}}` | *"receivable from your organisation"* for AR, *"payable to your organisation"* for AP |
""")

    subject_tpl = st.text_input("Subject Template", value=st.session_state.subject_tpl, key="subj_t")
    st.session_state.subject_tpl = subject_tpl
    body_tpl = st.text_area("Body Template", value=st.session_state.body_tpl, height=380, key="body_t")
    st.session_state.body_tpl = body_tpl

    if st.button("🔨 Generate Emails", type="primary", use_container_width=True):
        if not st.session_state.rows:
            st.warning("Upload data first.")
        elif not st.session_state.cfg:
            st.warning("⚠️ Please fill in the Configuration panel above and click **💾 Save Config** first.")
        else:
            drafts = []
            for row in st.session_state.rows:
                type_tag = "AP" if "AP" in str(row.get("Type (AR/AP)", "")).upper() else "AR"
                if type_tag in effective_filter and valid_email(row.get("Email ID", "")):
                    combined_cc = merge_cc(row.get("CC", ""), effective_global_cc)
                    drafts.append({
                        "party":   row.get("Party Name", ""),
                        "type":    type_tag,
                        "to":      str(row.get("Email ID", "")).strip(),
                        "cc":      combined_cc,
                        "subject": resolve(st.session_state.subject_tpl, row, cfg),
                        "body":    resolve(st.session_state.body_tpl, row, cfg),
                        "amount":  fmt_amount(row.get("Outstanding Balance", 0), row.get("Currency") or "INR"),
                        "include": True,
                    })
            # Pre-populate Review tab widget state so value= is respected on re-generate
            for i, d in enumerate(drafts):
                st.session_state[f"to_{i}"]  = d["to"]
                st.session_state[f"cc_{i}"]  = d["cc"]
                st.session_state[f"s_{i}"]   = d["subject"]
                st.session_state[f"b_{i}"]   = d["body"]
                st.session_state[f"inc_{i}"] = d["include"]
            st.session_state.drafts = drafts
            st.success(f"Generated {len(drafts)} emails. Go to Review Tab.")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — REVIEW
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    drafts = st.session_state.drafts
    if not drafts:
        st.info("Generate emails in Step 2.")
    for i, draft in enumerate(drafts):
        with st.expander(f"✉️ {draft['party']} | {draft['type']}", expanded=False):
            st.checkbox("Include this email in dispatch list", key=f"inc_{i}")
            draft["include"] = st.session_state[f"inc_{i}"]
            st.text_input("To",      key=f"to_{i}")
            draft["to"]      = st.session_state[f"to_{i}"]
            st.text_input("CC",      key=f"cc_{i}")
            draft["cc"]      = st.session_state[f"cc_{i}"]
            st.text_input("Subject", key=f"s_{i}")
            draft["subject"] = st.session_state[f"s_{i}"]
            st.text_area("Body",     key=f"b_{i}", height=500)
            draft["body"]    = st.session_state[f"b_{i}"]

    st.session_state.drafts = drafts

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — DISPATCH
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    # Resolve send settings from saved config if available
    eff_provider     = saved.get("send_provider", send_provider)
    eff_action       = saved.get("email_action",  email_action)
    eff_smtp_user    = saved.get("smtp_user",     smtp_user)
    eff_smtp_pass    = saved.get("smtp_pass",     smtp_pass)
    eff_from_name    = saved.get("from_name",     from_name)

    selected = [d for d in st.session_state.drafts if d.get("include")]
    if not selected:
        st.warning("No emails selected. Go to Step 3 and select emails to include.")
    else:
        st.write(f"**Ready to process:** {len(selected)} emails via `{eff_provider}`")

        if st.button(f"🚀 Execute: {eff_action} ({len(selected)} emails)", type="primary", use_container_width=True):
            if "SMTP" in eff_provider and (not eff_smtp_user or not eff_smtp_pass):
                st.error("❌ Please enter your SMTP Email and App Password in the Configuration panel and Save Config.")
            else:
                progress_bar = st.progress(0)
                status_text  = st.empty()
                results      = []
                start_time   = time.time()
                server = outlook = None

                try:
                    if "Desktop" in eff_provider:
                        import win32com.client, pythoncom
                        pythoncom.CoInitialize()
                        outlook = win32com.client.Dispatch("Outlook.Application")
                    elif "Gmail" in eff_provider:
                        context = ssl.create_default_context()
                        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
                        server.login(eff_smtp_user, eff_smtp_pass)
                    else:
                        server = smtplib.SMTP("smtp.office365.com", 587)
                        server.starttls()
                        server.login(eff_smtp_user, eff_smtp_pass)
                except Exception as e:
                    st.error(f"Connection Error: {e}")
                    st.stop()

                for idx, d in enumerate(selected):
                    status_text.markdown(f"**Processing:** {d['party']} ({idx+1}/{len(selected)})")
                    try:
                        if "Desktop" in eff_provider:
                            mail         = outlook.CreateItem(0)
                            mail.To      = d["to"]
                            mail.CC      = d.get("cc", "")
                            mail.Subject = d["subject"]
                            mail.Body    = d["body"]
                            if "Send" in eff_action:
                                mail.Send()
                                results.append({"Party": d["party"], "Email": d["to"], "CC": d.get("cc", ""), "Status": "✅ Sent via App"})
                            else:
                                mail.Save()
                                results.append({"Party": d["party"], "Email": d["to"], "CC": d.get("cc", ""), "Status": "✅ Draft Saved"})
                        else:
                            msg            = MIMEMultipart("alternative")
                            msg["Subject"] = d["subject"]
                            msg["From"]    = f"{eff_from_name or cfg.get('company','')} <{eff_smtp_user}>"
                            msg["To"]      = d["to"]
                            cc_list        = [c.strip() for c in d.get("cc", "").split(",") if c.strip()]
                            if cc_list: msg["Cc"] = ", ".join(cc_list)
                            msg.attach(MIMEText(d["body"], "plain"))
                            server.sendmail(eff_smtp_user, [d["to"]] + cc_list, msg.as_string())
                            results.append({"Party": d["party"], "Email": d["to"], "CC": d.get("cc", ""), "Status": "✅ Sent via SMTP"})
                            if idx < len(selected) - 1:
                                time.sleep(0.5)
                    except Exception as e:
                        results.append({"Party": d["party"], "Email": d["to"], "CC": d.get("cc", ""), "Status": f"❌ Error: {e}"})

                    progress_bar.progress((idx + 1) / len(selected))

                if server:
                    try: server.quit()
                    except: pass

                st.session_state.send_log = results
                status_text.empty()
                progress_bar.empty()

                success_count = sum(1 for r in results if "✅" in r["Status"])
                fail_count    = sum(1 for r in results if "❌" in r["Status"])
                time_taken    = round(time.time() - start_time, 1)

                st.markdown(f"""
                <div style="display:flex;gap:15px;margin-bottom:20px;">
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">✅ Success</div><div class="metric-val" style="color:#2e7d32;">{success_count}</div></div>
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">❌ Failed</div><div class="metric-val" style="color:#d32f2f;">{fail_count}</div></div>
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">⏱️ Time Taken</div><div class="metric-val">{time_taken}s</div></div>
                </div>
                """, unsafe_allow_html=True)

    if st.session_state.send_log:
        st.markdown('<div class="sec-head">Detailed Dispatch Log</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(st.session_state.send_log), use_container_width=True)
