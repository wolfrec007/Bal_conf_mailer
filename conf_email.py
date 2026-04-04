"""
Balance Confirmation – Bulk Email App (Streamlit)
==================================================
Run:
    pip install -r requirements.txt
    streamlit run balance_email_app.py
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
    initial_sidebar_state="collapsed", # Changed to collapsed since we don't use it anymore
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
# ══════════════════════════════════════════════════════════════════════════════

PLACEHOLDERS = {
    "{{party_name}}":    ("Party Name",              lambda r, cfg: str(r.get("Party Name", ""))),
    "{{contact}}":       ("Contact Person",           lambda r, cfg: str(r.get("Contact Person") or "Sir/Madam")),
    "{{email}}":         ("Email ID",                 lambda r, cfg: str(r.get("Email ID", ""))),
    "{{type}}":          ("Type (AR/AP)",             lambda r, cfg: "AP" if "AP" in str(r.get("Type (AR/AP)", "")).upper() else "AR"),
    "{{currency}}":      ("Currency",                 lambda r, cfg: str(r.get("Currency") or "INR")),
    "{{amount}}":        ("Outstanding Balance",      lambda r, cfg: fmt_amount(r.get("Outstanding Balance", 0), r.get("Currency") or "INR")),
    "{{due_date}}":      ("Due Date",                 lambda r, cfg: str(r.get("Due Date", ""))[:10] if r.get("Due Date") else "as agreed"),
    "{{reference}}":     ("Reference / Invoice No.",  lambda r, cfg: str(r.get("Reference / Invoice No.") or "N/A")),
    "{{remarks}}":       ("Remarks",                  lambda r, cfg: str(r.get("Remarks") or "")),
    "{{conf_date}}":     ("Confirmation Date",        lambda r, cfg: cfg.get("conf_date", "")),
    "{{company}}":       ("Your Company Name",        lambda r, cfg: cfg.get("company", "")),
    "{{signatory}}":     ("Signatory",                lambda r, cfg: cfg.get("signatory", "")),
    "{{reply_to}}":      ("Reply-To Email",           lambda r, cfg: cfg.get("reply_to", "")),
    "{{flow}}":          ("AR/AP Direction Text",     lambda r, cfg: "payable to your organisation" if "AP" in str(r.get("Type (AR/AP)", "")).upper() else "receivable from your organisation"),
}

DEFAULT_SUBJECT = "Balance Confirmation as on {{conf_date}} – {{party_name}} [Ref: {{reference}}]"

DEFAULT_BODY = """Dear {{contact}},

Greetings from {{company}}!

As part of our periodic balance confirmation exercise, we kindly request you to confirm the outstanding balance as on {{conf_date}}.

As per our records, the following amount is {{flow}}:

  Party Name          : {{party_name}}
  Outstanding Amount  : {{amount}}
  Reference           : {{reference}}
  Due Date            : {{due_date}}
  Remarks             : {{remarks}}

Kindly confirm the above balance by replying to this email. If there is any discrepancy, please share the relevant details and supporting documents so that we may reconcile at the earliest.

Your prompt response will be greatly appreciated.

For any queries, please write to us at: {{reply_to}}

Warm regards,
{{signatory}}
{{company}}"""

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def fmt_amount(value, currency="INR"):
    try: return f"{currency} {float(value):,.2f}"
    except (TypeError, ValueError): return f"{currency} 0.00"

def valid_email(e):
    return bool(e and re.match(r"[^@]+@[^@]+\.[^@]+", str(e).strip()))

def resolve(template, row, cfg):
    result = template
    for ph, (_, fn) in PLACEHOLDERS.items():
        result = result.replace(ph, fn(row, cfg))
    return result

def read_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Party List" not in wb.sheetnames:
        return None, "Sheet 'Party List' not found. Please use the provided template."
    ws = wb["Party List"]
    headers = [str(c.value).strip() if c.value else "" for c in ws[2]]
    rows = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(row): continue
        rec = dict(zip(headers, row))
        if rec.get("Party Name") and rec.get("Email ID"):
            rows.append(rec)
    return rows, None

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

for k, v in {"rows": [], "drafts": [], "send_log": [], "file_id": None, "subject_tpl": DEFAULT_SUBJECT, "body_tpl": DEFAULT_BODY}.items():
    if k not in st.session_state: st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
# HEADER & GLOBAL CONFIGURATION (MAIN AREA)
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div style="background: var(--secondary-background-color); padding:24px 32px 20px;border-radius:8px;border-bottom:4px solid #c8873a;margin-bottom:24px;">
  <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:2px;opacity:.7;text-transform:uppercase;margin-bottom:4px;">Financial Operations</div>
  <div style="font-size:24px;font-weight:700;">Balance Confirmation Mailer</div>
</div>
""", unsafe_allow_html=True)

with st.expander("⚙️ Global Configuration", expanded=True):
    col_cfg1, col_cfg2, col_cfg3 = st.columns(3)
    
    with col_cfg1:
        st.markdown("**Company Details**")
        company_name   = st.text_input("Company Name",            placeholder="Horizon Industries")
        signatory      = st.text_input("Signatory & Designation", placeholder="Ramesh Verma, CFO")
        reply_to_email = st.text_input("Reply-To Email",          placeholder="accounts@yourcompany.com")
        conf_date      = st.date_input("Confirmation Date",       value=datetime.today())
        conf_date_str  = conf_date.strftime("%d %B %Y")
        
    with col_cfg2:
        st.markdown("**Email Engine & Settings**")
        send_provider = st.selectbox("Select Email Provider", [
            "Desktop Outlook App (win32com)", 
            "Gmail (SMTP)", 
            "Office 365 / Outlook (SMTP)"
        ])

        email_action = "Send"
        smtp_user = smtp_pass = from_name = ""

        if "Desktop" in send_provider:
            email_action = st.radio("Action when processed:", ["Save as Drafts", "Send Immediately"])
        else:
            st.info("💡 SMTP Connections send emails immediately.")
            smtp_user = st.text_input("Sender Email Address", placeholder="you@domain.com")
            smtp_pass = st.text_input("App Password",  type="password")
            from_name = st.text_input("Display Name",  placeholder="Accounts Team")

    with col_cfg3:
        st.markdown("**Filters & CC**")
        global_cc = st.text_input("Global CC (comma-separated)",  placeholder="manager@co.com")
        type_filter = st.multiselect("Include types", ["AR", "AP"], default=["AR", "AP"])

cfg = {
    "company": company_name, "signatory": signatory, 
    "reply_to": reply_to_email, "conf_date": conf_date_str,
}

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
    subject_tpl = st.text_input("Subject Template", value=st.session_state.subject_tpl, key="subj_t")
    st.session_state.subject_tpl = subject_tpl
    body_tpl = st.text_area("Body Template", value=st.session_state.body_tpl, height=350, key="body_t")
    st.session_state.body_tpl = body_tpl

    if st.button("🔨 Generate Emails", type="primary", use_container_width=True):
        if not st.session_state.rows:
            st.warning("Upload data first.")
        else:
            drafts = []
            for row in st.session_state.rows:
                type_tag = "AP" if "AP" in str(row.get("Type (AR/AP)", "")).upper() else "AR"
                if type_tag in type_filter and valid_email(row.get("Email ID", "")):
                    drafts.append({
                        "party": row.get("Party Name", ""), "type": type_tag,
                        "to": str(row.get("Email ID", "")).strip(), "cc": global_cc,
                        "subject": resolve(st.session_state.subject_tpl, row, cfg),
                        "body": resolve(st.session_state.body_tpl, row, cfg),
                        "amount": fmt_amount(row.get("Outstanding Balance", 0), row.get("Currency") or "INR"),
                        "include": True,
                    })
            st.session_state.drafts = drafts
            st.success(f"Generated {len(drafts)} emails. Go to Review Tab.")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — REVIEW
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    drafts = st.session_state.drafts
    if not drafts: st.info("Generate emails in Step 2.")
    for i, draft in enumerate(drafts):
        static_title = f"✉️ {draft['party']} | {draft['type']}"
        with st.expander(static_title, expanded=False):
            draft["include"] = st.checkbox("Include this email in dispatch list", value=draft["include"], key=f"inc_{i}")
            draft["to"] = st.text_input("To", value=draft["to"], key=f"to_{i}")
            draft["cc"] = st.text_input("CC", value=draft["cc"], key=f"cc_{i}")
            draft["subject"] = st.text_input("Subject", value=draft["subject"], key=f"s_{i}")
            draft["body"] = st.text_area("Body", value=draft["body"], key=f"b_{i}", height=200)
    
    st.session_state.drafts = drafts

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — DISPATCH 
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    selected = [d for d in st.session_state.drafts if d.get("include")]
    if not selected:
        st.warning("No emails selected. Go to Step 3 and select emails to include.")
    else:
        st.write(f"**Ready to process:** {len(selected)} emails via `{send_provider}`")
        
        btn_label = f"🚀 Execute: {email_action} ({len(selected)} emails)"
        
        if st.button(btn_label, type="primary", use_container_width=True):
            if "SMTP" in send_provider and (not smtp_user or not smtp_pass):
                st.error("❌ Please enter your SMTP Email and App Password in the Configuration panel.")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                results = []
                start_time = time.time()
                
                # Setup Connection based on Provider
                server = None
                outlook = None
                
                try:
                    if "Desktop" in send_provider:
                        import win32com.client
                        import pythoncom
                        
                        pythoncom.CoInitialize() 
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        
                    elif "Gmail" in send_provider:
                        context = ssl.create_default_context()
                        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
                        server.login(smtp_user, smtp_pass)
                        
                    else: # Office 365
                        server = smtplib.SMTP("smtp.office365.com", 587)
                        server.starttls()
                        server.login(smtp_user, smtp_pass)
                        
                except Exception as e:
                    st.error(f"Connection Error: {e}")
                    st.stop()

                # Dispatch Loop
                for idx, d in enumerate(selected):
                    status_text.markdown(f"**Processing:** {d['party']} ({idx+1}/{len(selected)})")
                    
                    try:
                        if "Desktop" in send_provider:
                            mail = outlook.CreateItem(0)
                            mail.To = d["to"]
                            mail.CC = d.get("cc", "")
                            mail.Subject = d["subject"]
                            mail.Body = d["body"]
                            
                            if "Send" in email_action:
                                mail.Send()
                                results.append({"Party": d["party"], "Email": d["to"], "Status": "✅ Sent via App"})
                            else:
                                mail.Save()
                                results.append({"Party": d["party"], "Email": d["to"], "Status": "✅ Draft Saved"})
                                
                        else: # SMTP Sending
                            msg = MIMEMultipart("alternative")
                            msg["Subject"] = d["subject"]
                            msg["From"] = f"{from_name or company_name} <{smtp_user}>"
                            msg["To"] = d["to"]
                            cc_list = [c.strip() for c in d.get("cc", "").split(",") if c.strip()]
                            if cc_list: msg["Cc"] = ", ".join(cc_list)
                            msg.attach(MIMEText(d["body"], "plain"))
                            
                            server.sendmail(smtp_user, [d["to"]] + cc_list, msg.as_string())
                            results.append({"Party": d["party"], "Email": d["to"], "Status": "✅ Sent via SMTP"})
                            
                            if idx < len(selected) - 1:
                                time.sleep(0.5) 
                                
                    except Exception as e:
                        results.append({"Party": d["party"], "Email": d["to"], "Status": f"❌ Error: {e}"})
                    
                    progress_bar.progress((idx + 1) / len(selected))
                
                # Cleanup SMTP
                if server:
                    try: server.quit()
                    except: pass
                
                end_time = time.time()
                st.session_state.send_log = results
                
                status_text.empty()
                progress_bar.empty()
                
                # Display Summary Metrics
                success_count = sum(1 for r in results if "✅" in r["Status"])
                fail_count = sum(1 for r in results if "❌" in r["Status"])
                time_taken = round(end_time - start_time, 1)

                st.markdown(f"""
                <div style="display:flex;gap:15px;margin-bottom:20px;">
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">✅ Success</div><div class="metric-val" style="color:#2e7d32;">{success_count}</div></div>
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">❌ Failed</div><div class="metric-val" style="color:#d32f2f;">{fail_count}</div></div>
                    <div class="metric-box" style="flex:1;"><div style="font-size:12px;opacity:0.7;">⏱️ Time Taken</div><div class="metric-val">{time_taken}s</div></div>
                </div>
                """, unsafe_allow_html=True)

    if st.session_state.send_log:
        st.markdown('<div class="sec-head">Detailed Dispatch Log</div>', unsafe_allow_html=True)
        log_df = pd.DataFrame(st.session_state.send_log)
        st.dataframe(log_df, use_container_width=True)