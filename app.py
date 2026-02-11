import streamlit as st
import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURATION ---
SMTP_SERVER = "mail.drools.com" 
SMTP_PORT = 587
SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]

REMOVE_RSM_POD = ["ECOM & MT", "ESPT", "VET PHARMA", "EXPORT", "AQUA"]

DISPATCH_ORDER = [
    "Plant", "Location", "Customer_No", "Customer_Name",
    "Billing_Date", "Billing_Doc", "Bill_Amount",
    "Pending Days", "Disptch_Remark", "Yesterday Remarks",
    "Yesterday Standard Remarks", "RSM_Name", "ASM_Name", "RM", "Gross_Weight_Tons"
]

POD_ORDER = [
    "Plant", "Location", "Billing_Date", "Month", "Days", "Year",
    "Billing_Doc", "Customer_No", "Customer_Name",
    "Customer_City", "Local/Upcountry", "ASM_Name", "RSM_Name", 
    "Bill_Amount", "Dispatch_Date", "Disptch_Remark", "RM"
]

st.set_page_config(page_title="Drools Automation Hub", layout="wide")

def apply_excel_format(writer, sheet_name, df_obj):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#B8CCE4', 'border': 1, 'align': 'center'})
    cell_fmt = workbook.add_format({'border': 1})
    worksheet.conditional_format(0, 0, len(df_obj), len(df_obj.columns) - 1, {'type': 'no_errors', 'format': cell_fmt})
    for i, col in enumerate(df_obj.columns):
        worksheet.write(0, i, col, header_fmt)
        if not df_obj.empty:
            max_len = max(df_obj[col].astype(str).map(len).max(), len(str(col))) + 2
            worksheet.set_column(i, i, min(max_len, 50)) 

def send_email_smtp(to_email, subject, body_html, attachment_path, cc_emails=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        recipients = [to_email]
        if cc_emails and str(cc_emails) != 'nan':
            msg['Cc'] = cc_emails
            recipients.extend([email.strip() for email in cc_emails.split(';')])
        msg.attach(MIMEText(body_html, 'html'))
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
            msg.attach(part)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Mail failed: {e}")
        return False

st.title("📊 Drools Automation Hub")

with st.sidebar:
    raw_file = st.file_uploader("Upload ZSD_DISPATCH", type="xlsx")
    mapping_file = st.file_uploader("Upload Mapping.xlsx", type="xlsx")
    yesterday_file = st.file_uploader("Upload Yesterday Report", type="xlsx")

if raw_file and mapping_file:
    df = pd.read_excel(raw_file)
    mapping = pd.read_excel(mapping_file, sheet_name="Depot_Zone")
    email_df = pd.read_excel(mapping_file, sheet_name="Email_IDs")
    
    mapping.columns = mapping.columns.str.strip().str.upper()
    df["Plant"] = df["Plant"].astype(str).str.strip()
    mapping["PLANT"] = mapping["PLANT"].astype(str).str.strip()
    df = df.merge(mapping, left_on="Plant", right_on="PLANT", how="left")
    df = df.rename(columns={"LOCATION": "Location", "RM": "RM"})
    df = df[df["Plant_Name"].astype(str).str.upper().str.contains("DROOLS PET FOOD", na=False)]
    
    df["Dispatch_Date"] = pd.to_datetime(df["Dispatch_Date"], errors="coerce")
    df["Billing_Date"] = pd.to_datetime(df["Billing_Date"], errors="coerce")
    today = pd.Timestamp.now().normalize()

    # Yesterday Remarks logic
    df["Yesterday Remarks"] = ""
    df["Yesterday Standard Remarks"] = ""
    if yesterday_file:
        y_df = pd.read_excel(yesterday_file, sheet_name="Dispatch")
        y_df["Billing_Doc"] = y_df["Billing_Doc"].astype(str).str.strip()
        df["Billing_Doc_Match"] = df["Billing_Doc"].astype(str).str.strip()
        rem_map = y_df.set_index("Billing_Doc")["Disptch_Remark"].to_dict()
        df["Yesterday Remarks"] = df["Billing_Doc_Match"].map(rem_map).fillna("")

    df = df.sort_values(by=["Location", "Billing_Date"], ascending=[True, True])

    dispatch_df = df[df["Dispatch_Date"].isna()].copy()
    dispatch_df["Pending Days"] = (today - dispatch_df["Billing_Date"]).dt.days
    dispatch_df["Gross_Weight_Tons"] = (pd.to_numeric(dispatch_df["Gross_Weight"], errors="coerce") / 1000).round(3)
    dispatch_df = dispatch_df[[c for c in DISPATCH_ORDER if c in dispatch_df.columns]]

    pod_df = df[df["Dispatch_Date"].notna()].copy()
    pattern_p = "|".join(REMOVE_RSM_POD)
    pod_df = pod_df[~pod_df["RSM_Name"].str.upper().str.contains(pattern_p, na=False)]
    pod_df["Month"] = pod_df["Billing_Date"].dt.strftime("%b")
    pod_df = pod_df[[c for c in POD_ORDER if c in pod_df.columns]]

    # --- DISPATCH PIVOT WITH GRAND TOTAL ---
    dispatch_pivot = dispatch_df.groupby("Location").agg({"Billing_Doc":"count", "Bill_Amount":"sum", "Gross_Weight_Tons":"sum"}).rename(columns={"Billing_Doc":"Invoice Count"}).reset_index()
    d_total = pd.DataFrame({"Location":["Grand Total"], "Invoice Count":[dispatch_pivot["Invoice Count"].sum()], "Bill_Amount":[dispatch_pivot["Bill_Amount"].sum()], "Gross_Weight_Tons":[dispatch_pivot["Gross_Weight_Tons"].sum()]})
    dispatch_pivot = pd.concat([dispatch_pivot, d_total], ignore_index=True)

    # --- POD PIVOT ---
    pod_pivot = pd.pivot_table(pod_df, index=["RM", "Location"], columns="Month", values="Billing_Doc", aggfunc="count", fill_value=0, margins=True, margins_name="Grand Total").reset_index()
    pod_pivot.columns.name = None

    # --- DASHBOARD ---
    st.markdown("---")
    st.subheader("📌 Live Summary Dashboard")
    met1, met2, met3 = st.columns(3)
    met1.metric("Pending Invoices", f"{len(dispatch_df)}")
    met2.metric("Pending Amount", f"₹{dispatch_df['Bill_Amount'].sum():,.0f}")
    met3.metric("Processed POD Count", f"{len(pod_df)}")
    st.markdown("---")

    if st.button("🚀 Run Automation & Send Mails"):
        d_excel = dispatch_df.copy()
        p_excel = pod_df.copy()
        d_excel["Billing_Date"] = d_excel["Billing_Date"].dt.strftime('%d-%m-%Y')
        p_excel["Billing_Date"] = p_excel["Billing_Date"].dt.strftime('%d-%m-%Y')
        p_excel["Dispatch_Date"] = p_excel["Dispatch_Date"].dt.strftime('%d-%m-%Y')

        master_fname = f"Pending dispatches and pod {datetime.now().strftime('%d-%b-%Y')}.xlsx"
        
        with pd.ExcelWriter(master_fname, engine='xlsxwriter') as writer:
            d_excel.to_excel(writer, sheet_name="Dispatch", index=False)
            p_excel.to_excel(writer, sheet_name="Pod", index=False)
            dispatch_pivot.to_excel(writer, sheet_name="Dispatch pivot", index=False)
            pod_pivot.to_excel(writer, sheet_name="Pod pivot", index=False)
            apply_excel_format(writer, "Dispatch", d_excel)
            apply_excel_format(writer, "Pod", p_excel)
            apply_excel_format(writer, "Dispatch pivot", dispatch_pivot)
            apply_excel_format(writer, "Pod pivot", pod_pivot)

        email_map_to = dict(zip(email_df['Target'].astype(str).str.strip(), email_df['Email'].astype(str).str.strip()))
        email_map_cc = dict(zip(email_df['Target'].astype(str).str.strip(), email_df['CC'].astype(str).str.strip()))
        all_targets = set(list(dispatch_df['Location'].unique()) + list(pod_df['RM'].unique()))
        
        for target in all_targets:
            if str(target) in email_map_to:
                fname = f"Report_{target}.xlsx"
                is_loc = target in dispatch_df['Location'].values
                
                sub_disp_xl = d_excel[dispatch_df['Location' if is_loc else 'RM'] == target].copy()
                sub_pod_xl = p_excel[pod_df['Location' if is_loc else 'RM'] == target].copy()

                # --- DISPATCH TABLE IN MAIL (LOCATION ADDED) ---
                crit_data = dispatch_df[(dispatch_df['Location' if is_loc else 'RM'] == target) & (dispatch_df["Pending Days"] > 7)].copy()
                crit_data["Billing_Date"] = crit_data["Billing_Date"].dt.strftime('%d-%m-%Y')
                
                # Selecting columns for mail body table
                mail_cols = ["Location", "Billing_Doc", "Customer_Name", "Billing_Date", "Pending Days", "Bill_Amount"]
                t_disp = crit_data[mail_cols].to_html(index=False, border=1) if not crit_data.empty else "<p>No critical pending (>7 days).</p>"
                
                loc_sum = pod_pivot[pod_pivot['Location' if is_loc else 'RM'] == target]
                t_pod = loc_sum.to_html(index=False, border=1) if not loc_sum.empty else ""

                body = f"<html><body style='font-family: Calibri;'>Dear <b>{target}</b>,<br>Attached is your daily report.<br><h3 style='color:red;'>⚠️ Critical Pending (>7 Days)</h3>{t_disp}<h3>📦 POD Summary</h3>{t_pod}<br>Regards, <b>Aditya</b></body></html>"

                with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
                    sub_disp_xl.to_excel(writer, sheet_name="Dispatch", index=False)
                    sub_pod_xl.to_excel(writer, sheet_name="Pod", index=False)
                    apply_excel_format(writer, "Dispatch", sub_disp_xl)
                    apply_excel_format(writer, "Pod", sub_pod_xl)

                send_email_smtp(email_map_to[str(target)], f"Daily Report - {target}", body, fname, email_map_cc.get(str(target)))
                st.write(f"✅ Sent: {target}")

        send_email_smtp(SENDER_EMAIL, "Master Report Consolidated", "Attached is the main sheet.", master_fname)
        st.success("Main Sheet Sent to Admin & Target Mails Dispatched!")
        st.balloons()
