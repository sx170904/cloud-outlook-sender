import streamlit as st
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
from imap_tools import MailBox, AND

st.set_page_config(page_title="Outlook Cloud Sender", layout="wide")
st.title("📧 Outlook Universal Sender (Cloud)")

# --- 1. SIDEBAR: MANUAL LOGIN ---
with st.sidebar:
    st.header("1. Account Settings")
    # You can now type any email and its specific App Password here
    from_email = st.text_input("Send From (Outlook Email)")
    app_password = st.text_input("App Password (16 Characters)", type="password") 
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    st.markdown("---")
    st.warning("⚠️ **Reminder:** You must use an 'App Password' generated from your Microsoft Security settings, not your normal login password.")

# --- 2. UI: DRAFT & RECIPIENTS ---
st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")
with col2:
    uploaded_file = st.file_uploader("Upload Excel (Emails in 1st Column)", type=["xlsx"])

# --- 3. THE LOGIC ---
if st.button("🚀 Send Email(s)"):
    if not from_email or not app_password or not draft_subject:
        st.error("Please fill in Email, App Password, and Draft Subject.")
    else:
        try:
            # STEP A: FETCH THE DRAFT CONTENT
            with st.status("Logging into Outlook and searching drafts...") as status:
                # We connect to Outlook's cloud server directly
                with MailBox('outlook.office365.com').login(from_email, app_password, 'Drafts') as mb:
                    messages = list(mb.fetch(AND(subject=draft_subject)))
                    
                    if not messages:
                        st.error(f"Could not find a draft with subject: '{draft_subject}'")
                        st.stop()
                    
                    target_msg = messages[-1] 
                    body_content = target_msg.html if target_msg.html else target_msg.text
                status.update(label="Draft content retrieved!", state="complete")

            # STEP B: PREPARE RECIPIENTS
            bcc_list = []
            if uploaded_file:
                df = pd.read_excel(uploaded_file, header=None)
                all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

            # STEP C: SENDING FUNCTION
            def send_mail_batch(recipients_batch):
                msg = EmailMessage()
                msg['Subject'] = draft_subject
                msg['From'] = from_email
                msg['To'] = to_email if to_email else from_email
                if cc_email: msg['Cc'] = cc_email
                if recipients_batch: msg['Bcc'] = ", ".join(recipients_batch)
                
                msg.add_alternative(body_content, subtype='html')
                
                with smtplib.SMTP("smtp.office365.com", 587) as server:
                    server.starttls()
                    server.login(from_email, app_password)
                    server.send_message(msg)

            # STEP D: BATCH EXECUTION
            if bcc_list:
                total_batches = (len(bcc_list) + batch_size - 1) // batch_size
                for i in range(0, len(bcc_list), int(batch_size)):
                    batch_num = (i // int(batch_size)) + 1
                    current_batch = bcc_list[i : i + int(batch_size)]
                    
                    send_mail_batch(current_batch)
                    st.write(f"✅ Sent Batch {batch_num} of {total_batches}")
                    
                    if i + int(batch_size) < len(bcc_list):
                        time.sleep(5)
                st.success(f"🎉 All {len(bcc_list)} emails sent!")
            else:
                send_mail_batch([])
                st.success("✅ Single email sent successfully!")

        except Exception as e:
            st.error(f"Connection Error: {e}")
            st.info("Check if your App Password is correct and IMAP is enabled in Outlook settings.")