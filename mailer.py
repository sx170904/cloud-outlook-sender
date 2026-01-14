import streamlit as st
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
from imap_tools import MailBox, AND

st.set_page_config(page_title="Outlook Cloud Sender", layout="wide")
st.title("📧 Outlook Universal Sender (Cloud)")

# --- 1. GET THE FIXED PASSWORD FROM SECRETS ---
try:
    # This stays hidden and fixed
    FIXED_APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception:
    st.error("Missing App Password in Streamlit Secrets!")
    st.stop()

# --- 2. SIDEBAR: MANUAL EMAIL ENTRY ---
with st.sidebar:
    st.header("1. Account Settings")
    # You still enter your email manually here
    from_email = st.text_input("Your Outlook Email")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    st.markdown("---")
    st.info("🔐 App Password is fixed in the system backend.")

# --- 3. UI: DRAFT & RECIPIENTS ---
st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")
with col2:
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

# --- 4. THE LOGIC ---
if st.button("🚀 Send Email(s)"):
    if not from_email or not draft_subject:
        st.error("Please enter your Email and Draft Subject.")
    else:
        try:
            # STEP A: FETCH DRAFT (Using fixed password)
            with st.status("Accessing Outlook...") as status:
                with MailBox('outlook.office365.com').login(from_email, FIXED_APP_PASSWORD, 'Drafts') as mb:
                    messages = list(mb.fetch(AND(subject=draft_subject)))
                    if not messages:
                        st.error(f"Draft not found: '{draft_subject}'")
                        st.stop()
                    target_msg = messages[-1]
                    body_content = target_msg.html if target_msg.html else target_msg.text
                status.update(label="Draft retrieved!", state="complete")

            # STEP B: PREPARE RECIPIENTS
            bcc_list = []
            if uploaded_file:
                df = pd.read_excel(uploaded_file, header=None)
                all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

            # STEP C: SENDING (Using fixed password)
            def send_mail(batch):
                msg = EmailMessage()
                msg['Subject'] = draft_subject
                msg['From'] = from_email
                msg['To'] = to_email if to_email else from_email
                if cc_email: msg['Cc'] = cc_email
                if batch: msg['Bcc'] = ", ".join(batch)
                msg.add_alternative(body_content, subtype='html')
                
                with smtplib.SMTP("smtp.office365.com", 587) as server:
                    server.starttls()
                    server.login(from_email, FIXED_APP_PASSWORD)
                    server.send_message(msg)

            # STEP D: EXECUTION
            if bcc_list:
                for i in range(0, len(bcc_list), int(batch_size)):
                    send_mail(bcc_list[i : i + int(batch_size)])
                    st.write(f"✅ Sent Batch {i//batch_size + 1}")
                    if i + int(batch_size) < len(bcc_list): time.sleep(5)
                st.success("🎉 Done!")
            else:
                send_mail([])
                st.success("✅ Email sent!")

        except Exception as e:
            st.error(f"Error: {e}")