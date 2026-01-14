import streamlit as st
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
from imap_tools import MailBox, AND

st.set_page_config(page_title="Company Bulk Sender", layout="wide")

# --- 1. SIDEBAR: MANUAL EMAIL ENTRY ---
with st.sidebar:
    st.header("Login")
    from_email = st.text_input("Enter Company Email")
    batch_size = st.number_input("Batch Size", value=50)

# --- 2. LOGIC TO SELECT THE CORRECT FIXED PASSWORD ---
target_password = None
if from_email:
    # This looks into your Secrets table to find the matching password
    all_passwords = st.secrets["PASSWORDS"]
    if from_email in all_passwords:
        target_password = all_passwords[from_email]
    else:
        st.sidebar.error("This email is not registered in Secrets.")

# --- 3. THE SENDER LOGIC ---
st.title("📧 Bulk Emailer")
draft_subject = st.text_input("Draft Subject")
uploaded_file = st.file_uploader("Upload Recipients (Excel)", type=["xlsx"])

if st.button("🚀 Start Blasting"):
    if not target_password:
        st.error("Cannot proceed without a valid App Password.")
    elif not draft_subject or not uploaded_file:
        st.error("Please provide a subject and recipient list.")
    else:
        try:
            # STEP A: LOGIN & GET DRAFT
            with st.status("Logging in...") as status:
                with MailBox('outlook.office365.com').login(from_email, target_password, 'Drafts') as mb:
                    msgs = list(mb.fetch(AND(subject=draft_subject)))
                    if not msgs:
                        st.error("Draft not found!")
                        st.stop()
                    body = msgs[-1].html if msgs[-1].html else msgs[-1].text
                status.update(label="Login Success! Sending...", state="complete")

            # STEP B: LOAD EXCEL
            df = pd.read_excel(uploaded_file, header=None)
            emails = df.iloc[:, 0].dropna().astype(str).tolist()

            # STEP C: SEND IN BATCHES
            for i in range(0, len(emails), int(batch_size)):
                batch = emails[i : i + int(batch_size)]
                
                msg = EmailMessage()
                msg['Subject'] = draft_subject
                msg['From'] = from_email
                msg['To'] = from_email # Sends to yourself, BCCs the rest
                msg['Bcc'] = ", ".join(batch)
                msg.add_alternative(body, subtype='html')

                with smtplib.SMTP("smtp.office365.com", 587) as server:
                    server.starttls()
                    server.login(from_email, target_password)
                    server.send_message(msg)
                
                st.write(f"✅ Batch {i//batch_size + 1} sent.")
                time.sleep(5) # Pause to avoid company spam filters

            st.success("🎉 All emails sent successfully!")

        except Exception as e:
            st.error(f"Login Failed: {e}")
            st.info("Ensure IMAP is enabled in the Outlook Web settings for THIS specific email.")