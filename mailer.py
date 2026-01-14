import streamlit as st
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
from imap_tools import MailBox, AND

st.set_page_config(page_title="Company Bulk Sender", layout="wide")
st.title("📧 Outlook Universal Sender (Cloud)")

# --- 1. LOGIN LOGIC (Fixed Passwords from Secrets) ---
with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Enter Company Email")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    # Lookup the password based on the email entered
    target_password = None
    if from_email:
        try:
            # Looks in Secrets for the [PASSWORDS] table
            pass_table = st.secrets["PASSWORDS"]
            if from_email in pass_table:
                target_password = pass_table[from_email]
                st.success("✅ Password Found in System")
            else:
                st.error("❌ This email is not in the secret list.")
        except Exception:
            st.error("❌ Secrets not configured correctly.")

# --- 2. THE UI: TO, CC, AND DRAFT ---
st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject (Must match Outlook Draft exactly)")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Main Recipient)")
    cc_email = st.text_input("CC (Optional)")
with col2:
    st.info("The Excel file should have emails in the FIRST column.")
    uploaded_file = st.file_uploader("Upload Excel for BCC", type=["xlsx"])

# --- 3. THE SENDING LOGIC ---
if st.button("🚀 Send Email(s)"):
    if not from_email or not target_password:
        st.error("Please enter a valid company email.")
    elif not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        try:
            # STEP A: GET DRAFT CONTENT
            with st.status("Fetching your draft from Outlook...") as status:
                with MailBox('outlook.office365.com').login(from_email, target_password, 'Drafts') as mb:
                    messages = list(mb.fetch(AND(subject=draft_subject)))
                    if not messages:
                        st.error(f"Could not find draft with subject: '{draft_subject}'")
                        st.stop()
                    
                    target_msg = messages[-1]
                    body_content = target_msg.html if target_msg.html else target_msg.text
                status.update(label="Draft found! Sending...", state="complete")

            # STEP B: PREPARE BCC LIST
            bcc_list = []
            if uploaded_file:
                df = pd.read_excel(uploaded_file, header=None)
                all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                # Skip header if first row isn't an email
                bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

            # STEP C: SENDING FUNCTION
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
                    server.login(from_email, target_password)
                    server.send_message(msg)

            # STEP D: EXECUTION
            if bcc_list:
                for i in range(0, len(bcc_list), int(batch_size)):
                    current_batch = bcc_list[i : i + int(batch_size)]
                    send_mail(current_batch)
                    st.write(f"✅ Sent Batch {(i // int(batch_size)) + 1}")
                    if i + int(batch_size) < len(bcc_list):
                        time.sleep(5) # Delay for company safety
                st.success(f"🎉 Successfully sent to {len(bcc_list)} BCC recipients!")
            else:
                send_mail([])
                st.success(f"✅ Email sent successfully to {to_email if to_email else from_email}!")

        except Exception as e:
            st.error(f"Error: {e}")