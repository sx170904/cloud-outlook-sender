import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- 1. SETTINGS ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
# We use 'common' for personal accounts
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Email Sender", layout="wide")
st.title("📧 Outlook Email Blaster")

# --- 2. THE LOGIN LOGIC ---
if 'token' not in st.session_state:
    st.info("You need to link your Outlook account to start.")
    if st.button("🔑 Get Login Code"):
        # Create the MSAL app instance
        client = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        
        # Initiate the device flow
        flow = client.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            st.error("Error starting login. Check your Client ID in Secrets.")
        else:
            st.markdown(f"1. Go to: **{flow['verification_uri']}**")
            st.markdown(f"2. Enter this code: :blue[**{flow['user_code']}**]")
            
            # Wait for the user to finish login in the browser
            with st.spinner("Waiting for you to authorize..."):
                result = client.acquire_token_by_device_flow(flow)
                if "access_token" in result:
                    st.session_state.token = result["access_token"]
                    st.success("Login Successful!")
                    st.rerun()

# --- 3. THE SENDER INTERFACE ---
if 'token' in st.session_state:
    st.sidebar.success("Logged In")
    if st.sidebar.button("Log Out"):
        del st.session_state.token
        st.rerun()

    # User Inputs
    st.subheader("Message Details")
    draft_subject = st.text_input("Outlook Draft Subject", help="Must match your Draft subject exactly")
    
    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Main Recipient)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        uploaded_file = st.file_uploader("Upload Excel (Emails in 1st column)", type=["xlsx"])

    if st.button("🚀 Start Email Blast"):
        if not draft_subject or not uploaded_file:
            st.warning("Please enter a subject and upload an Excel file.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            
            # A. Fetch the Draft content
            draft_url = f"https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
            draft_res = requests.get(draft_url, headers=headers).json()
            
            if not draft_res.get('value'):
                st.error("Draft not found! Check the subject spelling in your Outlook Drafts folder.")
            else:
                body = draft_res['value'][0]['body']['content']
                
                # B. Read Excel
                df = pd.read_excel(uploaded_file, header=None)
                # Filter out empty rows and non-email strings
                bcc_emails = df.iloc[:, 0].dropna().astype(str).tolist()
                
                # C. Send Loop
                progress_bar = st.progress(0)
                for idx, email in enumerate(bcc_emails):
                    send_payload = {
                        "message": {
                            "subject": draft_subject,
                            "body": {"contentType": "HTML", "content": body},
                            "toRecipients": [{"emailAddress": {"address": to_email if to_email else "me@outlook.com"}}],
                            "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                            "bccRecipients": [{"emailAddress": {"address": email}}]
                        }
                    }
                    
                    r = requests.post("https://graph.microsoft.com/v1.0/me/sendMail", headers=headers, json=send_payload)
                    
                    if r.status_code == 202:
                        st.write(f"✅ Sent: {email}")
                    else:
                        st.error(f"❌ Error for {email}: {r.text}")
                    
                    # Update progress
                    progress_bar.progress((idx + 1) / len(bcc_emails))
                    time.sleep(1) # Safety delay to avoid spam blocks
                
                st.success(f"Done! Sent to {len(bcc_emails)} recipients.")