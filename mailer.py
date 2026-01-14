import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- CONFIG ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.title("📧 Custom Email Sender")

# --- LOGIN ---
if 'token' not in st.session_state:
    if st.button("🔑 Login to Microsoft"):
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        flow = app.initiate_device_flow(scopes=SCOPES)
        st.info(f"Go to: {flow['verification_uri']} and enter: {flow['user_code']}")
        with st.spinner("Waiting..."):
            result = app.acquire_token_by_device_flow(flow)
            if "access_token" in result:
                st.session_state.token = result["access_token"]
                st.rerun()

# --- SENDER UI ---
if 'token' in st.session_state:
    st.subheader("Configuration")
    
    # NEW: Input for the 'From' address
    from_email = st.text_input("Send From (Your email or shared mailbox)", help="Must be an email you have permission to use.")
    
    draft_subject = st.text_input("Outlook Draft Subject")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Main Recipient)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        uploaded_file = st.file_uploader("Upload Excel for BCC", type=["xlsx"])

    if st.button("🚀 Start Email Blast"):
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        
        # 1. Get the Draft
        # If sending from a different mailbox, we change the URL from /me/ to /users/EMAIL/
        base_url = "https://graph.microsoft.com/v1.0/me"
        if from_email:
            base_url = f"https://graph.microsoft.com/v1.0/users/{from_email}"

        draft_url = f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
        r = requests.get(draft_url, headers=headers).json()
        
        if 'error' in r:
            st.error(f"Error: {r['error']['message']}")
        elif not r.get('value'):
            st.error(f"Draft not found in {from_email if from_email else 'your'} account.")
        else:
            body = r['value'][0]['body']['content']
            df = pd.read_excel(uploaded_file, header=None)
            bcc_list = df.iloc[:, 0].dropna().astype(str).tolist()

            for email in bcc_list:
                send_payload = {
                    "message": {
                        "subject": draft_subject,
                        "body": {"contentType": "HTML", "content": body},
                        "toRecipients": [{"emailAddress": {"address": to_email if to_email else from_email}}],
                        "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                        "bccRecipients": [{"emailAddress": {"address": email}}]
                    }
                }
                # Send the mail via the chosen 'From' account
                send_url = f"{base_url}/sendMail"
                requests.post(send_url, headers=headers, json=send_payload)
                st.write(f"✅ Sent to {email}")
                time.sleep(1)
            st.success("Complete!")