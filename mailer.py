import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- 1. CONFIGURATION ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

# --- 2. THE BRIDGE (Handles the "Return" from Microsoft) ---
# When you login in the NEW TAB, Microsoft sends you back here.
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        # This tab now turns into the real Outlook website to avoid iframe errors
        st.markdown(f"""
            <script>window.top.location.href = "https://outlook.office.com/mail/";</script>
            <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
                <h2 style="color: #0078d4;">✅ Account Linked!</h2>
                <p>Opening your Outlook Inbox now...</p>
                <p><b>Go back to your original tab to start sending!</b></p>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. MAIN UI CONFIGURATION ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if 'token' in st.session_state:
        st.success("🟢 Connected to Outlook")
        if st.button("🔌 Logout / Reset"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

# --- 4. SENDER FORM (ALWAYS VISIBLE) ---
st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject", placeholder="Enter the exact subject of your Outlook Draft")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)", placeholder="direct.recipient@example.com")
    cc_email = st.text_input("CC (Optional)")
with col2:
    st.info("Upload Excel: Emails must be in the first column.")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

# --- 5. THE SWITCH (LOGIN vs SEND BUTTON) ---
if 'token' not in st.session_state:
    st.warning("⚠️ Action Required: You must link your Outlook account to enable sending.")
    
    # Generate the Microsoft Login URL
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    # This link opens in a NEW TAB to prevent "Refused to Connect"
    st.markdown(f"""
        <div style="text-align: center; margin-top: 20px;">
            <a href="{auth_url}" target="_blank" style="
                background-color: #0078d4; color: white; padding: 20px 50px; 
                text-decoration: none; border-radius: 8px; font-weight: bold; 
                display: inline-block; font-size: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            ">🔗 CLICK TO LOGIN & OPEN OUTLOOK</a>
        </div>
    """, unsafe_allow_html=True)
    
    # Button to refresh the main page once the user is done logging in
    if st.button("🔄 I have logged in - Unlock Send Button"):
        st.rerun()

else:
    # --- 6. THE SENDING LOGIC (EXACTLY YOUR BUSINESS LOGIC) ---
    if st.button("🚀 START EMAIL BLAST"):
        if not draft_subject:
            st.error("Please enter the Draft Email Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            # Handle "Send From" specific user or 'me'
            base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"
            
            try:
                # 1. Search for the Draft
                draft_res = requests.get(
                    f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", 
                    headers=headers
                ).json()

                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"❌ Could not find draft with subject: '{draft_subject}'. Please check your Outlook Drafts folder.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    
                    # 2. Get Recipients from Excel
                    bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                        # Simple logic: skip header if no '@' found in first cell
                        bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                    if not to_email and not bcc_list:
                        st.error("❌ No recipients found in 'To' field or Excel file.")
                    else:
                        # 3. Batch Sending Loop
                        total_emails = len(bcc_list)
                        st.info(f"Processing {total_emails} emails in batches of {batch_size}...")
                        
                        for i in range(0, max(len(bcc_list), 1), int(batch_size)):
                            batch = bcc_list[i : i + int(batch_size)]
                            payload = {
                                "message": {
                                    "subject": draft_subject,
                                    "body": {"contentType": "HTML", "content": body_content},
                                    "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                    "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                                    "bccRecipients": [{"emailAddress": {"address": e}} for e in batch]
                                }
                            }
                            # Send the mail via Graph API
                            resp = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                            
                            if resp.status_code == 202:
                                st.write(f"✅ Batch {i//batch_size + 1} sent successfully.")
                            else:
                                st.error(f"Failed to send batch: {resp.text}")
                            
                            # Pause to avoid spam filters
                            if i + int(batch_size) < len(bcc_list):
                                time.sleep(5)
                        
                        st.success("🎉 All emails have been sent successfully!")
            except Exception as e:
                st.error(f"An error occurred: {e}")