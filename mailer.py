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

# --- 2. THE LOGIN BRIDGE (Runs in the New Tab) ---
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        # Force the tab to become Outlook immediately to avoid the iframe block
        st.markdown(f"""
            <script>window.top.location.href = "https://outlook.office.com/mail/";</script>
            <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
                <h2 style="color: #0078d4;">✅ Account Linked!</h2>
                <p>Redirecting to Outlook Inbox...</p>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. MAIN UI ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if 'token' in st.session_state:
        st.success("🟢 Connected")
        if st.button("🔌 Logout"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject", placeholder="Match your Outlook Draft name exactly")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")
with col2:
    st.info("Excel: Emails must be in the FIRST column.")
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

# --- 4. THE LOGIN vs SEND SWITCH ---
if 'token' not in st.session_state:
    # This script makes the app check for your login every 2 seconds
    st.markdown("""<script>setInterval(function(){ window.parent.location.reload(); }, 2000);</script>""", unsafe_allow_html=True)
    
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.warning("⚠️ Action Required: Link your account to enable sending.")
    
    st.markdown(f"""
        <div style="text-align: center; margin-top: 20px;">
            <a href="{auth_url}" target="_blank" style="
                background-color: #0078d4; color: white; padding: 20px 60px; 
                text-decoration: none; border-radius: 10px; font-weight: bold; 
                display: inline-block; font-size: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);
            ">🔗 CLICK TO LOGIN & OPEN OUTLOOK</a>
        </div>
    """, unsafe_allow_html=True)

else:
    # --- 5. THE FULL FEATURE SENDING ENGINE ---
    if st.button("🚀 START EMAIL BLAST"):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"
            try:
                # 1. Search for Draft
                draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()
                
                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"❌ Draft '{draft_subject}' not found.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    
                    # 2. Get Recipients from Excel
                    bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                        bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                    # 3. Batch Sending Loop
                    st.info(f"Processing in batches of {batch_size}...")
                    
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
                        
                        resp = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        
                        if resp.status_code == 202:
                            st.write(f"✅ Batch {i//batch_size + 1} Sent")
                        else:
                            st.error(f"Error in batch {i//batch_size + 1}: {resp.text}")
                        
                        if i + int(batch_size) < len(bcc_list):
                            time.sleep(5)
                            
                    st.success("🎉 All tasks finished!")
            except Exception as e:
                st.error(f"Error: {e}")