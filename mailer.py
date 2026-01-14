import streamlit as st
import msal
import requests
import pandas as pd
import time
import os

# --- 1. CONFIGURATION ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

# File path to sync login between tabs
TOKEN_FILE = "session_token.txt"

# --- 2. THE LOGIN HANDLER (Runs in Tab B) ---
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        # Save token to a file so Tab A can see it
        with open(TOKEN_FILE, "w") as f:
            f.write(result["access_token"])
        
        st.markdown("""
            <div style="text-align:center; margin-top:100px; font-family:sans-serif;">
                <h1 style="color: #25D366;">✅ LOGIN SUCCESSFUL</h1>
                <p style="font-size: 20px;">You can now close this tab.</p>
                <p>Go back to the original page and click <b>ACTIVATE SENDER</b>.</p>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. MAIN APP UI ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")
st.title("📧 Outlook Universal Sender")

# Sidebar for settings
with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Force Logout"):
        if os.path.exists(TOKEN_FILE):
            os.remove(TOKEN_FILE)
        if 'token' in st.session_state:
            del st.session_state.token
        st.rerun()

# --- 4. THE SYNC GATE ---
# Check if a token exists in the file (from Tab B)
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE, "r") as f:
        st.session_state.token = f.read()

# UI Switch
if 'token' not in st.session_state:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.info("👋 Follow these steps carefully:")
    
    # Force New Tab
    st.markdown(f"""
        <div style="text-align: center; margin: 30px 0;">
            <a href="{auth_url}" target="_blank" style="
                background-color: #0078d4; color: white; padding: 20px 50px; 
                text-decoration: none; border-radius: 10px; font-weight: bold; 
                display: inline-block; font-size: 20px;
            ">1. LOGIN (NEW TAB)</a>
        </div>
    """, unsafe_allow_html=True)

    if st.button("2. ✅ ACTIVATE SENDER", use_container_width=True, type="primary"):
        # This forces the app to look at the 'session_token.txt' file
        st.rerun()

else:
    # --- 5. SENDER UI (FULL FEATURES) ---
    st.success("🟢 Account Linked Successfully")
    st.subheader("2. Draft & Recipients")
    draft_subject = st.text_input("Draft Email Subject", placeholder="Match your Outlook Draft name exactly")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Optional)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

    if st.button("🚀 START EMAIL BLAST", use_container_width=True, type="primary"):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"
            try:
                # Find Draft
                draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()
                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"❌ Draft '{draft_subject}' not found.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                        bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                    # Sending Loop
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
                        r = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        if r.status_code == 202:
                            st.write(f"✅ Batch {i//batch_size + 1} Sent")
                        else:
                            st.error(f"Error: {r.text}")
                        time.sleep(5)
                    st.success("🎉 Process Finished!")
            except Exception as e:
                st.error(f"Error: {e}")