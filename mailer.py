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

# --- 2. THE ERROR-PROOF TOKEN HANDLER ---
if "code" in st.query_params:
    try:
        msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
        # We exchange the code for a token
        result = msal_app.acquire_token_by_authorization_code(
            st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
        
        # Clean the URL and refresh to show the Sender UI
        st.query_params.clear()
        st.rerun()
    except Exception:
        # If the code already expired or failed, just clear it and let the user try again
        st.query_params.clear()
        st.rerun()

# --- 3. UI SETUP ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")
st.title("📧 Outlook Universal Sender")

# SIDEBAR
with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if 'token' in st.session_state:
        st.success("🟢 Connected")
        if st.button("🔌 Logout / Switch"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

# --- 4. PAGE SWITCH ---
if 'token' not in st.session_state:
    # --- LOGIN VIEW ---
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.info("👋 Welcome! Please login to unlock the Sender tools.")
    st.markdown(f"""
        <div style="text-align: center; margin-top: 40px;">
            <a href="{auth_url}" target="_top" style="
                background-color: #0078d4; color: white; padding: 25px 80px; 
                text-decoration: none; border-radius: 12px; font-weight: bold; 
                display: inline-block; font-size: 22px;
            ">🔑 LOGIN TO OUTLOOK</a>
        </div>
    """, unsafe_allow_html=True)
else:
    # --- SENDER VIEW (All Features Restored) ---
    st.subheader("2. Draft & Recipients")
    draft_subject = st.text_input("Draft Email Subject", placeholder="Match your Outlook Draft name exactly")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Optional)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        st.info("Upload Excel: Emails must be in the FIRST column.")
        uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if st.button("🚀 START EMAIL BLAST", type="primary", use_container_width=True):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            # Use 'me' or specific email
            user_path = f"users/{from_email}" if from_email else "me"
            base_url = f"https://graph.microsoft.com/v1.0/{user_path}"
            
            try:
                # 1. Search for Draft
                draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()
                
                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"❌ Draft '{draft_subject}' not found. Check your Outlook Drafts.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    
                    # 2. Get Excel Data
                    bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                        bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                    # 3. Batch Sending Loop
                    st.info(f"Sending to {len(bcc_list)} recipients...")
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
                            st.write(f"✅ Sent Batch {i//batch_size + 1}")
                        else:
                            st.error(f"Failed: {r.text}")
                        
                        if i + int(batch_size) < len(bcc_list):
                            time.sleep(5)
                    st.success("🎉 Process Complete!")
            except Exception as e:
                st.error(f"Connection Error: {e}")