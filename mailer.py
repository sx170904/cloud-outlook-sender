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

# --- 2. THE LOGIN TAB HANDLER ---
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        # Success message for the NEW TAB only
        st.markdown("""
            <div style="text-align:center; margin-top:100px; font-family:sans-serif;">
                <h1 style="color: #25D366;">✅ Login Successful</h1>
                <p style="font-size: 20px;">You can now close this tab and return to the original window.</p>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. MAIN APP UI ---
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

# --- 4. THE MANUAL FLOW GATE ---
if 'token' not in st.session_state:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.info("👋 Follow these steps to connect:")
    
    # FORCED NEW TAB BUTTON (HTML)
    st.markdown(f"""
        <div style="text-align: center; margin: 30px 0;">
            <a href="{auth_url}" target="_blank" style="
                background-color: #0078d4; color: white; padding: 20px 50px; 
                text-decoration: none; border-radius: 10px; font-weight: bold; 
                display: inline-block; font-size: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            ">1. CLICK TO LOGIN (OPENS NEW TAB)</a>
        </div>
    """, unsafe_allow_html=True)

    # MANUAL CONNECT BUTTON
    if st.button("2. ✅ I HAVE LOGGED IN - ACTIVATE SENDER", use_container_width=True, type="primary"):
        # This forces the app to check the session for the token saved by the other tab
        st.rerun()

else:
    # --- 5. THE SENDER UI (FULL FEATURES) ---
    st.subheader("2. Draft & Recipients")
    draft_subject = st.text_input("Draft Email Subject", placeholder="Match your Outlook Draft name exactly")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Optional)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

    if st.button("🚀 START EMAIL BLAST", use_container_width=True):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"
            try:
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
                        requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        st.write(f"✅ Batch {i//batch_size + 1} Sent")
                        time.sleep(5)
                    st.success("🎉 Process Finished!")
            except Exception as e:
                st.error(f"Error: {e}")