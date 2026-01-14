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

# --- 2. THE POPUP BREAKOUT (THIS FIXES THE BUTTON NO RESPONSE) ---
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        
        # We use window.top.location.href to break out of the iframe
        st.markdown(f"""
            <div style="text-align:center; margin-top:50px; font-family: sans-serif;">
                <h1 style="color: #25D366;">✅ Login Successful</h1>
                <p style="font-size: 18px;">Click the button below to switch this window to Outlook.</p>
                <button onclick="window.top.location.href='https://outlook.office.com/mail/';" style="
                    background-color: #0078d4; color: white; padding: 20px 40px; 
                    border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 20px;
                ">🚀 OPEN OUTLOOK INBOX</button>
            </div>
            <script>
                // Auto-jump attempt using Top-Level navigation
                setTimeout(function(){{
                    window.top.location.href = 'https://outlook.office.com/mail/';
                }}, 1000);
            </script>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. MAIN SENDER UI (YOUR ORIGINAL DESIGN) ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")

st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if 'token' in st.session_state:
        if st.button("🔌 Logout / Switch Account"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")

with col2:
    st.info("The Excel file should have emails in the **first column**.")
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

# --- 4. DYNAMIC BUTTON LOGIC ---
if 'token' not in st.session_state:
    # Refresh the main page every 4 seconds to detect the login
    st.markdown("""<script>setInterval(function(){ window.parent.location.reload(); }, 4000);</script>""", unsafe_allow_html=True)
    
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.warning("⚠️ Action Required: Please link your Outlook account.")
    
    login_html = f"""
    <div style="text-align: center;">
        <button onclick="window.open('{auth_url}', 'OutlookApp', 'width=1100,height=850')" style="
            background-color: #0078d4; color: white; padding: 18px 45px; 
            border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 18px;">
            🔗 LOGIN & OPEN OUTLOOK
        </button>
    </div>
    """
    st.components.v1.html(login_html, height=100)

else:
    # --- 5. YOUR ORIGINAL SENDING FUNCTION ---
    if st.button("🚀 Send Email(s)"):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

            try:
                draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()

                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"Could not find draft: '{draft_subject}'.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                        bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                    if not to_email and not bcc_list:
                        st.error("No recipients found.")
                    else:
                        for i in range(0, max(len(bcc_list), 1), int(batch_size)):
                            batch = bcc_list[i : i + int(batch_size)]
                            payload = {
                                "message": {
                                    "subject": draft_subject,
                                    "body": {"contentType": "HTML", "content": body_content},
                                    "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                    "bccRecipients": [{"emailAddress": {"address": e}} for e in batch]
                                }
                            }
                            requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                            st.write(f"✅ Batch sent successfully.")
                            time.sleep(5)
                        st.success("🎉 Complete!")
            except Exception as e:
                st.error(f"Error: {e}")