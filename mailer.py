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
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" # MUST match Azure Web Redirect

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Outlook Pro Sender", layout="wide")

# --- 2. AUTHENTICATION HELPER ---
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- 3. LOGIN INTERFACE (The "WhatsApp Web" Style Popup) ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)

    st.info("Please link your Microsoft Outlook account to begin.")

    # JAVASCRIPT POPUP LOGIC
    # This creates a small controlled window just for the login
    popup_js = f"""
    <script>
    function openLoginPopup() {{
        const width = 600, height = 600;
        const left = (window.innerWidth / 2) - (width / 2);
        const top = (window.innerHeight / 2) - (height / 2);
        window.open('{auth_url}', 'MSAL_Login', `width=${{width}},height=${{height}},top=${{top}},left=${{left}},status=no,menubar=no,toolbar=no`);
    }}
    </script>
    <div style="text-align: center; padding: 50px;">
        <button onclick="openLoginPopup()" style="
            background-color: #25D366; 
            color: white; 
            padding: 15px 40px; 
            border: none; 
            border-radius: 30px; 
            font-size: 18px; 
            font-weight: bold; 
            cursor: pointer;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        ">
            🔗 OPEN OUTLOOK LOGIN
        </button>
        <p style="margin-top: 20px; color: #666;">A new window will open to authorize your account.</p>
    </div>
    """
    st.components.v1.html(popup_js, height=300)

    # Detect if we are returning from the login window
    if "code" in st.query_params:
        try:
            result = msal_app.acquire_token_by_authorization_code(
                st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
            )
            if "access_token" in result:
                st.session_state.token = result["access_token"]
                st.query_params.clear()
                st.rerun()
        except Exception as e:
            st.error(f"Login failed: {e}")
    st.stop()

# --- 4. MAIN UI (The Blasting Application) ---
st.title("🚀 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", help="Leave blank to use primary")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    if st.button("🔌 Disconnect Account"):
        del st.session_state.token
        st.rerun()

st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")

with col2:
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

if st.button("🚀 START EMAIL BLAST"):
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        try:
            # A. Find the Draft
            draft_query = f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
            r = requests.get(draft_query, headers=headers).json()

            if 'value' not in r or len(r['value']) == 0:
                st.error(f"Could not find draft: '{draft_subject}'")
            else:
                body_content = r['value'][0]['body']['content']
                
                # B. Excel Logic
                bcc_list = []
                if uploaded_file:
                    df = pd.read_excel(uploaded_file, header=None)
                    all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                    bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                # C. Batch Sending Logic
                if not to_email and not bcc_list:
                    st.error("No recipients found.")
                else:
                    total_emails = len(bcc_list)
                    for i in range(0, max(total_emails, 1), int(batch_size)):
                        batch = bcc_list[i:i + int(batch_size)]
                        batch_num = (i // int(batch_size)) + 1
                        
                        payload = {
                            "message": {
                                "subject": draft_subject,
                                "body": {"contentType": "HTML", "content": body_content},
                                "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                                "bccRecipients": [{"emailAddress": {"address": email}} for email in batch]
                            }
                        }
                        res = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        
                        if res.status_code == 202:
                            st.write(f"✅ Sent Batch {batch_num}")
                        else:
                            st.error(f"Error: {res.text}")

                        if (i + int(batch_size)) < total_emails:
                            st.info("⏳ Waiting 5 seconds...")
                            time.sleep(5)
                    st.success("🎉 Blasting Complete!")
        except Exception as e:
            st.error(f"Error: {e}")