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
# Use your app's actual URL
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Outlook Pro Sender", layout="wide")

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- 2. LOGIN & POPUP LOGIC ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    # Check if the "Code" is in the URL (The Popup sent it back)
    if "code" in st.query_params:
        msal_app = get_msal_app()
        result = msal_app.acquire_token_by_authorization_code(
            st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
            # REFRESH THE MAIN PAGE TO UNLOCK UI
            st.query_params.clear()
            
            # This JS tells the small window to go to Outlook, while the main app reruns
            st.markdown("""
                <script>
                window.location.href = 'https://outlook.office.com/mail/';
                </script>
            """, unsafe_allow_html=True)
            st.rerun()

    msal_app = get_msal_app()
    # Force account selection
    auth_url = msal_app.get_authorization_request_url(
        SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account"
    )

    st.info("Link your account to see the Sender UI here, while Outlook opens in a new window.")

    # WhatsApp-style Popup Logic
    popup_js = f"""
    <script>
    function openProLogin() {{
        const width = 1000, height = 800;
        const left = (window.screen.width / 2) - (width / 2);
        const top = (window.screen.height / 2) - (height / 2);
        // This opens the window. Once login is done, it redirects back to the main app URL,
        // which triggers the 'code' detection above.
        window.open('{auth_url}', 'OutlookLogin', `width=${{width}},height=${{height}},top=${{top}},left=${{left}}`);
    }}
    </script>
    <div style="text-align: center; padding: 40px;">
        <button onclick="openProLogin()" style="
            background-color: #0078d4; color: white; padding: 18px 45px; 
            border: none; border-radius: 50px; font-size: 20px; font-weight: bold; 
            cursor: pointer; box-shadow: 0 4px 15px rgba(0,120,212,0.4);
        ">
            🔗 CONNECT & OPEN OUTLOOK
        </button>
    </div>
    """
    st.components.v1.html(popup_js, height=200)
    st.stop()

# --- 3. MAIN SENDER UI (Unlocked after login) ---
st.title("🚀 Pro Email Blaster Control Panel")
st.success("✅ Account Connected. Outlook is open in your other window.")

with st.sidebar:
    st.header("⚙️ Settings")
    from_email = st.text_input("Send From (Optional)")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Disconnect & Switch Account"):
        # Global Logout logic
        logout_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri={REDIRECT_URI}"
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.markdown(f'<meta http-equiv="refresh" content="0;URL=\'{logout_url}\'" />', unsafe_allow_html=True)

# Layout for inputs
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Email Content")
    draft_subject = st.text_input("Outlook Draft Subject", placeholder="Check your other window for the subject name")

with col2:
    st.subheader("2. Recipients")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# --- 4. EXECUTION LOGIC ---
if st.button("🔥 START EMAIL BLAST"):
    if not draft_subject or not uploaded_file:
        st.error("Provide Subject and Excel file.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        with st.spinner("Searching for draft in your Outlook..."):
            draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()

        if 'value' not in draft_res or len(draft_res['value']) == 0:
            st.error("Draft not found! Make sure the subject matches exactly in the Outlook window.")
        else:
            body = draft_res['value'][0]['body']['content']
            df = pd.read_excel(uploaded_file, header=None)
            recipients = df.iloc[:, 0].dropna().astype(str).tolist()
            
            # Skip header if necessary
            if recipients and "@" not in recipients[0]:
                recipients = recipients[1:]

            pbar = st.progress(0)
            for i in range(0, len(recipients), int(batch_size)):
                batch = recipients[i : i + int(batch_size)]
                payload = {
                    "message": {
                        "subject": draft_subject,
                        "body": {"contentType": "HTML", "content": body},
                        "bccRecipients": [{"emailAddress": {"address": e}} for e in batch]
                    }
                }
                requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                st.write(f"✅ Batch sent to {len(batch)} people.")
                pbar.progress(min((i + int(batch_size)) / len(recipients), 1.0))
                if i + int(batch_size) < len(recipients):
                    time.sleep(5)
            st.success("🏁 All batches sent!")