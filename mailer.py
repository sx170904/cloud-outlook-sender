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

# YOUR SPECIFIC REDIRECT URL
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Outlook Universal Sender", layout="wide")

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- 2. THE POPUP REDIRECT (This runs in the small window) ---
if "code" in st.query_params:
    msal_app = get_msal_app()
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        
        # This script forces the small window to BECOME Outlook immediately
        st.markdown("""
            <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
                <h2 style="color: #0078d4;">Opening Outlook...</h2>
                <p>Please wait while we load your inbox.</p>
            </div>
            <script>
                // Using replace ensures the browser doesn't keep the "Login Successful" page in history
                window.location.replace('https://outlook.office.com/mail/');
            </script>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. THE MAIN PAGE HANDLER ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    # This keeps the main page checking for the login every 3 seconds
    st.markdown("""
        <script>
        setInterval(function() {
            window.parent.location.reload();
        }, 3000); 
        </script>
    """, unsafe_allow_html=True)

    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(
        SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account"
    )

    st.info("💡 Link your account. Once you sign in, this window will unlock and the popup will show Outlook.")
    
    popup_js = f"""
    <script>
    function openOutlookLogin() {{
        const w = 1100, h = 850;
        const left = (window.screen.width/2)-(w/2), top = (window.screen.height/2)-(h/2);
        window.open('{auth_url}', 'OutlookLogin', `width=${{w}},height=${{h}},top=${{top}},left=${{left}},resizable=yes`);
    }}
    </script>
    <div style="text-align: center; padding: 20px;">
        <button onclick="openOutlookLogin()" style="
            background-color: #25D366; color: white; padding: 20px 50px; 
            border: none; border-radius: 40px; font-size: 20px; font-weight: bold; cursor: pointer;
            box-shadow: 0 4px 15px rgba(37, 211, 102, 0.4);
        ">🔗 OPEN OUTLOOK & LOGIN</button>
    </div>
    """
    st.components.v1.html(popup_js, height=200)
    st.stop()

# --- 4. SENDER UI (Unlocked original web page) ---
# (The rest of the UI and Logic stays exactly the same)
st.title("🚀 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Optional")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Disconnect / Switch Account"):
        logout_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri={REDIRECT_URI}"
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.markdown(f'<meta http-equiv="refresh" content="0;URL=\'{logout_url}\'" />', unsafe_allow_html=True)

st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")

with col2:
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

send_btn = st.button("🚀 START EMAIL BLAST")

if send_btn:
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"
        try:
            draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()
            if 'value' not in draft_res or len(draft_res['value']) == 0:
                st.error(f"Draft '{draft_subject}' not found.")
            else:
                body_content = draft_res['value'][0]['body']['content']
                bcc_list = []
                if uploaded_file:
                    df = pd.read_excel(uploaded_file, header=None)
                    all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                    bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                for i in range(0, len(bcc_list) if bcc_list else 1, int(batch_size)):
                    batch = bcc_list[i : i + int(batch_size)] if bcc_list else []
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
                    st.write(f"✅ Sent batch starting at index {i}")
                    time.sleep(5)
                st.success("🎉 Process Finished!")
        except Exception as e:
            st.error(f"Error: {e}")