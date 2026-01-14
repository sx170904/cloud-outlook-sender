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

# --- 2. POPUP HANDLER (Runs inside the small window after login) ---
if "code" in st.query_params:
    msal_app = get_msal_app()
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        
        # This UI shows in the small window briefly before jumping to Outlook
        st.markdown(f"""
            <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
                <h2 style="color: #25D366;">✅ Login Successful!</h2>
                <p>Redirecting this window to your Outlook Inbox...</p>
                <script>
                    // Redirect the small window to the actual Outlook Web App
                    setTimeout(function(){{
                        window.location.href = 'https://outlook.office.com/mail/';
                    }}, 1500);
                </script>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. LOGIN INTERFACE (Runs on the main web page) ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    # AUTO-DETECT SCRIPT: The main page "pings" itself to see if the token exists
    # This prevents the original page from staying stuck on the login screen
    st.markdown("""
        <script>
        var checkLogin = setInterval(function() {
            // We refresh the main page to check if the session_state.token is now set
            window.parent.location.reload();
        }, 4000); 
        </script>
    """, unsafe_allow_html=True)

    msal_app = get_msal_app()
    # prompt="select_account" allows you to switch accounts every time
    auth_url = msal_app.get_authorization_request_url(
        SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account"
    )

    st.info("💡 Link your account. Once authorized, Outlook will open here in a new window.")
    
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

# --- 4. MAIN SENDER UI (Unlocked original web page) ---
st.title("🚀 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Optional")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Disconnect / Switch Account"):
        # Fully clears session and redirects to Microsoft Logout
        logout_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri={REDIRECT_URI}"
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.markdown(f'<meta http-equiv="refresh" content="0;URL=\'{logout_url}\'" />', unsafe_allow_html=True)

st.subheader("2. Draft & Recipients")
draft_subject = st.text_input("Draft Email Subject", placeholder="Match your Outlook Draft exactly")

col1, col2 = st.columns(2)
with col1:
    to_email = st.text_input("To (Optional)")
    cc_email = st.text_input("CC (Optional)")

with col2:
    st.info("The Excel file should have emails in the **first column**.")
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

send_btn = st.button("🚀 START EMAIL BLAST")

# --- 5. LOGIC (Your original win32 logic ported to Web API) ---
if send_btn:
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        try:
            # Find Draft by Subject
            draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()

            if 'value' not in draft_res or len(draft_res['value']) == 0:
                st.error(f"Draft '{draft_subject}' not found. Check the Outlook window to verify the name.")
            else:
                body_content = draft_res['value'][0]['body']['content']
                
                # Excel Logic
                bcc_list = []
                if uploaded_file:
                    df = pd.read_excel(uploaded_file, header=None)
                    all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                    bcc_list = all_rows[1:] if all_rows and "@" not in all_rows[0] else all_rows

                if not to_email and not bcc_list:
                    st.error("No recipients found.")
                else:
                    total_batches = (len(bcc_list) + batch_size - 1) // batch_size if bcc_list else 1
                    
                    for i in range(0, max(len(bcc_list), 1), int(batch_size)):
                        batch = bcc_list[i:i + int(batch_size)]
                        batch_num = (i // int(batch_size)) + 1
                        
                        payload = {
                            "message": {
                                "subject": draft_subject,
                                "body": {"contentType": "HTML", "content": body_content},
                                "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                                "bccRecipients": [{"emailAddress": {"address": e}} for e in batch]
                            }
                        }
                        
                        res = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        
                        if res.status_code == 202:
                            st.write(f"✅ Sent Batch {batch_num} of {total_batches}")
                        else:
                            st.error(f"Error: {res.text}")

                        if batch_num < total_batches:
                            time.sleep(5)
                    
                    st.success("🎉 Email Blast Completed Successfully!")

        except Exception as e:
            st.error(f"Critical Error: {e}")