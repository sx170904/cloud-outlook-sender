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
# Ensure this matches your Azure Portal Redirect URI
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Outlook Universal Sender", layout="wide")

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- 2. THE POPUP & REDIRECT LOGIC ---

# Step A: Check if we are inside the Popup and just finished login
if "code" in st.query_params:
    msal_app = get_msal_app()
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.token = result["access_token"]
        
        # THIS IS THE PART THAT SHOWS MICROSOFT OUTLOOK
        # We use JavaScript to:
        # 1. Refresh the Main Opener window (the original web)
        # 2. Redirect THIS small window to Outlook Web
        st.markdown("""
            <script>
                if (window.opener) {
                    window.opener.location.reload(); 
                }
                window.location.href = 'https://outlook.office.com/mail/';
            </script>
            <div style="text-align:center; margin-top:50px;">
                <h2>Login Successful!</h2>
                <p>Redirecting this window to Outlook...</p>
            </div>
        """, unsafe_allow_html=True)
        st.stop()

# --- 3. LOGIN INTERFACE (Original Web View) ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(
        SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account"
    )

    st.info("💡 Click the button below. Login in the popup, and it will become your Outlook window.")
    
    # Popup JS
    popup_js = f"""
    <script>
    function openOutlookLogin() {{
        const w = 1100, h = 800;
        const left = (window.screen.width/2)-(w/2), top = (window.screen.height/2)-(h/2);
        window.open('{auth_url}', 'OutlookLogin', `width=${{w}},height=${{h}},top=${{top}},left=${{left}},resizable=yes,scrollbars=yes`);
    }}
    </script>
    <div style="text-align: center; padding: 30px;">
        <button onclick="openOutlookLogin()" style="
            background-color: #25D366; color: white; padding: 18px 40px; 
            border: none; border-radius: 35px; font-size: 20px; font-weight: bold; cursor: pointer;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        ">🔗 LOGIN & OPEN OUTLOOK</button>
    </div>
    """
    st.components.v1.html(popup_js, height=200)
    st.stop()

# --- 4. MAIN SENDER UI (Unlocked original web) ---
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Disconnect / Switch"):
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

send_btn = st.button("🚀 Send Email(s)")

# --- 5. SENDING LOGIC ---
if send_btn:
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        # Correctly format the URL for 'from_email'
        if from_email:
            base_url = f"https://graph.microsoft.com/v1.0/users/{from_email}"
        else:
            base_url = "https://graph.microsoft.com/v1.0/me"

        try:
            # Search for the draft
            draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()

            if 'value' not in draft_res or len(draft_res['value']) == 0:
                st.error(f"Could not find draft: '{draft_subject}' in the account.")
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
                    
                    st.success("🎉 All tasks complete!")

        except Exception as e:
            st.error(f"An error occurred: {e}")