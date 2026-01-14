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
# MUST match your Azure Web Redirect URI exactly
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

# ---------- UI (EXACTLY YOUR ORIGINAL DESIGN) ----------
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- LOGIN FLOW (WhatsApp Style) ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    
    # Check for returning code
    if "code" in st.query_params:
        msal_app = get_msal_app()
        result = msal_app.acquire_token_by_authorization_code(
            st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
            st.query_params.clear()
            # This JS forces the current window (the popup) to become Outlook
            st.markdown("<script>window.location.href='https://outlook.office.com/mail/';</script>", unsafe_allow_html=True)
            st.rerun()

    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")

    st.info("💡 Click the button below to link your Outlook. A separate window will open.")
    
    # JavaScript to open the controlled popup
    popup_js = f"""
    <script>
    function openOutlookLogin() {{
        const w = 1000, h = 800;
        const left = (window.screen.width/2)-(w/2), top = (window.screen.height/2)-(h/2);
        window.open('{auth_url}', 'OutlookLogin', `width=${{w}},height=${{h}},top=${{top}},left=${{left}}`);
    }}
    </script>
    <div style="text-align: center; padding: 20px;">
        <button onclick="openOutlookLogin()" style="
            background-color: #25D366; color: white; padding: 15px 35px; 
            border: none; border-radius: 30px; font-size: 18px; font-weight: bold; cursor: pointer;
        ">🔗 OPEN OUTLOOK & LOGIN</button>
    </div>
    """
    st.components.v1.html(popup_js, height=150)
    st.stop()

# ---------- LOGGED IN UI (YOUR ORIGINAL DESIGN) ----------
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    st.info("💡 A 5-second pause is applied between each batch for safety.")
    
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
    st.info("The Excel file should have emails in the **first column**.")
    uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

send_btn = st.button("🚀 Send Email(s)")

# ---------- LOGIC (YOUR ORIGINAL BATCHING LOGIC) ----------
if send_btn:
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        try:
            # Find Draft
            draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()

            if 'value' not in draft_res or len(draft_res['value']) == 0:
                st.error(f"Could not find draft: '{draft_subject}'")
            else:
                body_content = draft_res['value'][0]['body']['content']
                
                # Excel Recipient Logic
                bcc_list = []
                if uploaded_file:
                    df = pd.read_excel(uploaded_file, header=None)
                    all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                    if all_rows and "@" not in all_rows[0]:
                        bcc_list = all_rows[1:]
                    else:
                        bcc_list = all_rows

                if not to_email and not bcc_list:
                    st.error("No recipients found.")
                else:
                    # BATCHING LOGIC
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
                            if bcc_list:
                                st.write(f"✅ Sent Batch {batch_num} of {total_batches}")
                            else:
                                st.success(f"✅ Email sent successfully!")
                        else:
                            st.error(f"Error: {res.text}")

                        # YOUR 5-SECOND PAUSE
                        if batch_num < total_batches:
                            countdown = st.empty()
                            for s in range(5, 0, -1):
                                countdown.info(f"⏳ Waiting {s} seconds before next batch...")
                                time.sleep(1)
                            countdown.empty()
                    
                    if bcc_list: st.success("🎉 All batches sent successfully!")

        except Exception as e:
            st.error(f"An error occurred: {e}")