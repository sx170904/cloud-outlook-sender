import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- 1. CONFIGURATION ---
# These MUST be in your Streamlit Cloud Secrets
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common" 
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# MUST match your Azure Web Redirect URI exactly
REDIRECT_URI = "https://your-app-name.streamlit.app" 

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Pro Email Blaster", layout="wide")

# --- 2. AUTHENTICATION HELPER ---
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- 3. LOGIN INTERFACE ---
if 'token' not in st.session_state:
    st.title("📧 Outlook Bulk Email Sender")
    
    msal_app = get_msal_app()
    # We generate the login URL
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    
    st.info("Authentication Required")
    
    # This is the "Universal Fix" - A direct link is much harder for browsers to block
    st.markdown(f"""
        <div style="text-align: center; padding: 20px; border: 2px dashed #ccc; border-radius: 10px;">
            <h3>Step 1: Authorization</h3>
            <p>Click the link below to open the Microsoft Login page.</p>
            <a href="{auth_url}" target="_top" style="
                background-color: #0078d4; 
                color: white; 
                padding: 15px 30px; 
                text-decoration: none; 
                border-radius: 5px; 
                font-weight: bold;
                display: inline-block;
            ">👉 LOG IN WITH MICROSOFT OUTLOOK</a>
            <p style="margin-top: 15px; font-size: 0.8em; color: gray;">
                (If the page doesn't open, right-click the button and select 'Open in new tab')
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Handle the "Code" returned by Microsoft
    if "code" in st.query_params:
        with st.spinner("Finalizing Login..."):
            try:
                result = msal_app.acquire_token_by_authorization_code(
                    st.query_params["code"], 
                    scopes=SCOPES, 
                    redirect_uri=REDIRECT_URI
                )
                if "access_token" in result:
                    st.session_state.token = result["access_token"]
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error(f"Login Failed: {result.get('error_description')}")
            except Exception as e:
                st.error(f"An error occurred: {e}")
    st.stop()

# --- 4. MAIN APPLICATION INTERFACE ---
st.title("🚀 Pro Email Blaster")

with st.sidebar:
    st.header("⚙️ Settings")
    from_email = st.text_input("Send From (Optional)", placeholder="e.g. info@company.com")
    batch_size = st.number_input("Batch Size", min_value=1, value=50)
    delay_seconds = st.number_input("Batch Delay (Seconds)", min_value=1, value=10)
    
    if st.button("Log Out"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Email Content")
    draft_subject = st.text_input("Outlook Draft Subject", placeholder="Match your draft exactly")

with col2:
    st.subheader("2. Recipients")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# --- 5. EXECUTION LOGIC ---
if st.button("🔥 START EMAIL BLAST"):
    if not draft_subject or not uploaded_file:
        st.error("Missing Subject or Excel File!")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        # Search for Draft
        draft_query = f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
        response = requests.get(draft_query, headers=headers).json()

        if 'value' not in response or len(response['value']) == 0:
            st.error("Draft not found! Check the subject name and 'From' address.")
        else:
            body_content = response['value'][0]['body']['content']
            df = pd.read_excel(uploaded_file, header=None)
            recipient_list = df.iloc[:, 0].dropna().astype(str).tolist()
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i in range(0, len(recipient_list), batch_size):
                batch = recipient_list[i : i + batch_size]
                
                for recipient in batch:
                    payload = {
                        "message": {
                            "subject": draft_subject,
                            "body": {"contentType": "HTML", "content": body_content},
                            "toRecipients": [{"emailAddress": {"address": recipient}}]
                        }
                    }
                    send_res = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                    
                    if send_res.status_code == 202:
                        st.write(f"✅ Sent: {recipient}")
                    else:
                        st.error(f"❌ Failed: {recipient}")
                
                # Progress Update
                progress = min((i + batch_size) / len(recipient_list), 1.0)
                progress_bar.progress(progress)
                
                if i + batch_size < len(recipient_list):
                    status_text.warning(f"Waiting {delay_seconds}s for next batch...")
                    time.sleep(delay_seconds)
            
            status_text.success("🏁 Blasting Complete!")