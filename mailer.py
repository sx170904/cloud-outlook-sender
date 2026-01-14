import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- CONFIG ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = st.secrets["MS_TENANT_ID"] # Use 'common'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# Use your actual Streamlit URL here
REDIRECT_URI = "https://your-app-name.streamlit.app" 
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Pro Email Blaster", layout="wide")

# --- MSAL HELPER ---
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

# --- AUTHENTICATION FLOW ---
if 'token' not in st.session_state:
    st.title("🔒 Access Required")
    st.write("Please log in with your Microsoft account to use the blasting tool.")
    
    # Create the login URL
    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    
    # This button opens the Microsoft login page in a new tab
    st.markdown(f'<a href="{auth_url}" target="_self"><button style="background-color:#0078d4;color:white;padding:10px 24px;border:none;border-radius:4px;cursor:pointer;">Login with Microsoft Outlook</button></a>', unsafe_allow_html=True)

    # Handle the redirect back from Microsoft
    query_params = st.query_params
    if "code" in query_params:
        result = msal_app.acquire_token_by_authorization_code(
            query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
            st.rerun()
    st.stop()

# --- MAIN APP INTERFACE ---
st.title("🚀 Pro Email Blaster")

with st.sidebar:
    st.success("Account Authenticated")
    from_email = st.text_input("Send From (Email Alias/Shared)", placeholder="leave blank for default")
    batch_size = st.number_input("Batch Size (Emails per group)", min_value=1, value=50)
    delay_seconds = st.number_input("Delay between batches (seconds)", min_value=1, value=10)
    if st.button("Logout"):
        del st.session_state.token
        st.rerun()

# Sender UI
col1, col2 = st.columns(2)
with col1:
    draft_subject = st.text_input("Outlook Draft Subject")
    to_display = st.text_input("Display 'To' Name (optional)", "Valued Customer")

with col2:
    uploaded_file = st.file_uploader("Upload Excel (Emails in 1st Column)", type=["xlsx"])

if st.button("Start Blasting"):
    if not draft_subject or not uploaded_file:
        st.error("Missing Draft Subject or Excel File!")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        
        # Determine Base URL (Me or another User)
        base_endpoint = "https://graph.microsoft.com/v1.0/me"
        if from_email:
            base_endpoint = f"https://graph.microsoft.com/v1.0/users/{from_email}"

        # 1. Get Draft Content
        draft_url = f"{base_endpoint}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
        draft_data = requests.get(draft_url, headers=headers).json()

        if 'value' not in draft_data or len(draft_data['value']) == 0:
            st.error("Draft not found. Check spelling and 'From' address.")
        else:
            body_content = draft_data['value'][0]['body']['content']
            df = pd.read_excel(uploaded_file, header=None)
            emails = df.iloc[:, 0].dropna().astype(str).tolist()

            # 2. Loop with Batching
            for i in range(0, len(emails), batch_size):
                current_batch = emails[i:i + batch_size]
                st.write(f"📦 Processing batch {i//batch_size + 1}...")
                
                for recipient in current_batch:
                    payload = {
                        "message": {
                            "subject": draft_subject,
                            "body": {"contentType": "HTML", "content": body_content},
                            "toRecipients": [{"emailAddress": {"address": recipient}}],
                        }
                    }
                    send_url = f"{base_endpoint}/sendMail"
                    res = requests.post(send_url, headers=headers, json=payload)
                    
                    if res.status_code == 202:
                        st.write(f"✅ Sent: {recipient}")
                    else:
                        st.error(f"❌ Failed {recipient}: {res.text}")
                    
                if i + batch_size < len(emails):
                    st.info(f"⏳ Waiting {delay_seconds}s for next batch...")
                    time.sleep(delay_seconds)
            
            st.success("All batches completed!")