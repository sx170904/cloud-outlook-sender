import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- 1. CONFIGURATION ---
# Replace these with your actual IDs or ensure they are in Streamlit Secrets
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common"  # Works for both Personal and Work accounts
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# IMPORTANT: This must match your Azure Portal Redirect URI exactly
# Use "http://localhost:8501" for local testing
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
    st.info("Please log in to your Microsoft account to continue.")
    
    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    
    # HTML Button to break out of Streamlit's iframe
    login_btn_html = f"""
    <a href="{auth_url}" target="_top">
        <button style="
            background-color: #0078d4; color: white; padding: 12px 24px;
            border: none; border-radius: 4px; font-size: 16px; cursor: pointer;
        ">Log in with Microsoft Outlook</button>
    </a>
    """
    st.markdown(login_btn_html, unsafe_allow_html=True)

    # Handle the redirect logic
    if "code" in st.query_params:
        auth_code = st.query_params["code"]
        result = msal_app.acquire_token_by_authorization_code(
            auth_code, scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
            st.query_params.clear()
            st.rerun()
    st.stop()

# --- 4. MAIN APPLICATION INTERFACE ---
st.title("🚀 Pro Email Blaster")

with st.sidebar:
    st.header("⚙️ Settings")
    from_email = st.text_input("Send From (Optional)", placeholder="e.g. info@company.com", help="Leave blank to use your primary account.")
    batch_size = st.number_input("Batch Size", min_value=1, value=50, help="How many emails to send before pausing.")
    delay_seconds = st.number_input("Batch Delay (Seconds)", min_value=1, value=10, help="How long to wait between batches.")
    
    if st.button("Log Out"):
        del st.session_state.token
        st.rerun()

# Layout for inputs
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Email Content")
    draft_subject = st.text_input("Outlook Draft Subject", placeholder="Must match your draft exactly")
    st.caption("Tip: Create a draft in Outlook first with the body and subject you want.")

with col2:
    st.subheader("2. Recipients")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    st.caption("Ensure emails are in the first column (Column A).")

# --- 5. EXECUTION LOGIC ---
if st.button("🔥 START EMAIL BLAST"):
    if not draft_subject or not uploaded_file:
        st.error("Please provide both a Draft Subject and an Excel file.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        
        # Determine API Endpoint
        base_url = "https://graph.microsoft.com/v1.0/me"
        if from_email:
            base_url = f"https://graph.microsoft.com/v1.0/users/{from_email}"

        # Step A: Fetch the Draft
        with st.spinner("Searching for draft..."):
            draft_query = f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
            response = requests.get(draft_query, headers=headers).json()

        if 'value' not in response or len(response['value']) == 0:
            st.error(f"Could not find a draft with subject '{draft_subject}' in {from_email if from_email else 'your'} account.")
        else:
            draft_msg = response['value'][0]
            body_content = draft_msg['body']['content']
            
            # Step B: Load Recipients
            df = pd.read_excel(uploaded_file, header=None)
            recipient_list = df.iloc[:, 0].dropna().astype(str).tolist()
            total_emails = len(recipient_list)
            
            st.info(f"Loaded {total_emails} recipients. Starting batches...")
            
            # Step C: Send in Batches
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i in range(0, total_emails, batch_size):
                batch = recipient_list[i : i + batch_size]
                
                for index, recipient in enumerate(batch):
                    send_payload = {
                        "message": {
                            "subject": draft_subject,
                            "body": {"contentType": "HTML", "content": body_content},
                            "toRecipients": [{"emailAddress": {"address": recipient}}]
                        }
                    }
                    
                    send_res = requests.post(f"{base_url}/sendMail", headers=headers, json=send_payload)
                    
                    if send_res.status_code == 202:
                        st.write(f"✅ Sent: {recipient}")
                    else:
                        st.error(f"❌ Failed: {recipient} | {send_res.text}")
                
                # Progress and Batch Delay
                current_progress = min((i + batch_size) / total_emails, 1.0)
                progress_bar.progress(current_progress)
                
                if i + batch_size < total_emails:
                    status_text.warning(f"Batch complete. Waiting {delay_seconds}s before next batch...")
                    time.sleep(delay_seconds)
            
            status_text.success("🏁 All emails have been processed!")