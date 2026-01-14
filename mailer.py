import streamlit as st
import msal
import requests
import pandas as pd
import time

# --- 1. CONFIGURATION (Must be in Streamlit Secrets) ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common" 
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" # Update this!

SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

st.set_page_config(page_title="Outlook Universal Sender", layout="wide")

# --- 2. AUTHENTICATION LOGIC ---
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

if 'token' not in st.session_state:
    st.title("📧 Outlook Universal Sender")
    st.warning("Please sign in to connect your Outlook account.")
    
    msal_app = get_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    
    # The "Window Pop-out" Login Button
    st.markdown(f"""
        <a href="{auth_url}" target="_top" style="
            background-color: #0078d4; color: white; padding: 12px 24px;
            text-decoration: none; border-radius: 4px; font-weight: bold;
        ">Click to Login with Microsoft Outlook</a>
    """, unsafe_allow_html=True)

    if "code" in st.query_params:
        result = msal_app.acquire_token_by_authorization_code(
            st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            st.session_state.token = result["access_token"]
            st.rerun()
    st.stop()

# --- 3. UI (EXACTLY LIKE YOUR ORIGINAL CODE) ---
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", help="Leave blank to use logged-in account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    st.info("💡 A 5-second pause is applied between each batch for safety.")
    if st.button("Logout"):
        del st.session_state.token
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

send_btn = st.button("🚀 Send Email(s)")

# --- 4. LOGIC (WEB VERSION OF YOUR WIN32 LOGIC) ---
if send_btn:
    if not draft_subject:
        st.error("Please enter the Draft Subject.")
    else:
        headers = {'Authorization': f"Bearer {st.session_state.token}"}
        # Use 'me' or a specific user
        base_url = f"https://graph.microsoft.com/v1.0/{f'users/{from_email}' if from_email else 'me'}"

        try:
            # Step A: Find the Draft
            draft_query = f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true"
            r = requests.get(draft_query, headers=headers).json()

            if 'value' not in r or len(r['value']) == 0:
                st.error(f"Could not find draft: '{draft_subject}'")
            else:
                body_content = r['value'][0]['body']['content']
                
                # Step B: Excel Recipient Logic
                bcc_list = []
                if uploaded_file:
                    df = pd.read_excel(uploaded_file, header=None)
                    all_rows = df.iloc[:, 0].dropna().astype(str).tolist()
                    
                    if all_rows and "@" not in all_rows[0]:
                        st.write(f"ℹ️ Skipping header row: '{all_rows[0]}'")
                        bcc_list = all_rows[1:]
                    else:
                        bcc_list = all_rows

                # Step C: Sending Logic
                if not to_email and not bcc_list:
                    st.error("No recipients found.")
                else:
                    total_batches = (len(bcc_list) + batch_size - 1) // batch_size if bcc_list else 1
                    
                    # Batch Loop
                    for i in range(0, max(len(bcc_list), 1), int(batch_size)):
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
                        
                        send_res = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                        
                        if send_res.status_code == 202:
                            if bcc_list:
                                st.write(f"✅ Sent Batch {batch_num} of {total_batches}")
                            else:
                                st.success(f"✅ Single email sent successfully to {to_email}")
                        else:
                            st.error(f"Error sending: {send_res.text}")

                        # Countdown Timer
                        if batch_num < total_batches:
                            countdown_placeholder = st.empty()
                            for seconds_left in range(5, 0, -1):
                                countdown_placeholder.info(f"⏳ Waiting {seconds_left} seconds before next batch...")
                                time.sleep(1)
                            countdown_placeholder.empty()

                    if bcc_list:
                        st.success(f"🎉 All batches sent successfully!")

        except Exception as e:
            st.error(f"An error occurred: {e}")