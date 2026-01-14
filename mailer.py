import streamlit as st
import msal
import requests
import pandas as pd
import time
import os

# --- 1. CONFIGURATION ---
CLIENT_ID = st.secrets["MS_CLIENT_ID"]
CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://cloud-outlook-sender-kn4vdkgrcmxz7pfk5lfp3f.streamlit.app/" 
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]

TOKEN_FILE = "session_token.txt"

# --- 2. LOGIN HANDLER ---
if "code" in st.query_params:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(
        st.query_params["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        with open(TOKEN_FILE, "w") as f:
            f.write(result["access_token"])
        st.query_params.clear()
        st.rerun()

# --- 3. MAIN APP UI ---
st.set_page_config(page_title="Outlook Universal Sender", layout="wide")
st.title("📧 Outlook Universal Sender")

with st.sidebar:
    st.header("1. Account Settings")
    from_email = st.text_input("Send From (Account Email)", placeholder="Default Account")
    batch_size = st.number_input("BCC Batch Size", value=50, min_value=1)
    
    if st.button("🔌 Force Logout"):
        if os.path.exists(TOKEN_FILE): os.remove(TOKEN_FILE)
        if 'token' in st.session_state: del st.session_state.token
        st.rerun()

# --- 4. THE SYNC GATE ---
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE, "r") as f:
        st.session_state.token = f.read()

if 'token' not in st.session_state:
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI, prompt="select_account")
    
    st.markdown(f"""
        <div style="text-align: center; margin: 30px 0;">
            <a href="{auth_url}" target="_blank" style="
                background-color: #0078d4; color: white; padding: 20px 50px; 
                text-decoration: none; border-radius: 10px; font-weight: bold; 
                display: inline-block; font-size: 20px;
            ">1. LOGIN (NEW TAB)</a>
        </div>
    """, unsafe_allow_html=True)
    if st.button("2. ✅ ACTIVATE SENDER", use_container_width=True, type="primary"):
        st.rerun()

else:
    st.success("🟢 Account Linked")
    st.subheader("2. Draft & Recipients")
    draft_subject = st.text_input("Draft Email Subject")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("To (Optional)")
        cc_email = st.text_input("CC (Optional)")
    with col2:
        uploaded_file = st.file_uploader("Upload Excel (Optional)", type=["xlsx"])

    if st.button("🚀 START EMAIL BLAST", use_container_width=True, type="primary"):
        if not draft_subject:
            st.error("Please enter the Draft Subject.")
        else:
            headers = {'Authorization': f"Bearer {st.session_state.token}"}
            user_path = f"users/{from_email}" if from_email else "me"
            base_url = f"https://graph.microsoft.com/v1.0/{user_path}"
            
            try:
                # 1. FIND DRAFT
                draft_res = requests.get(f"{base_url}/messages?$filter=subject eq '{draft_subject}' and isDraft eq true", headers=headers).json()
                
                if 'value' not in draft_res or len(draft_res['value']) == 0:
                    st.error(f"❌ Draft '{draft_subject}' not found.")
                else:
                    body_content = draft_res['value'][0]['body']['content']
                    
                    # 2. EXTRACT BCC FROM EXCEL
                    final_bcc_list = []
                    if uploaded_file:
                        df = pd.read_excel(uploaded_file, header=None)
                        raw_list = df.iloc[:, 0].dropna().astype(str).tolist()
                        if raw_list and "@" not in raw_list[0]:
                            final_bcc_list = raw_list[1:]
                        else:
                            final_bcc_list = raw_list

                    # 3. VERIFY ANY RECIPIENTS
                    if not final_bcc_list and not to_email and not cc_email:
                        st.error("❌ No recipients found.")
                    else:
                        if final_bcc_list:
                            total_batches = (len(final_bcc_list) + int(batch_size) - 1) // int(batch_size)
                            
                            for i in range(0, len(final_bcc_list), int(batch_size)):
                                batch_num = (i // int(batch_size)) + 1
                                current_batch = final_bcc_list[i : i + int(batch_size)]
                                
                                # Payload includes To and CC for EVERY batch iteration
                                payload = {
                                    "message": {
                                        "subject": draft_subject,
                                        "body": {"contentType": "HTML", "content": body_content},
                                        "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                        "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else [],
                                        "bccRecipients": [{"emailAddress": {"address": e}} for e in current_batch]
                                    }
                                }
                                
                                r = requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                                if r.status_code == 202:
                                    st.write(f"✅ Sent Batch {batch_num} of {total_batches} (Rows {i+1} to {i+len(current_batch)})")
                                else:
                                    st.error(f"Error: {r.text}")

                                if batch_num < total_batches:
                                    countdown_placeholder = st.empty()
                                    for s in range(5, 0, -1):
                                        countdown_placeholder.info(f"⏳ Waiting {s} seconds...")
                                        time.sleep(1)
                                    countdown_placeholder.empty()
                            st.success(f"🎉 Process Complete! Sent to {len(final_bcc_list)} BCC recipients.")
                        
                        else:
                            # Single send logic for just To/CC
                            payload = {
                                "message": {
                                    "subject": draft_subject,
                                    "body": {"contentType": "HTML", "content": body_content},
                                    "toRecipients": [{"emailAddress": {"address": to_email}}] if to_email else [],
                                    "ccRecipients": [{"emailAddress": {"address": cc_email}}] if cc_email else []
                                }
                            }
                            requests.post(f"{base_url}/sendMail", headers=headers, json=payload)
                            st.success(f"✅ Email sent successfully.")
            
            except Exception as e:
                st.error(f"Error: {e}")