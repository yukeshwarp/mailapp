import streamlit as st
import msal
import requests
import json
import os

# Azure app registration details
# Azure app registration details
CLIENT_ID = os.getenv("CLIENT_ID")  
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  
TENANT_ID = os.getenv("TENANT_ID")  
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"  
SCOPE = ["https://graph.microsoft.com/.default"] # Permission to read mail correct scope

def get_access_token():  
    app = msal.ConfidentialClientApplication(  
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET  
    )  
    result = app.acquire_token_for_client(scopes=SCOPE)  
    if "access_token" in result:  
        return result["access_token"]  
    else:  
        st.error(f"Error acquiring token: {result.get('error_description')}")  
        return None  
  
def fetch_emails(access_token, user_id):  
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"  
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}  
    response = requests.get(url, headers=headers)  
    if response.status_code == 200:  
        return response.json().get("value", [])  
    else:  
        st.error(f"Error fetching emails: {response.status_code} - {response.text}")  
        return []  
  
st.title("Outlook Mail Viewer")  
  
user_email = st.text_input("Enter User Email")  
  
if st.button("Fetch Emails"):  
    token = get_access_token()  
    if token:  
        if user_email:  
            mails = fetch_emails(token, user_email)  
            # for mail in mails:  
            #     #with st.expander(mail["subject"]):  
            #     print(mail)
            #     st.write(f"**From:** {mail['from']['emailAddress']['address']}")  
            #     st.write(f"**Received:** {mail['receivedDateTime']}")  
            #     st.write(f"**Body:** {mail.get('body', 'No preview available')}")  
            #     st.write("---")
            st.chat_input
        else:  
            st.error("Error reading mail") 
