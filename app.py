import streamlit as st
import msal
import requests
import json
import os
from openai import AzureOpenAI
import html2text

# Azure app registration details
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]  # Permission to read mail correct scope

# LLM setup
client = AzureOpenAI(
    azure_endpoint=os.getenv("LLM_ENDPOINT"),
    api_key=os.getenv("LLM_KEY"),
    api_version="2024-10-01-preview",
)

def query_responder(query, mails):
    """Respond to user query using the mail details."""
    # Create a prompt to ask the LLM for a response based on the emails.
    mail_details = "\n".join([f"Subject: {mail['subject']}\nFrom: {mail['from']['emailAddress']['address']}\nBody: {mail['bodyPreview']}" for mail in mails])
    prompt = f"Respond to the user's query based on the following email details:\n{mail_details}\n\nUser's Query: {query}"

    # Use LLM to respond to the user query based on mail details.
    response = client.chat.completions.create(
        model="gpt-4o",  # Replace with your model ID
        messages=[{
            "role": "system",
            "content": "You are a helpful assistant.",
        },
        {
            "role": "user",
            "content": prompt
        }],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

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
  
def fetch_emails(access_token, user_email):  
    """Fetch emails from the user's inbox."""  
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$select=subject,from,body"  
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
            for mail in mails:  
                #with st.expander(mail["subject"]):  
                # print(mail)
                # st.write(f"**From:** {mail['from']['emailAddress']['address']}")  
                # #st.write(f"**Received:** {mail['receivedDateTime']}")  
                # st.write(f"**Body:** {mail.get('body', 'No preview available')}")  
                # st.write("---")
                h = html2text.HTML2Text()  
                h.ignore_links = True  

                mail_details = "\n".join([  
                    f"Subject: {mail['subject']}\n"  
                    f"From: {mail['from']['emailAddress']['address']}\n"  
                    f"Body: {h.handle(mail['body']['content']) if mail['body']['contentType'] == 'html' else mail['body']['content']}"  
                    for mail in mails  
                ]) 
                st.write(mail_details)
        else:  
            st.error("Error reading mail") 
