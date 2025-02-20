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
    #mail_details = "\n".join([f"Subject: {mail['subject']}\nFrom: {mail['from']['emailAddress']['address']}\nBody: {mail['bodyPreview']}" for mail in mails])
    
    h = html2text.HTML2Text()  
    h.ignore_links = True  

    mail_details = "\n".join([  
        f"Subject: {mail['subject']}\n"  
        f"From: {mail['from']['emailAddress']['address']}\n"  
        f"Body: {h.handle(mail['body']['content']) if mail['body']['contentType'] == 'html' else mail['body']['content']}"  
        for mail in mails  
    ])

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

# def query_responder(query, mails):
#     """Respond to user query using the mail details."""
#     h = html2text.HTML2Text()  
#     h.ignore_links = True  

#     mail_details = "\n".join([  
#         f"Subject: {mail['subject']}\n"  
#         f"From: {mail['from']['emailAddress']['address']}\n"  
#         f"Body: {h.handle(mail['body']['content']) if mail['body']['contentType'] == 'html' else mail['body']['content']}"  
#         for mail in mails  
#     ])

#     prompt = f"Respond to the user's query based on the following email details:\n{mail_details}\n\nUser's Query: {query}"
    
#     response = client.chat.completions.create(
#         model="gpt-4o",
#         messages=[
#             {"role": "system", "content": "You are a helpful assistant."},
#             {"role": "user", "content": prompt}
#         ],
#         temperature=0.5,
#     )
#     return response.choices[0].message.content.strip()

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
    """Fetch all emails from the user's inbox using pagination."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$select=subject,from,bodyPreview,body"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    all_mails = []
    
    while url:  
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            all_mails.extend(data.get("value", []))  
            url = data.get("@odata.nextLink")  
        else:
            st.error(f"Error fetching emails: {response.status_code} - {response.text}")
            break  

    return all_mails

# Streamlit UI
st.title("docQuest Mailer")

# Email input
user_email = st.text_input("Enter User Email")  

# Store emails in session state
if "mails" not in st.session_state:
    st.session_state.mails = []

if st.button("Fetch mails"):
    token = get_access_token()  
    if token and user_email:  
        mails = fetch_emails(token, user_email)
        if mails:
            st.session_state.mails = mails  # Store fetched mails in session state
            st.success(f"Found {len(mails)} email(s).")
        else:
            st.warning("No emails found.")
    else:
        st.error("Enter a valid email address or check authentication.")

# **Query input always visible**  
user_query = st.text_input("Ask a question about the emails")  

if st.button("Ask"):
    if not st.session_state.mails:
        st.error("No emails available. Fetch emails first.")
    elif not user_query:
        st.error("Enter a query to ask.")
    else:
        st.write("Answering the query based on the emails...")
        answer = query_responder(user_query, st.session_state.mails)
        st.write(f"**Answer:** {answer}")
