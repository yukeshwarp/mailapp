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
    h = html2text.HTML2Text()  
    h.ignore_links = True  
    mails_s = mails[0:40]
    mail_details = "\n".join([  # Combine relevant email details
        f"Subject: {mail.get('subject', 'No Subject')}\n"
        f"From: {mail.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}\n"
        f"Received: {mail.get('receivedDateTime', 'Unknown Time')}\n"  # Include received time
        f"Body: {h.handle(mail['body']['content']) if mail.get('body', {}).get('contentType') == 'html' else mail.get('body', {}).get('content', 'No Content')}"
        for mail in mails_s
    ])

    prompt = f"Respond to the user's query based on the following email details:\n{mail_details}\n\nUser's Query: {query}"
    
    response = client.chat.completions.create(
        model="gpt-4o",  # Specify GPT model
        messages=[  # Provide prompt to the model
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ],
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
    """Fetch all emails from the user's inbox using pagination."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$select=subject,from,bodyPreview,body,receivedDateTime"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    all_mails = []
    
    while url:  # Loop through paginated results
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            all_mails.extend(data.get("value", []))  # Append emails to the list
            url = data.get("@odata.nextLink")  # Get next page URL
        else:
            st.error(f"Error fetching emails: {response.status_code} - {response.text}")
            break  # Stop if there's an error

    return all_mails

def convert_emails_to_json(mails):
    """Convert email data to JSON format after processing HTML content."""
    h = html2text.HTML2Text()  
    h.ignore_links = True  

    for mail in mails:
        # Convert HTML body to plain text
        if mail.get('body', {}).get('contentType') == 'html':
            mail['body']['content'] = h.handle(mail['body']['content'])
    
    # Convert the processed mails list to JSON
    return json.dumps(mails, indent=4)

# Streamlit app UI
st.title("Outlook Mail Viewer with QA")  

user_email = st.text_input("Enter User Email")  
user_query = st.text_input("Ask a question about the emails")  # New input field for queries

if st.button("Ask"):  
    token = get_access_token()  
    if token:  
        if user_email:  
            mails = fetch_emails(token, user_email)  
            st.write(f"Found {len(mails)} email(s)")

            # Check if a user has asked a question
            if user_query:
                st.write("Answering the query based on the emails...")

                # Use the query_responder function to generate the answer
                answer = query_responder(user_query, mails)
                st.write(f"Answer: {answer}")
                
            # Convert mails list to JSON after processing HTML content
            mails_json = convert_emails_to_json(mails)

            # Display the JSON content in the app
            st.json(mails_json)

            # Create a download button to download the JSON file
            st.download_button(
                label="Download Emails as JSON", 
                data=mails_json, 
                file_name="emails.json", 
                mime="application/json"
            )
                
        else:  
            st.error("Please enter a valid email address.") 
    else:
        st.error("Error acquiring access token.")
