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
SCOPE = ["https://graph.microsoft.com/.default"]

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
    mails_s = mails[:30]  # Limit to 30 emails to avoid long context

    mail_details = "\n".join([
        f"Subject: {mail.get('subject', 'No Subject')}\n"
        f"From: {mail.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}\n"
        f"Received: {mail.get('receivedDateTime', 'Unknown Time')}\n"
        f"Importance: {mail.get('importance', 'Normal')}\n"
        f"Priority: {mail.get('priority', 'Normal')}\n"
        f"Has Attachment: {mail.get('hasAttachments', False)}\n"
        f"Categories: {', '.join(mail.get('categories', [])) if mail.get('categories') else 'None'}\n"
        f"Conversation ID: {mail.get('conversationId', 'N/A')}\n"
        f"Weblink: {mail.get('webLink', 'No Link')}\n"
        f"Body: {h.handle(mail['body']['content']) if mail.get('body', {}).get('contentType') == 'html' else mail.get('body', {}).get('content', 'No Content')}"
        for mail in mails_s
    ])

    prompt = f"Respond to the user's query based on the following email details:\n{mail_details}\n\nUser's Query: {query}"
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

def get_access_token():  
    """Authenticate and get access token."""
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
    """Fetch emails with additional metadata from Outlook."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?" \
          f"$select=subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview,body,receivedDateTime,sentDateTime," \
          f"importance,priority,hasAttachments,categories,conversationId,conversationIndex,isRead,isDraft,webLink"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    all_mails = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            all_mails.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # Get next page URL
        else:
            st.error(f"Error fetching emails: {response.status_code} - {response.text}")
            break

    return all_mails

def format_recipients(recipients):
    """Format recipients list as a readable string."""
    if not recipients:
        return "None"
    return ", ".join([recipient.get("emailAddress", {}).get("address", "Unknown") for recipient in recipients])

def convert_emails_to_json(mails):
    """Convert email data to JSON format after processing HTML content."""
    h = html2text.HTML2Text()
    h.ignore_links = True  

    processed_mails = []
    for mail in mails:
        processed_mails.append({
            "subject": mail.get("subject", "No Subject"),
            "from": mail.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender"),
            "to": format_recipients(mail.get("toRecipients", [])),
            "cc": format_recipients(mail.get("ccRecipients", [])),
            "bcc": format_recipients(mail.get("bccRecipients", [])),
            "receivedDateTime": mail.get("receivedDateTime", "Unknown Time"),
            "sentDateTime": mail.get("sentDateTime", "Unknown Time"),
            "importance": mail.get("importance", "Normal"),
            "priority": mail.get("priority", "Normal"),
            "hasAttachments": mail.get("hasAttachments", False),
            "categories": mail.get("categories", []),
            "conversationId": mail.get("conversationId", "N/A"),
            "conversationIndex": mail.get("conversationIndex", "N/A"),
            "isRead": mail.get("isRead", False),
            "isDraft": mail.get("isDraft", False),
            "webLink": mail.get("webLink", "No Link"),
            "body": h.handle(mail["body"]["content"]) if mail.get("body", {}).get("contentType") == "html" else mail.get("body", {}).get("content", "No Content")
        })
    
    return json.dumps(processed_mails, indent=4)

# Streamlit UI
st.title("Outlook Mail Viewer with QA")  

user_email = st.text_input("Enter User Email")  
user_query = st.text_input("Ask a question about the emails")

if st.button("Ask"):
    token = get_access_token()
    if token:
        if user_email:
            mails = fetch_emails(token, user_email)
            st.write(f"Found {len(mails)} email(s)")

            if user_query:
                st.write("Answering the query based on the emails...")
                answer = query_responder(user_query, mails)
                st.write(f"Answer: {answer}")

            # Convert to JSON
            mails_json = convert_emails_to_json(mails)
            st.json(mails_json)

            # JSON Download Button
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
