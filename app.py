import streamlit as st
import msal
import requests
import json
import os
from openai import AzureOpenAI

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
    """Get access token using MSAL."""
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
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        st.error(f"Error fetching emails: {response.status_code} - {response.text}")
        return []

st.title("Outlook Mail Viewer and LLM Assistant")

# Chat history
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# User input (email address)
user_email = "Yukeshwar@docu3c.com"

# Chat interaction
user_input = st.chat_input("Ask me something about your email")

if user_input:
    # Add user message to chat history
    st.session_state.chat_history.append({"role": "user", "content": user_input})

    # Get access token
    token = get_access_token()
    if token:
        if user_email:
            mails = fetch_emails(token, user_email)
            if mails:
                # Send the emails to the LLM for generating the response
                assistant_response = query_responder(user_input, mails)
                
                # Add assistant's response to chat history
                st.session_state.chat_history.append({
                    "role": "assistant",
                    "content": f"Here's the response to your query:\n{assistant_response}"
                })
            else:
                st.session_state.chat_history.append({
                    "role": "assistant",
                    "content": "I couldn't find any emails in your inbox. Try again later."
                })
        else:
            st.session_state.chat_history.append({
                "role": "assistant",
                "content": "Please enter a valid user email."
            })

    # Display chat history
    for message in st.session_state.chat_history:
        st.chat_message(message["role"]).markdown(message["content"])

    # Generate the response using LLM for further user queries
    if user_input.lower() in ["show me more", "tell me more", "get more details"]:
        st.session_state.chat_history.append({"role": "user", "content": "Show me more emails."})
        # Fetch the next set of emails or details, then display
        # You can add logic to fetch more emails or continue the conversation
        st.experimental_rerun()

