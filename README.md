# OneNote Notifier

A Python application that monitors OneNote updates in shared workspaces, generates AI-powered summaries of modified notes, and sends email notifications.

## Overview

This utility monitors shared OneNote notebooks for changes, uses AI to create concise summaries of modified content, and automatically sends these summaries via email. Perfect for teams wanting to stay updated on shared notebook changes without manual checking.

## Features

- Real-time monitoring of shared OneNote notebooks using Microsoft Graph API
- AI-powered summarization of modified notes using OpenAI's API
- Automated email notifications via Gmail API
- Support for multiple notebooks and sections
- Configurable monitoring intervals
- Secure authentication handling

## Prerequisites

- Python 3.8 or higher
- Microsoft 365 account with access to shared OneNote notebooks
- Google account for sending emails
- OpenAI API key for AI summarization

## Setup Instructions

1. **Microsoft Graph API Setup**
   - Register your application in the Azure Portal
   - Grant necessary permissions for OneNote access
   - Note down your client ID and secret

2. **Gmail API Setup**
   - Enable Gmail API in Google Cloud Console
   - Configure OAuth 2.0 credentials
   - Download credentials file

3. **OpenAI API Setup**
   - Create an OpenAI account
   - Generate API key
   - Save the key securely

4. **Python Environment Setup**
   ```bash
   # Create and activate virtual environment
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   
   # Install required packages
   pip install msal openai google-api-python-client google-auth-httplib2 google-auth-oauthlib
   ```

## Implementation Guide

### 1. Configure Microsoft Graph API Authentication

1. Create a configuration file `config.py`:
   ```python
   MICROSOFT_CLIENT_ID = "your_client_id"
   MICROSOFT_CLIENT_SECRET = "your_client_secret"
   MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common"
   MICROSOFT_SCOPE = ["Notes.Read.All", "Notes.ReadWrite.All"]
   ```

2. Set up Microsoft Graph authentication:
   ```python
   import msal
   
   def get_graph_client():
       app = msal.ConfidentialClientApplication(
           MICROSOFT_CLIENT_ID,
           authority=MICROSOFT_AUTHORITY,
           client_credential=MICROSOFT_CLIENT_SECRET
       )
       result = app.acquire_token_silent(MICROSOFT_SCOPE, account=None)
       if not result:
           result = app.acquire_token_for_client(scopes=MICROSOFT_SCOPE)
       return result.get('access_token')
   ```

### 2. Implement OneNote Monitoring

1. Create a monitoring script `monitor.py`:
   ```python
   import requests
   import time
   from datetime import datetime, timedelta
   
   def setup_subscription():
       token = get_graph_client()
       headers = {'Authorization': f'Bearer {token}'}
       subscription = {
           'changeType': 'updated',
           'notificationUrl': 'your_webhook_url',
           'resource': '/me/onenote/notebooks',
           'expirationDateTime': (datetime.utcnow() + timedelta(days=2)).isoformat() + 'Z',
           'clientState': 'your_client_state'
       }
       response = requests.post(
           'https://graph.microsoft.com/v1.0/subscriptions',
           headers=headers,
           json=subscription
       )
       return response.json()
   
   def get_notebook_changes():
       token = get_graph_client()
       headers = {'Authorization': f'Bearer {token}'}
       response = requests.get(
           'https://graph.microsoft.com/v1.0/me/onenote/pages',
           headers=headers
       )
       return response.json().get('value', [])
   ```

### 3. Implement AI Summarization

1. Create a summarization script `summarize.py`:
   ```python
   import openai
   from config import OPENAI_API_KEY
   
   openai.api_key = OPENAI_API_KEY
   
   def summarize_content(content):
       response = openai.ChatCompletion.create(
           model="gpt-4",
           messages=[
               {"role": "system", "content": "Summarize the following OneNote content concisely:"},
               {"role": "user", "content": content}
           ],
           max_tokens=150
       )
       return response.choices[0].message['content']
   
   def process_page_changes(pages):
       summaries = []
       for page in pages:
           content = get_page_content(page['id'])  # Implement this function
           summary = summarize_content(content)
           summaries.append({
               'title': page['title'],
               'summary': summary,
               'lastModified': page['lastModifiedDateTime']
           })
       return summaries
   ```

### 4. Implement Email Notifications

1. Create an email notification script `notify.py`:
   ```python
   from google.oauth2.credentials import Credentials
   from google_auth_oauthlib.flow import InstalledAppFlow
   from google.auth.transport.requests import Request
   from googleapiclient.discovery import build
   import base64
   from email.mime.text import MIMEText
   import os
   
   SCOPES = ['https://www.googleapis.com/auth/gmail.send']
   
   def get_gmail_service():
       creds = None
       if os.path.exists('token.json'):
           creds = Credentials.from_authorized_user_file('token.json', SCOPES)
       if not creds or not creds.valid:
           if creds and creds.expired and creds.refresh_token:
               creds.refresh(Request())
           else:
               flow = InstalledAppFlow.from_client_secrets_file(
                   'credentials.json', SCOPES)
               creds = flow.run_local_server(port=0)
           with open('token.json', 'w') as token:
               token.write(creds.to_json())
       return build('gmail', 'v1', credentials=creds)
   
   def send_email_notification(summaries, to_email):
       service = get_gmail_service()
       
       email_content = "OneNote Updates Summary:\n\n"
       for summary in summaries:
           email_content += f"Page: {summary['title']}\n"
           email_content += f"Last Modified: {summary['lastModified']}\n"
           email_content += f"Summary: {summary['summary']}\n\n"
       
       message = MIMEText(email_content)
       message['to'] = to_email
       message['subject'] = 'OneNote Updates Summary'
       
       raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
       service.users().messages().send(userId='me', body={'raw': raw}).execute()
   ```

### 5. Main Application

1. Create the main script `app.py`:
   ```python
   import time
   from monitor import get_notebook_changes
   from summarize import process_page_changes
   from notify import send_email_notification
   
   def main():
       while True:
           try:
               # Get changes from OneNote
               changed_pages = get_notebook_changes()
               
               if changed_pages:
                   # Generate summaries
                   summaries = process_page_changes(changed_pages)
                   
                   # Send email notification
                   send_email_notification(summaries, 'your_email@gmail.com')
               
               # Wait for next check
               time.sleep(300)  # Check every 5 minutes
           except Exception as e:
               print(f"Error: {e}")
               time.sleep(60)  # Wait 1 minute before retrying
   
   if __name__ == "__main__":
       main()
   ```

## Running the Application

1. Set up your environment variables:
   ```bash
   export MICROSOFT_CLIENT_ID="your_client_id"
   export MICROSOFT_CLIENT_SECRET="your_client_secret"
   export OPENAI_API_KEY="your_openai_api_key"
   ```

2. Run the application:
   ```bash
   python app.py
   ```

## Security Considerations

- Store sensitive credentials in environment variables or a secure vault
- Use HTTPS for all API communications
- Implement proper error handling and logging
- Regularly rotate API keys and tokens
- Monitor API usage and implement rate limiting

## Troubleshooting

Common issues and solutions:

1. Authentication Errors
   - Verify API credentials are correct
   - Check if tokens have expired
   - Ensure proper permissions are granted

2. Rate Limiting
   - Implement exponential backoff
   - Monitor API usage
   - Adjust polling frequency

3. Missing Updates
   - Verify webhook configuration
   - Check network connectivity
   - Review API permissions
