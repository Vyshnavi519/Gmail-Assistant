from __future__ import print_function
import os.path
import base64
import email
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying scopes, delete token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    """Shows basic email info: sender, subject, and snippet."""
    creds = None
    # Load existing credentials if available
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If no valid credentials available, let user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save credentials for future use
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Build Gmail API service
    service = build('gmail', 'v1', credentials=creds)

    # Get list of messages
    results = service.users().messages().list(userId='me', maxResults=5).execute()
    messages = results.get('messages', [])

    if not messages:
        print("No emails found.")
    else:
        print("Latest Emails:")
        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            headers = msg_data['payload']['headers']

            subject = None
            sender = None
            for header in headers:
                if header['name'] == 'Subject':
                    subject = header['value']
                if header['name'] == 'From':
                    sender = header['value']

            snippet = msg_data.get('snippet')
            print(f"From: {sender}")
            print(f"Subject: {subject}")
            print(f"Snippet: {snippet}")
            print("="*50)

if __name__ == '__main__':
    main()
