import os
import base64
import email
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_email_body(msg):
    """Extracts the full body of the email."""
    if msg.is_multipart():
        # Iterate through parts
        for part in msg.walk():
            # Check for the plain text content
            content_type = part.get_content_type()
            if content_type == "text/plain":
                return part.get_payload(decode=True).decode()  # Return plain text body
            elif content_type == "text/html":
                # Optionally extract text from HTML emails too
                return part.get_payload(decode=True).decode()  # Return HTML body as well
    else:
        # For non-multipart emails
        return msg.get_payload(decode=True).decode()  
    return ""

def main():
    """Fetches customer feedback and complaint emails from Gmail."""
    creds = None
    # The file token.json stores the user's access and refresh tokens.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        
        # Updated refined query to fetch relevant emails
        query = "feedback OR complaint OR replacement OR issue OR problem OR concern OR disappointed OR satisfied OR experience"
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query).execute()
        messages = results.get('messages', [])

        # Handle pagination if there are more messages
        while 'nextPageToken' in results:
            results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query, pageToken=results['nextPageToken']).execute()
            messages.extend(results.get('messages', []))

        if not messages:
            print('No relevant messages found.')
        else:
            print('Relevant messages:')
            for message in messages:
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                
                # Extract and decode the email content
                if 'data' in msg['payload']['parts'][0]['body']:
                    email_msg = email.message_from_bytes(base64.urlsafe_b64decode(msg['payload']['parts'][0]['body']['data'].encode('UTF-8')))
                else:
                    continue  # Skip if no data is available

                # Get the full email body
                body = get_email_body(email_msg)
                
                print(f"Message ID: {msg['id']}\nFull Message:\n{body}\n")

    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
