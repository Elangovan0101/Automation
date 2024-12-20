import os
import base64
import email
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from textblob import TextBlob

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_email_body(msg):
    """Extracts the full body of the email."""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            try:
                if content_type == "text/plain":
                    return part.get_payload(decode=True).decode()  # Return plain text body
                elif content_type == "text/html":
                    return part.get_payload(decode=True).decode()  # Return HTML body as fallback
            except Exception as e:
                print(f"Error decoding part: {e}")
                continue  # Skip part on failure
    else:
        try:
            return msg.get_payload(decode=True).decode()
        except Exception as e:
            print(f"Error decoding non-multipart email: {e}")
            return ""  # Return empty if decoding fails
    return ""  # Fallback for any unhandled cases

def extract_key_info(body):
    """Extracts key details from the email body."""
    customer_name = "Not Found"
    order_id = "Not Specified"
    feedback_category = "General"
    sentiment = "Neutral"
    
    # Convert body to lowercase for easier keyword matching
    lower_body = body.lower()

    # Use regular expressions to extract customer name and order ID
    name_match = re.search(r"(my name is|i am|this is|i’m)\s+([a-zA-Z]+)", lower_body)
    if name_match:
        customer_name = name_match.group(2)

    order_id_match = re.search(r"(order id is|order number is|my order id is|order id)\s+([a-zA-Z0-9\-]+)", lower_body)
    if order_id_match:
        order_id = order_id_match.group(2)

    # Determine feedback category based on keywords
    if "damaged" in lower_body and "return" in lower_body:
        feedback_category = "Return/Exchange - Damaged"
    elif "payment" in lower_body:
        feedback_category = "Payment Issue"
    elif "shipping" in lower_body or "delivery" in lower_body:
        feedback_category = "Shipping Issue"
    elif "product" in lower_body or "quality" in lower_body or "broke" in lower_body:
        feedback_category = "Product Quality"
    elif "disappointed" in lower_body:
        feedback_category = "Product Issue"
    elif "return" in lower_body or "exchange" in lower_body:
        feedback_category = "Return/Exchange"
    elif "complaint" in lower_body or "issue" in lower_body:
        feedback_category = "General Complaint"
    elif "suggestion" in lower_body or "recommend" in lower_body:
        feedback_category = "Suggestion"

    # Use TextBlob to determine the sentiment of the message
    analysis = TextBlob(body)
    sentiment = "Positive" if analysis.sentiment.polarity > 0 else "Negative" if analysis.sentiment.polarity < 0 else "Neutral"

    # Check if the message contains essential details
    if customer_name == "Not Found" or order_id == "Not Specified":
        return None  # Exclude messages without required information

    return {
        "Customer Name": customer_name,
        "Order ID": order_id,
        "Feedback Category": feedback_category,
        "Sentiment": sentiment
    }

def main():
    """Fetches customer feedback and complaint emails from Gmail."""
    creds = None
    if os.path.exists('token1.json'):
        creds = Credentials.from_authorized_user_file('token1.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token1.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)

        # Query to fetch emails containing feedback/complaint-related keywords
        query = "shipping OR delivery OR delay OR late OR feedback OR complaint OR replacement OR issue OR problem OR concern OR disappointed OR satisfied OR experience OR damaged"
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query).execute()
        messages = results.get('messages', [])

        if not messages:
            print('No relevant messages found.')
            return
        
        print('Relevant messages found:')
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()

            # Process the message payload
            payload = msg.get('payload', {})
            headers = payload.get('headers', [])
            subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')

            # Extract body content
            body = ""
            if 'parts' in payload:
                for part in payload['parts']:
                    try:
                        if 'body' in part and 'data' in part['body']:
                            email_msg = email.message_from_bytes(base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8')))
                            body = get_email_body(email_msg)
                            break
                    except Exception as e:
                        print(f"Error extracting part: {e}")
                        continue
            elif 'body' in payload and 'data' in payload['body']:
                try:
                    email_msg = email.message_from_bytes(base64.urlsafe_b64decode(payload['body']['data'].encode('UTF-8')))
                    body = get_email_body(email_msg)
                except Exception as e:
                    print(f"Error decoding body: {e}")

            # Extract key information from the email body
            extracted_info = extract_key_info(body)

            # Only print messages that have required fields
            if extracted_info:
                print(f"Message ID: {msg['id']}\nSubject: {subject}\nExtracted Info: {extracted_info}\n")

    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
