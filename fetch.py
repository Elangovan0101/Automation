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
        # Iterate through parts to find plain text or HTML
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
        # Handle non-multipart email
        try:
            return msg.get_payload(decode=True).decode()
        except Exception as e:
            print(f"Error decoding non-multipart email: {e}")
            return ""  # Return empty if decoding fails
    return ""  # Fallback for any unhandled cases

def extract_key_info(body):
    """Extracts key details from the email body."""
    # Initialize variables
    customer_name = "Not Found"
    order_id = "Not Found"
    feedback_category = "General"
    sentiment = "Neutral"  # Initialize sentiment description

    # Convert the body to lower case for easier matching
    lower_body = body.lower()

    # Extract customer name
    if "my name is" in lower_body:
        try:
            customer_name = lower_body.split("my name is ")[1].split()[0].strip()  # Get the name after 'my name is'
        except IndexError:
            customer_name = "Not Found"
    elif "i am" in lower_body:
        try:
            customer_name = lower_body.split("i am ")[1].split()[0].strip()  # Get the name after 'I am'
        except IndexError:
            customer_name = "Not Found"

    # Extract order ID
    if "order id is" in lower_body:
        try:
            order_id = lower_body.split("order id is ")[1].split()[0].strip()  # Get the order ID after 'order id is'
        except IndexError:
            order_id = "Not Found"

    # Determine feedback category and sentiment description
    positive_keywords = ["satisfied", "happy", "great", "excellent", "love", "good", "awesome", "wonderful"]
    negative_keywords = ["disappointed", "broke", "bad", "hate", "terrible", "awful", "poor", "not happy", "issue"]

    # Check for positive sentiments
    if any(keyword in lower_body for keyword in positive_keywords):
        feedback_category = "Product"
        sentiment = "Positive"  # Use descriptive sentiment
    # Check for negative sentiments
    elif any(keyword in lower_body for keyword in negative_keywords):
        feedback_category = "Product"
        sentiment = "Negative"  # Use descriptive sentiment
    else:
        sentiment = "Neutral"  # Default to Neutral if no keywords found

    # Check for specific feedback scenarios
    if "payment" in lower_body:
        feedback_category = "Payment Issue"
    elif "shipping" in lower_body or "delivery" in lower_body:
        feedback_category = "Shipping Issue"
    elif "product" in lower_body or "quality" in lower_body or "broke" in lower_body:
        feedback_category = "Product Quality"
    elif "disappointed" in lower_body:
        feedback_category = "Product Issue"
        sentiment = "Negative"
    elif "satisfied" in lower_body or "happy" in lower_body:
        feedback_category = "Product Feedback"
        sentiment = "Positive"
    elif "return" in lower_body or "exchange" in lower_body:
        feedback_category = "Return/Exchange"
    elif "complaint" in lower_body or "issue" in lower_body:
        feedback_category = "General Complaint"
    elif "suggestion" in lower_body or "recommend" in lower_body:
        feedback_category = "Suggestion"

    return {
        "Customer Name": customer_name,
        "Order ID": order_id,
        "Feedback Category": feedback_category,
        "Sentiment": sentiment  # Change sentiment score to descriptive
    }

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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)

        # Query to fetch emails containing feedback/complaint-related keywords
        query = "feedback OR complaint OR replacement OR issue OR problem OR concern OR disappointed OR satisfied OR experience"
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
                        continue  # Skip and try the next part
            elif 'body' in payload and 'data' in payload['body']:
                try:
                    email_msg = email.message_from_bytes(base64.urlsafe_b64decode(payload['body']['data'].encode('UTF-8')))
                    body = get_email_body(email_msg)
                except Exception as e:
                    print(f"Error decoding body: {e}")

            # Extract key information from the email body
            extracted_info = extract_key_info(body)

            # Exclude messages that do not have required fields
            if extracted_info["Customer Name"] != "Not Found" and extracted_info["Order ID"] != "Not Found":
                print(f"Message ID: {msg['id']}\nSubject: {subject}\nExtracted Info: {extracted_info}\n")

    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
