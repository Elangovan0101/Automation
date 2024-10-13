import os
import base64
import email
import re
import time
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from textblob import TextBlob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Define the Google Form link
form_link = "https://forms.gle/gFX7TmZBBcrW69Up8"

# Google API scopes
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
    
    lower_body = body.lower()

    # Extract customer name and order ID
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

    # Use TextBlob to analyze sentiment
    analysis = TextBlob(body)
    sentiment = "Positive" if analysis.sentiment.polarity > 0 else "Negative" if analysis.sentiment.polarity < 0 else "Neutral"

    if customer_name == "Not Found" or order_id == "Not Specified":
        return None

    return {
        "Customer Name": customer_name,
        "Order ID": order_id,
        "Feedback Category": feedback_category,
        "Sentiment": sentiment
    }

def submit_to_google_form(entry):
    """Submits extracted data to Google Form using Selenium."""
    options = webdriver.ChromeOptions()
    options.add_argument("-incognito")
    browser = webdriver.Chrome(options=options)

    try:
        browser.get(form_link)

        # Customer Name
        customer_name_field = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i1']"))
        )
        customer_name_field.clear()
        customer_name_field.send_keys(entry["Customer Name"])

        # Order ID
        order_id_field = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i5']"))
        )
        order_id_field.clear()
        order_id_field.send_keys(entry["Order ID"])

        # Feedback Category
        feedback_category_field = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i9']"))
        )
        feedback_category_field.clear()
        feedback_category_field.send_keys(entry["Feedback Category"])

        # Sentiment
        sentiment = entry["Sentiment"]
        if sentiment == "Positive":
            sentiment_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Positive']"))
            )
        elif sentiment == "Neutral":
            sentiment_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Neutral']"))
            )
        elif sentiment == "Negative":
            sentiment_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Negative']"))
            )
        sentiment_field.click()

        # Submit
        submit_button = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']"))
        )
        submit_button.click()

    finally:
        browser.quit()

def main():
    """Fetches customer feedback and submits to Google Form."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)
        query = "shipping OR delivery OR delay OR feedback OR complaint OR replacement OR issue"
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query).execute()
        messages = results.get('messages', [])

        if not messages:
            print('No relevant messages found.')
            return

        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()

            # Extract body content
            body = ""
            payload = msg.get('payload', {})
            if 'parts' in payload:
                for part in payload['parts']:
                    if 'body' in part and 'data' in part['body']:
                        email_msg = email.message_from_bytes(base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8')))
                        body = get_email_body(email_msg)
                        break

            elif 'body' in payload and 'data' in payload['body']:
                email_msg = email.message_from_bytes(base64.urlsafe_b64decode(payload['body']['data'].encode('UTF-8')))
                body = get_email_body(email_msg)

            # Extract key information
            extracted_info = extract_key_info(body)
            if extracted_info:
                print(f"Submitting: {extracted_info}")
                submit_to_google_form(extracted_info)

    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
