import os
import smtplib
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Google Sheets API setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SPREADSHEET_ID = '1zmwj4wKU6QQ_9QtDj8U1IY1Pjjck5zc-8Ga6sQBzJso'  # Update with your Google Sheet ID
RANGE_NAME = 'A2:E8'  # Adjusted range to fetch all data

def get_all_submissions():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials_new.json', SCOPES)  # Updated filename
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    # Connect to the Google Sheets API
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    
    # Fetch the data from the Google Sheet
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])
    
    print("Fetched values from Google Sheet:", values)  # Debugging line
    
    if not values:
        print('No data found in the Google Sheet.')
        return None
    
    return values  # Return all fetched values

def send_email(summaries):
    # Email configuration
    sender_email = "elangovansanjay2003@gmail.com"
    receiver_email = "sec22ad028@sairamtap.edu.in"
    password = "Elangovanelan0101"
    
    # Create the email subject and body
    subject = "New Customer Feedback Summary"
    
    # Create the email body
    body = "A new feedback has been received with the following details:\n\n"
    
    for summary in summaries:
        body += f"""
        Customer Name: {summary[0]}
        Order ID: {summary[1]}
        Feedback Category: {summary[2]}
        Sentiment: {summary[3]}
        --------------------------
        """
    
    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    # Send the email using the SMTP server
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        print("Summary email sent successfully.")
    except Exception as e:
        print(f"Error sending email: {e}")

def main():
    # Fetch all submissions from the Google Sheet
    submissions = get_all_submissions()
    
    if submissions:
        # Send email notification with all submissions
        send_email(submissions)
    else:
        print("No recent submissions to process.")

if __name__ == "__main__":
    main()
