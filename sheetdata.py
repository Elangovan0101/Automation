import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1zmwj4wKU6QQ_9QtDj8U1IY1Pjjck5zc-8Ga6sQBzJso"
SAMPLE_RANGE_NAME = "A2:E8"  # Adjusted to include all relevant columns

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # Use the service account key file
    creds = service_account.Credentials.from_service_account_file(
        'sheet.json', scopes=SCOPES
    )

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return

        print("Timestamp, Customer Name, Order ID, Feedback Category, Sentiment:")
        for row in values:
            # Check if the row has at least five elements
            if len(row) >= 5:
                print(f"{row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}")  # Print all relevant columns
            else:
                print("Row has fewer than five columns.")
    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
