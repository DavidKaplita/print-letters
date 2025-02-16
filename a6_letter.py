import os.path

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from reportlab.lib.pagesizes import A6
from reportlab.lib.units import inch

SERVICE_ACCOUNT_FILE = 'credentials.json'

# Replace with the actual spreadsheet ID
SPREADSHEET_ID = '1dfeeMOBRfDtSgIEF9L8ArMhytRc8kZ_T5iJdXUKHG80'

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive.file']

creds = None

try:
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            # Save the credentials for the next run

        #with open("token.json", "w") as token:
        #    token.write(creds.to_json())
except HttpError as err:
    print(f"An error occurred: {err}")
    exit


def get_spreadsheet_data():
    """Retrieves data from the specified Google Spreadsheet."""

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Get the data
        ranges = [
            "Guests!B2:B",
            "Guests!J2:J",
            "Guests!K2:K",
        ]

        result = service.spreadsheets().values().batchGet(
            spreadsheetId=SPREADSHEET_ID, ranges=ranges, majorDimension="ROWS").execute()
        names = result.get('valueRanges', [])[0]['values']
        addresses_street = result.get('valueRanges', [])[1]['values']
        addresses_city = result.get('valueRanges', [])[2]['values']

        if not names:
            print('No data found.')
            return []

        # Reformat the data into a list of dictionaries
        data = []
        for i in range(min(len(names), len(addresses_street), len(addresses_city))):
            name = names[i][0] if names and names[i] else ""
            address_street = addresses_street[i][0] if addresses_street and addresses_street[i] else ""
            address_city = addresses_city[i][0] if addresses_city and addresses_city[i] else ""

            if name == "" or address_street == "" or address_city == "" or address_city.lower() == "international":
                continue

            data.append({
                'name': name,
                'address': address_street,
                'city_state_zip': address_city
            })
        return data

    except HttpError as err:
        print(f"An error occurred: {err}")
        return []


def create_google_doc(data):
    """Creates a new Google Doc with A6 page format, adds data, and centers it."""

    try:
        service = build('docs', 'v1', credentials=creds)

        # Create a new Google Doc
        document = {'title': 'Save The Date Print'}
        document = service.documents().create(body=document).execute()
        new_document_id = document.get('documentId')
        print(f"New Google Doc created with ID: {new_document_id}")

        # Add data to the document with formatting and page breaks
        requests = []
        current_index = 1  # Initialize the index

        for item in data:
            requests.append({
                'insertTable': {
                    'rows': 1,
                    'columns': 3,
                    'location': {'index': current_index}
                }
            })
            current_index += 1

            # Insert text into the table cells
            requests.append({
                'insertText': {
                    'location': {
                        'index': current_index
                    },
                    'text': item['name']
                }
            })
            current_index += 1

            requests.append({
                'insertText': {
                    'location': {
                        'index': current_index
                    },
                    'text': item['address']
                }
            })
            current_index += 1

            requests.append({
                'insertText': {
                    'location': {
                        'index': current_index
                    },
                    'text': item['city_state_zip']
                }
            })
            current_index += 1

            """
            # Center the text in all cells of the table
            requests.append({
                'updateTableCellStyle': {
                    'tableCellStyle': {
                        'contentAlignment': 'CONTENT_ALIGNMENT_LEFT'
                    },
                    'fields': 'contentAlignment',
                    'tableRange': {
                        'startIndex': current_index-4,
                        'endIndex': current_index-1
                    }
                }
            })
            """
            # Add page break after each guest (adjust index as needed)
            requests.append({
                'insertPageBreak': {
                    # Adjust the index if you add more elements
                    'location': {'index': current_index}
                }
            })
            current_index += 1

        # Send the batch update request
        result = service.documents().batchUpdate(
            documentId=new_document_id, body={'requests': requests}).execute()

        # Share the document with davidkaplita@gmail.com with viewer permissions
        # Use the Drive API for sharing
        drive_service = build('drive', 'v3', credentials=creds)
        permission = {
            'type': 'user',
            'role': 'Editor',  # Grant 'reader' (editor) permissions
            'emailAddress': 'davidkaplita@gmail.com'
        }
        drive_service.permissions().create(
            fileId=new_document_id, body=permission).execute()

        # Set the page format to A6 (manual step still required)
        print("Data written to the new Google Doc. Please manually set the page size to A6 in the document.")
        print(
            f"\nOpen the document here: https://docs.google.com/document/d/{new_document_id}/edit\n")
        return new_document_id

    except HttpError as err:
        print(f"An error occurred: {err}")
        return None


if __name__ == '__main__':
    spreadsheet_data = get_spreadsheet_data()
    if spreadsheet_data:
        create_google_doc(spreadsheet_data)
