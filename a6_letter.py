import os.path

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError

from reportlab.lib.pagesizes import A6
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

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


def create_pdf(recipients, pdf_filename="letters.pdf"):
    """
    Creates a PDF with A6 pages, each containing name, address, and city_state_zip
    from the 'data' array, centered on the page.  Includes a return address in the
    upper left if provided.

    Args:
        data: A list of dictionaries, where each dictionary contains
              'name', 'address', and 'city_state_zip' keys.
        output_filename: The name of the PDF file to be created.
        return_address: A string containing the return address, or None.
    """
    A6_Landscape_Oversize = (188*mm, 105*mm)
    A6_Landscape = (148*mm, 105*mm)
    doc = SimpleDocTemplate(
        pdf_filename, pagesize=A6_Landscape_Oversize, showBoundary=0)
    centered_style = ParagraphStyle(name="CenterNormal",
                                    fontSize=17, # Font Size
                                    leading=12, # Line height
                                    leftIndent=25,
                                    alignment=TA_LEFT)

    top_spacer = Spacer(1, 0.5 * inch)
    spacer = Spacer(1, 0.05 * inch)

    story = []

    for i, item in enumerate(recipients):
        if i > 0:  # Add page break *before* all but the first item
            story.append(PageBreak())

        name = item.get("name", "")
        address = item.get("address", "")
        city_state_zip = item.get("city_state_zip", "")

        name_paragraph = Paragraph(name, centered_style)
        address_paragraph = Paragraph(address, centered_style)
        city_state_zip_paragraph = Paragraph(city_state_zip, centered_style)

        story.append(top_spacer)
        story.append(name_paragraph)
        story.append(spacer)
        story.append(address_paragraph)
        story.append(spacer)
        story.append(city_state_zip_paragraph)


    doc.build(story)


if __name__ == '__main__':
    #from_information = {
    #    'name': 'David & Akhokem',
    #    'address': '1214 Falconer Way',
    #    'city_state_zip': 'Austin, Tx, 78748'
    #}
    from_information = None
    spreadsheet_data = get_spreadsheet_data()
    if spreadsheet_data:
        create_pdf(spreadsheet_data)
