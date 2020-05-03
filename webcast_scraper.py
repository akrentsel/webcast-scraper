from __future__ import print_function
import pickle
import os.path
import random
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup
from os import listdir, rename
from os.path import isfile, join

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID of the Webcast spreadsheet in Google Drive (seen in the URL).
# https://docs.google.com/spreadsheets/d/<id_here>/edit
WEBCAST_SPREADSHEET_ID = '<insert_id_here>'

# The directory where the html files are found, and the folder to move them to
# (within that directory) once done processing.
HTML_FILE_DIRECTORY = "./webcast_html/"
PROCESSED_FOLDER = "processed/"

# Number of columns between starts of new class records within a sheet.
SHEET_SPACING = 3

def main():
    """
    Sets up Google Sheets API object, and gets list of all files in the
    HTML_FILE_DIRECTORY, and calls to process them one by one.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Create the sheet object that will be used for all writing and reading.
    # This object will be passed around to functions that need to write/read.
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Process each file in the webcast_html directory:
    listed_files = [f for f in listdir(HTML_FILE_DIRECTORY) if isfile(join(HTML_FILE_DIRECTORY, f))]
    for filename in listed_files:
        # We ignore .DS_Store, which is a MacOS autogenereated metadata file.
        if filename == ".DS_Store":
            continue
        print("Scraping " + filename)
        process_html_file(HTML_FILE_DIRECTORY, filename, sheet)

def column_letter(i):
    """
    Return the name for the ith column of a Google sheet, 0-indexed.
    A - Z for columns 0 through 25, then AA-AZ, BA-AZ, ..., AZ-ZZ for > 25.
    """
    if i < 26:
        return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[i]
    else:
        first_letter_index = (i // 26) - 1
        second_letter_index = i % 26
        return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[first_letter_index] + "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[second_letter_index]


def create_new_sheet(title, sheet, num_columns):
    """
    Create a new sheet in the Google Sheets spreadsheet specified by
    WEBCAST_SPREADSHEET_ID. Creates the tab with a random colors.
    """
    request_body = {
        'requests': [{
            'addSheet': {
                'properties': {
                    'title': title,
                    'tabColor': {
                        'red': random.random(),
                        'green': random.random(),
                        'blue': random.random()
                    },
                    'gridProperties': {
                        'rowCount': 100,
                        'columnCount': num_columns
                    }
                }
            }
        }]
    }
    response = sheet.batchUpdate(spreadsheetId=WEBCAST_SPREADSHEET_ID, body=request_body).execute()
    print("Successfully created sheet \"" + title + "\" with " + str(num_columns) + " columns.")

def format_range_val(sheet_title, start_cell):
    """
    Formats range val in A1 notation. This is 'sheet name'!A1.
    """
    return "'" + sheet_title + "'!" + str(start_cell)

def insert_values(data, sheet):
    """
    Insert values to a given sheet in the Google Sheets spreadsheet specified by
    WEBCAST_SPREADSHEET_ID.
    """
    body = {
        'valueInputOption': "USER_ENTERED",
        'data': data
    }

    result = sheet.values().batchUpdate(spreadsheetId=WEBCAST_SPREADSHEET_ID, body=body).execute()
    print('{0} cells updated.'.format(result.get('totalUpdatedCells')))

def process_html_file(file_path, file_name, sheet):
    """
    Reads in an HTML file located at file_path + file_name, extracts all
    iterations of a class from that page, creates a new sheet in the
    WEBCAST_SPREADSHEET_ID spreadsheet with the name of the class (parsed from
    the HTML file), extracts the YouTube URLs from the page, and write all
    iterations of the class to the spreadsheet in a batch.

    Note: This will fail if a sheet with the given name already exists, so if
    this script fails while running and you restart it, make sure to delete the
    most recently created sheet if that sheet didn't finish being updated.
    """
    # Parse html website
    sheet_title = ""
    titles_seen = []

    f = open(file_path + file_name)
    soup = BeautifulSoup(f.read(), 'html.parser')

    class_iteration_sections = soup.find_all("div", class_="openberkeley-collapsible-container")

    sheet_title = soup.find_all("h1",class_="title")[0].text
    create_new_sheet(sheet_title, sheet, len(class_iteration_sections) * SHEET_SPACING)


    # To not hit the 100 writes per second limit, will use batchUpdate:
    data = []
    # We add entries of the form
    # {
    #     'range': range_name,
    #     'values': values
    # }
    # to the data list.

    start_index = 0;
    for section in class_iteration_sections:
        values_to_write = []

        iteration_title = section.find("h2", class_="openberkeley-collapsible-controller")

        # Skip content that is accidentally duplicated on the Webcast site.
        if not iteration_title or iteration_title.text in titles_seen:
            continue
        titles_seen.append(iteration_title.text)
        values_to_write.append([iteration_title.text, ""])
        link_list = section.find_all("a", attrs={"rel":"noreferrer"})
        for link in link_list:
            if "youtube" not in link.get("href"):
                continue
            values_to_write.append([link.text, link.get("href")])

        # Print starting from cell <letter>1, where letter = SHEET_SPACING *class_iteration
        start_cell = column_letter(start_index) + "1"
        data.append({
            'range': format_range_val(sheet_title, start_cell),
            'values': values_to_write
        })
        start_index += SHEET_SPACING

    insert_values(data, sheet)

    # Move processed file into a processed folder, to keep unprocessed and
    # processed files separate.
    os.rename(file_path + file_name, file_path + PROCESSED_FOLDER + file_name)

if __name__ == '__main__':
    main()
