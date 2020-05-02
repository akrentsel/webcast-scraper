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

# The ID and range of a sample spreadsheet.
WEBCAST_SPREADSHEET_ID = '17_o9tf34OCBVheyxjimPocUNg5JskP4nNwQ0blH3Zvc'
# SAMPLE_RANGE_NAME = 'Class Data!A2:E'

HTML_FILE_DIRECTORY = "./webcast_html/"
PROCESSED_FOLDER = "processed/"

# Number of columns between starts of new class records.
SHEET_SPACING = 3

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Process each file in the webcast_html directory:
    listed_files = [f for f in listdir(HTML_FILE_DIRECTORY) if isfile(join(HTML_FILE_DIRECTORY, f))]
    for filename in listed_files:
        if filename == ".DS_Store":
            continue
        print("Scraping " + filename)
        process_html_file(HTML_FILE_DIRECTORY, filename, sheet)

    # request_body = {
    #     'requests': [{
    #         'addSheet': {
    #             'properties': {
    #                 'title': 'test_title2',
    #                 'tabColor': {
    #                     'red': random.random(),
    #                     'green': random.random(),
    #                     'blue': random.random()
    #                 }
    #             }
    #         }
    #     }]
    # }
    #
    # values = [["a", "b"], ["c", "d"]]
    #
    # body = {
    #     'values': values
    # }
    #
    # sheet = service.spreadsheets()
    # response = sheet.batchUpdate(spreadsheetId=WEBCAST_SPREADSHEET_ID, body=request_body).execute()
    # print("Successfully added new sheet.")
    #
    #
    # # result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, valueInputOption="USER_ENTERED", range="A1", body=body).execute()
    # # print('{0} cells updated.'.format(result.get('updatedCells')))

# Return the ith letter of the alphabet, 0-indexed
def alphabet_letter(i):
    if i < 26:
        return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[i]
    else:
        first_letter_index = (i // 26) - 1
        second_letter_index = i % 26
        return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[first_letter_index] + "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[second_letter_index]


# Create a new sheet in the Google Sheets spreadsheet specified by
# WEBCAST_SPREADSHEET_ID. Creates the tab with a random colors.
def create_new_sheet(title, sheet, num_columns):
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

# Formats range val in A1 notation. This is 'sheet name'!A1.
def format_range_val(sheet_title, start_cell):
    return "'" + sheet_title + "'!" + str(start_cell)

# Insert values to a given sheet in the Google Sheets spreadsheet specified by
# WEBCAST_SPREADSHEET_ID.

def insert_values(sheet_title, values_to_write, start_cell, sheet):
        body = {
            'values': values_to_write
        }

        result = sheet.values().update(spreadsheetId=WEBCAST_SPREADSHEET_ID, valueInputOption="USER_ENTERED", range=format_range_val(sheet_title, start_cell), body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))

def process_html_file(file_path, file_name, sheet):
    # Parse html website
    sheet_title = ""
    titles_seen = []

    f = open(file_path + file_name)
    soup = BeautifulSoup(f.read(), 'html.parser')
    # print(soup.prettify())

    class_iteration_sections = soup.find_all("div", class_="openberkeley-collapsible-container")

    sheet_title = soup.find_all("h1",class_="title")[0].text
    create_new_sheet(sheet_title, sheet, len(class_iteration_sections) * SHEET_SPACING)

    start_index = 0;
    for section in class_iteration_sections:
        values_to_write = []

        iteration_title = section.find("h2", class_="openberkeley-collapsible-controller")

        # Skip content that is recorded twice.
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
        start_cell = alphabet_letter(start_index) + "1"
        insert_values(sheet_title, values_to_write, start_cell, sheet)
        start_index += SHEET_SPACING

    os.rename(file_path + file_name, file_path + PROCESSED_FOLDER + file_name)

if __name__ == '__main__':
    main()
