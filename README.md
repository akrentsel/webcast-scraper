# UC Berkeley Webcast Link Scraper

This is a simple website scraper that processes HTML pages from UC Berkeley's webcast directory, located at [http://webcast.berkeley.edu/](http://webcast.berkeley.edu/), and writes them to a Google Sheet in the user's Google Drive.

I used this script to create this [Master Webcast Spreadsheet](https://docs.google.com/spreadsheets/d/17_o9tf34OCBVheyxjimPocUNg5JskP4nNwQ0blH3Zvc/edit?ouid=105312008256795736262&usp=sheets_home&ths=true) in May 2020.


## Usage

This script does **not** read directly from the website itself. It could be updated to do so, but this would require saving the user's cookie and passing it along with the requests, as the webcast directory is only accessible to logged in users. 

Instructions:
1. Clone this repo
2. Go to Google Sheets and create a new spreadsheet. The URL will be of the form `https://docs.google.com/spreadsheets/d/<id_here>/edit` - take note of the numbers and letters in the `<id_here>` part of the URL.
3. Past the id into the `WEBCAST_SPREADSHEET_ID`  variable initialization in the code.
4. Make a new folder in the cloned repo, called `webcast_html/`. Inside of that folder, create a new folder called `processed/`
5. Navigate to [http://webcast.berkeley.edu/](http://webcast.berkeley.edu/), sign in with your @berkeley.edu account, and open the webcast page(s) for whatever classes you want to save (for example, for the [page for CS61b](https://coursecapture.berkeley.edu/compsci-61b))
6. For each webcast page, hit Command (or Control) + S, and choose `Webpage, HTML Only` for the Format, and save the HTML file to the `webcast_html` directory you created above.
7. Run the script with `python3 webcast_scraper.py`. The first time you run it, you will need to authenticate with your Google account. 

Voila! You should see the Google sheet get filled in. The script will create a new tab in the spreadsheet for each webcast class file that you provide in the `webcast_html` folder.
