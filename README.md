# contacts-birthday-sheet

A Google Apps Script that reads all your Google Contacts and their birthdays and writes them to a Google Sheet, sorted by month and day.

## Features

- Pulls all contacts with birthdays from Google Contacts via the People API
- Uses the "Speichern unter" (File as) name when set, otherwise falls back to the display name
- Sorts contacts by month → day
- Writes results to a sheet named **Birthdays** with columns: `Name | Birthday | Day | Month | Year`
- Formatted header row and auto-resized columns

## Setup

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Go to **Extensions → Apps Script**
3. Delete the default code and paste the contents of `birthday_to_sheet.gs`
4. Click **Save**
5. Enable the People API: click **Services** (+ icon) → find **Google People API** → **Add**
6. Select `syncBirthdaysToSheet` from the function dropdown and click **Run**
7. Grant the requested permissions when prompted

## Functions

| Function | Description |
|---|---|
| `syncBirthdaysToSheet` | Main function — fetches contacts and writes to the sheet |
| `debugNameFields` | Logs raw API data for the first 20 contacts to help diagnose name field issues |
