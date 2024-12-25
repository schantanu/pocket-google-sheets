# Pocket to Google Sheets Integration

A Google Apps Script project that syncs your Pocket articles with Google Sheets, enabling easy management and organization of your saved Pocket articles utilizing Google Sheets features. This project utilizes the Pocket v3 API to enable Google Sheets integration.

## Features

- Import all Pocket articles into Google Sheets
- Add new articles to Pocket directly from Google Sheets
- Update/delete articles in bulk
- View detailed article metadata (word count, reading time, tags, etc.)
- Filter and sort articles using Google Sheets features

## Setup

1. Create a new Google Sheet
2. Open Script Editor (Extensions > Apps Script)
3. Create the following files in your Apps Script project with the same name and case:
   - `main.gs`
   - `func_helper.gs`
   - `func_sheet.gs`
   - `sidebar_user_guide.html`
   - `sidebar_app_guide.html`
4. Create 'Script' file for `.gs` and 'HTML' file for `.html` extension
5. Once saved, refresh your Google Sheet to see the "Pocket Editor" menu
6. Follow the "User Guide" for the next steps to perform

## Usage

The extension adds a "Pocket Editor" menu with the following options:

- ℹ️ **User Guide**: User guide to understand the various features
- 🔓 **Pocket App Guide**: Setup instructions and app authentication
- 🛠️ **Setup/Reset Sheets**: Initialize or reset the Google Sheets
- ⬇️ **Get Articles**: Import articles from Pocket
- ⬆️ **Add Articles**: Add new articles to Pocket
- 🔄 **Update Articles**: Sync deleted articles back to Pocket

## Sheets Structure

### Articles Sheet
Displays all your Pocket articles with metadata:
- Item ID, Domain, Authors, Title, URL, Tags
- Time stamps (Added, Updated, Read, Favorite)
- Article properties (Word Count, Listen Duration, Status)

### Add Articles Sheet
Used for adding new articles to Pocket:
- URL column for article links
- Tags column for comma-separated tags

## Requirements

- Google account
- Pocket account
- Pocket API access

## License

MIT License

## Author

schantanu