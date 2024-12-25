/**
 * Update the Font for the whole Spreadsheet.
 * Note: Need to execute only once.
 */
function updateFont() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSpreadsheetTheme().setFontFamily(defaultFont);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Set the Font Size for a given sheet.
 * @param {string} sheetName - The sheet to change the font size of.
 */
function setFontSize(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Set Font Size for whole sheet
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    range.setFontSize(fontSize);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Insert sheet, if it does not exist else clear formatting and data.
 * @param {string} sheetName - The sheet name.
 */
function resetSheet(sheetName, clearType = 'formats') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Insert new sheet if not exists
    if (!sheet) {
      ss.insertSheet(sheetName);
    } else {
      // Clear based on clearType parameter
      switch(clearType) {
        case 'all':
          sheet.clear();          // Clear formatting and data
          break;
        case 'formats':
        default:
          sheet.clearFormats();   // Clear formatting only
      }
    }

    // Check if the sheet exists and has a filter, then remove the filter
    if (sheet && typeof sheet.getFilter === 'function') {
      const filter = sheet.getFilter();
      if (filter) {
        filter.remove();
      }
    }

    // Set Font Size
    setFontSize(sheetName);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
* Sets up the articles spreadsheet with formatted headers and column configuration
* This sheet stores the main Pocket articles data including metadata and status
*
* @function setupArticlesSheet
* @throws {Error} If sheet setup fails
 */
function setupArticlesSheet() {
  try {
    // Reset sheet
    resetSheet(articlesSheetName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(articlesSheetName);

    // Setup headers
    const headers = [[
      'Item ID','Domain','Authors','Title','URL','Tags','Time Added','Time Updated','Favorite','Status',
      'Word Count','Listen Duration (mins)','Time Read','Time Favorited','Is Article?','Has Video?','Has Image?'
    ]];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

    // Format headers
    headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('Bold');
    headerRange.setBackground('#c9daf8');
    sheet.setFrozenRows(1);

    // Add header notes
    const headerNotes = {
      'Favorite': '1 = favorited item',
      'Status': '1 = archived, 2 = deleted',
      'Is Article?': '1 = item is an article',
      'Has Video?': '1 = contains videos, 2 = is a video',
      'Has Image?': '1 = contains images, 2 = is an image'
    };

    // Add after setting headers
    Object.entries(headerNotes).forEach(([header, note]) => {
      const col = headers[0].indexOf(header) + 1;
      if (col > 0) {
        sheet.getRange(1, col).setNote(note);
      }
    });

    // Set Column widths
    sheet.setColumnWidth(1, 80);      // Item ID
    sheet.setColumnWidth(2, 150);     // Domain
    sheet.setColumnWidth(3, 150);     // Authors
    sheet.setColumnWidth(4, 500);     // Title
    sheet.setColumnWidth(5, 150);     // URL
    sheet.setColumnWidth(6, 160);     // Tags
    sheet.setColumnWidth(7, 150);     // Time Added
    sheet.setColumnWidth(8, 150);     // Time Updated
    sheet.setColumnWidth(9, 80);      // Favorite
    sheet.setColumnWidth(10, 80);     // Status
    sheet.setColumnWidth(11, 100);    // Word Count
    sheet.setColumnWidth(12, 165);    // Listen Duration
    sheet.setColumnWidth(13, 150);    // Time Read
    sheet.setColumnWidth(14, 150);    // Time Favorited
    sheet.setColumnWidth(15, 100);    // Is Article?
    sheet.setColumnWidth(16, 100);    // Has Video?
    sheet.setColumnWidth(17, 100);    // Has Image?

    // Set Wrap Strategy
    articlesDataRange = sheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns());
    articlesDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
* Sets up the add articles spreadsheet with formatted headers and column configuration
* This sheet is used to add articles to Pocket articles data including tags
*
* @function setupAddArticlesSheet
* @throws {Error} If sheet setup fails
 */
function setupAddArticlesSheet() {
  try {
    // Reset sheet
    resetSheet(addArticlesSheetName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(addArticlesSheetName);

    // Setup headers
    const headers = [['Add URLs to Pocket','Add Tags']];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

    // Format Headers
    headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('Bold');
    headerRange.setBackground('#c9daf8');
    sheet.setFrozenRows(1);

    // Add header notes
    const headerNotes = {
      'Add URLs to Pocket': 'Add only URLs in this column',
      'Add Tags': `Add optional tags in this column. For multiple tags per URL use commas, for e.g., 'books, fiction'`
    };

    // Add after setting headers
    Object.entries(headerNotes).forEach(([header, note]) => {
      const col = headers[0].indexOf(header) + 1;
      if (col > 0) {
        sheet.getRange(1, col).setNote(note);
      }
    });

    // Set Column widths
    sheet.setColumnWidth(1, 500);      // URL
    sheet.setColumnWidth(2, 160);      // Tags

    // Set Wrap Strategy
    addArticlesDataRange = sheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns());
    addArticlesDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } catch (error) {
    Logger.log(error.stack);
  }
}