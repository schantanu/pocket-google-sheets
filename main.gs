// Set Constants
const defaultFont = 'Open Sans';
const fontSize = 9;
const articlesSheetName = 'Articles';
const addArticlesSheetName = 'Add Articles';

/**
 * Creates menu items in the Google Sheets UI.
 * @returns {void}
 */
function onOpen() {
  // Create Pocket App UI
  SpreadsheetApp.getUi()
    .createMenu('✍️ Pocket App')
    .addItem('ℹ️ User Guide', 'showUserGuide')
    .addItem('🔓 Pocket App Guide', 'showAppGuideSidebar')
    .addItem('🛠️ Setup/Reset Sheets', 'setupSheets')
    .addSeparator()
    .addItem('⬇️ Get Articles', 'getArticles')
    .addItem('⬆️ Add Articles', 'addArticles')
    .addItem('🔄 Update Articles', 'updateArticles')
    .addToUi();
}

/**
 * Shows user guide.
 */
function showUserGuide() {
  const html = HtmlService.createTemplateFromFile('sidebar_user_guide')
    .evaluate()
    .setTitle('User Guide');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows app authentication creation sidebar.
 */
function showAppGuideSidebar() {
  const html = HtmlService.createTemplateFromFile('sidebar_app_guide')
    .evaluate()
    .setTitle('Pocket App Guide');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Creates or resets sheets for article management.
 * @returns {void}
 * @throws {Error} If sheet creation fails
 */
function setupSheets() {
  try {
    // Set default font
    updateFont();

    // Setup Articles Sheet
    setupArticlesSheet();

    // Setup Add Articles Sheet
    setupAddArticlesSheet();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Retrieves script properties for Pocket API authentication.
 * @returns {Properties} Object containing consumer key and access token
 * @throws {Error} If required properties are missing
 */
function getScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const consumerKey = scriptProperties.getProperty('consumer_key');
  const accessToken = scriptProperties.getProperty('access_token');

  if (!consumerKey || !accessToken) {
    SpreadsheetApp.getUi().alert(`To run this script, please follow the 'Pocket App Guide' and setup the app first.`);
    return null;
  }

  return { consumerKey, accessToken };
}

/**
 * Stores the current state of article IDs in the script cache.
 * @throws {Error} If unable to store the initial state
 */
function storeInitialState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(articlesSheetName);

    if (!sheet) {
      throw new Error(`Sheet '${articlesSheetName}' not found`);
    }

    if (sheet.getLastRow() <= 1) {
      Logger.log('No data found in the sheet to store.');
      return;
    }

    const itemIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(id => id);

    const cache = CacheService.getScriptCache();
    cache.put("previous_item_ids", JSON.stringify(itemIds));
  } catch (error) {
    Logger.log(`Store Initial State Error: ${error.stack}`);
    throw error;
  }
}

/**
 * Fetches and processes articles from Pocket API.
 * @returns {void}
 * @throws {Error} If API calls fail
 */
function getArticles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(articlesSheetName);

    if (!sheet) {
      setupArticlesSheet();
      sheet = ss.getSheetByName(articlesSheetName);
    }

    // Get and validate credentials
    const properties = getScriptProperties();
    if (!properties) {
      throw new Error('Missing authentication credentials');
    }

    const url = 'https://getpocket.com/v3/get';
    const headers = [
      'item_id', 'domain_metadata.name', 'authors', 'resolved_title', 'resolved_url', 'tags', 'time_added',
      'time_updated', 'favorite', 'status', 'word_count', 'listen_duration_estimate', 'time_read', 'time_favorited',
      'is_article', 'has_video', 'has_image'
    ];
    const batchSize = 30;
    let allArticles = [];
    let offset = 0;

    while (true) {
      const payload = {
        'consumer_key': properties.consumerKey,
        'access_token': properties.accessToken,
        'detailType': 'complete',
        'state': 'all',
        'count': batchSize,
        'offset': offset
      };

      const options = {
        'method': 'POST',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();

      Logger.log(`API Status: ${statusCode}, Offset: ${offset}`);
      Logger.log(`Response Preview: ${responseText.substring(0, 200)}`);

      if (statusCode !== 200) {
        throw new Error(`API request failed. This might be due to the API rate limit.
          \nPlease try again later.\n\n
          More Info:\nhttps://getpocket.com/developer/docs/rate-limits
          \n\nResponse Code: ${statusCode}\nResponse Text: ${responseText}`);
      }

      const responseData = JSON.parse(responseText);

      if (!responseData || typeof responseData !== 'object') {
        throw new Error('Invalid API response format');
      }

      if (!responseData.list || Object.keys(responseData.list).length === 0) {
        Logger.log('No more articles found');
        break;
      }

      const currentBatchSize = Object.keys(responseData.list).length;
      Logger.log(`Received batch of ${currentBatchSize} articles`);

      const batchArticles = Object.values(responseData.list)
        .filter(article => article && article.status !== "2")
        .map(article => {
          article.time_added = parseInt(article.time_added) || 0;
          return article;
        });

      allArticles.push(...batchArticles);
      offset += batchSize;
    }

    Logger.log(`Total articles retrieved: ${allArticles.length}`);

    // Sort articles by time_added in descending order
    allArticles.sort((a, b) => b.time_added - a.time_added);

    // Convert articles to values array
    const values = allArticles.map(article => headers.map(key => {
      let value = key.split('.').reduce((prev, curr) => (prev && prev[curr] ? prev[curr] : ''), article);

      // Convert Unix timestamps to readable dates
      if (['time_added', 'time_updated', 'time_read', 'time_favorited'].includes(key)) {
        value = value === '0' ? '' : new Date(value * 1000).toLocaleString();
      }

      // Convert listen_duration_estimate from seconds to minutes
      if (key === 'listen_duration_estimate' && value) {
        value = Math.round(value / 60);
      }

      // Extract tags as a comma-separated string
      if (key === 'tags' && article.tags) {
        value = Object.keys(article.tags).join(', ');
      }

      // Extract authors as a comma-separated string
      if (key === 'authors' && article.authors) {
        value = Object.values(article.authors).map(author => author.name).join(', ');
      }

      return value;
    }));

    // Clear the sheet contents
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent();
    }

    // Write the data to the sheet
    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    }

    // Set filter and sort the data by Sort ID
    const filter = sheet.getFilter();
    if (!filter) {
      sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
    } else {
      filter.remove();
      sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
    }

    // Sort by time_added
    if (sheet.getLastRow() > 1 && sheet.getLastColumn() > 0) {
      sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).sort({ column: 7, ascending: false });
    }

    // Cache urls
    storeInitialState();

    SpreadsheetApp.getUi().alert(`Successfully retrieved ${allArticles.length} articles`);

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/**
 * Adds articles to Pocket from the Add Articles sheet.
 * Processes and removes successful uploads while preserving failed entries.
 * Handles articles sequentially to avoid row skipping issues.
 *
 * @throws {Error} If sheet access fails or API calls fail
 * @returns {void}
 */
function addArticles() {
  try {
    // Initialize sheet access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(addArticlesSheetName);
    const BATCH_SIZE = 10; // Number of successful rows to delete at once

    // Validate sheet exists
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`The '${addArticlesSheetName}' sheet was not found. Please run 'Setup/Reset Sheets' first.`);
      return;
    }

    const properties = getScriptProperties();
    let lastRow = sheet.getLastRow(); // Mutable to track row count after deletions

    // Check for data
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert(`No articles found in the '${addArticlesSheetName}' sheet.`);
      return;
    }

    // Setup processing variables
    const url = 'https://getpocket.com/v3/add';
    const stats = { successCount: 0, failureCount: 0 };
    const successfulRows = []; // Track successful rows for batch deletion

    // Process each row sequentially to maintain accuracy
    for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
      const rowData = sheet.getRange(rowIndex, 1, 1, 2).getValues()[0];
      const articleUrl = rowData[0];
      const tags = rowData[1] ? rowData[1].split(',').map(tag => tag.trim()).join(',') : '';

      // Skip empty rows
      if (!articleUrl) {
        Logger.log(`Row ${rowIndex}: Skipped empty URL.`);
        stats.failureCount++;
        continue;
      }

      // Prepare API payload
      const payload = {
        consumer_key: properties.consumerKey,
        access_token: properties.accessToken,
        url: articleUrl,
      };

      if (tags) {
        payload.tags = tags;
      }

      // Attempt API call
      try {
        const response = UrlFetchApp.fetch(url, {
          method: 'POST',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          stats.successCount++;
          successfulRows.push(rowIndex);
          Logger.log(`Row ${rowIndex}: Added successfully - ${articleUrl}`);
        } else {
          stats.failureCount++;
          Logger.log(`Row ${rowIndex}: Failed - ${articleUrl}. Response: ${response.getContentText()}`);

          throw new Error(`API request failed. This might be due to the API rate limit.
            \nPlease try again later.\n\n
            More Info:\nhttps://getpocket.com/developer/docs/rate-limits
            \n\nResponse Code: ${statusCode}\nResponse Text: ${responseText}`);
        }
      } catch (error) {
        stats.failureCount++;
        Logger.log(`Row ${rowIndex}: Error - ${error.message}`);
      }

      // Batch delete successful rows to improve performance
      if (successfulRows.length >= BATCH_SIZE) {
        successfulRows.sort((a, b) => b - a); // Sort descending for bottom-up deletion
        successfulRows.forEach(rowNum => {
          sheet.deleteRow(rowNum);
        });
        successfulRows.length = 0;
        rowIndex -= BATCH_SIZE; // Adjust index for deleted rows
        lastRow = sheet.getLastRow(); // Update total row count
      }
    }

    // Delete any remaining successful rows
    if (successfulRows.length > 0) {
      successfulRows.sort((a, b) => b - a);
      successfulRows.forEach(rowNum => {
        sheet.deleteRow(rowNum);
      });
    }

    // Refresh articles list if any were added
    if (stats.successCount > 0) {
      getArticles();
    }

    // Show results to user
    SpreadsheetApp.getUi().alert(
      `Articles processed.\nSuccess: ${stats.successCount}\nFailed: ${stats.failureCount}\n` +
      (stats.failureCount > 0 ? '\nFailed articles remain in the sheet for retry.' : '')
    );

  } catch (error) {
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert(`${error.message}\nPartially processed articles are saved.`);
  }
}

/**
 * Updates articles in Pocket based on spreadsheet changes.
 * @returns {Promise<void>}
 * @throws {Error} If update fails
 */
function updateArticles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(articlesSheetName);

    if (!sheet) {
      SpreadsheetApp.getUi().alert(`The '${articlesSheetName}' sheet was not found. Please setup the sheet by running the 'Get Articles' script.`);
      return;
    }

    // Get consumer key and access token
    const properties = getScriptProperties();

    // Get the current state of the Item IDs in the sheet
    const currentItemIds = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().filter(id => id);

    // Get previously stored Item IDs from the cache
    const cache = CacheService.getScriptCache();
    const previousItemIds = JSON.parse(cache.get('previous_item_ids') || '[]');

    // Find Item IDs that have been removed
    const removedItemIds = previousItemIds.filter(id => !currentItemIds.includes(id));
    if (removedItemIds.length === 0) {
      SpreadsheetApp.getUi().alert('No articles were removed.');
      return;
    }

    const url = 'https://getpocket.com/v3/send';

    // Remove each Item ID that has been deleted from the Pocket account
    removedItemIds.forEach(function (itemId) {
      const payload = {
        'consumer_key': properties.consumerKey,
        'access_token': properties.accessToken,
        'actions': [{
          'action': 'delete',
          'item_id': itemId
        }]
      };
      const options = {
        'method': 'POST',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload)
      };
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode === 200) {
        Logger.log('Article removed successfully: ' + itemId);
      } else {
        Logger.log('Failed to remove article: ' + itemId);
      }
    });

    // Update cache with the current Item IDs
    cache.put('previous_item_ids', JSON.stringify(currentItemIds));

    SpreadsheetApp.getUi().alert('Articles removed successfully from Pocket!\nGetting latest Articles from Pocket.');

    // Get Articles
    getArticles();

  } catch (error) {
    Logger.log(error.stack);
  }
}