/**
* Obtains a request token from Pocket API and stores credentials.
* @param {string} consumerKey - Pocket API consumer key
* @returns {string} Authorization URL for user to authorize app
* @throws {Error} If consumer key is invalid or API request fails
*/
function getRequestToken(consumerKey) {
  try {
    // API endpoint for requesting token
    const url = 'https://getpocket.com/v3/oauth/request';

    // Configure API request
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        consumer_key: consumerKey,
        redirect_uri: "https://getpocket.com/home"
      }),
      muteHttpExceptions: true
    };

    // Make API request
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error('Invalid consumer key');
    }

    // Extract and store tokens
    const requestToken = response.getContentText().split("=")[1].trim();
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperties({
      "consumer_key": consumerKey,
      "request_token": requestToken
    });

    // Return authorization URL
    return `https://getpocket.com/auth/authorize?request_token=${requestToken}&redirect_uri=https://getpocket.com/home`;
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
* Completes OAuth flow by exchanging request token for access token.
* @returns {boolean} True if authorization successful
* @throws {Error} If authorization fails or tokens are missing
*/
function completeAuth() {
  try {
    // Get stored tokens
    const properties = PropertiesService.getScriptProperties();
    const requestToken = properties.getProperty("request_token");
    const consumerKey = properties.getProperty("consumer_key");

    if (!requestToken || !consumerKey) {
      throw new Error('Missing authentication tokens');
    }

    // Exchange request token for access token
    const response = UrlFetchApp.fetch("https://getpocket.com/v3/oauth/authorize", {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify({
        consumer_key: consumerKey,
        code: requestToken
      }),
      muteHttpExceptions: true
    });

    // Verify response
    if (response.getResponseCode() !== 200) {
      throw new Error('Authorization failed');
    }

    // Extract and store access token
    const accessToken = response.getContentText().match(/access_token=([^&]*)/)[1].trim();
    properties.setProperty("access_token", accessToken);
    return true;
  } catch (error) {
    Logger.log(error.stack);
  }
}