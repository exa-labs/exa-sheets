function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Exa AI')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('About/Help', 'showAbout')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Exa AI Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  var message = 'Exa AI for Google Sheets\n\n' +
                'Version: 1.0.0\n\n' +
                'This add-on provides powerful AI-driven search and analysis capabilities using the Exa API.\n\n' +
                'Key Features:\n' +
                '• EXA_ANSWER: Generate AI answers from web searches\n' +
                '• EXA_SEARCH: Perform web searches\n' +
                '• EXA_CONTENTS: Extract content from URLs\n' +
                '• EXA_FINDSIMILAR: Find similar web pages\n\n' +
                'For detailed documentation and support, open the sidebar and navigate to the Docs tab.\n\n' +
                'Visit https://exa.ai for more information about the Exa API.\n\n' +
                'Visit https://github.com/BrandonGuocamole/exa-for-sheets the full project repository.';
  
  ui.alert('About Exa AI', message, ui.ButtonSet.OK);
}

/**
 * Retrieves all stored API keys
 * @return {Object} Object containing all saved API keys and metadata
 */
function getAllApiKeys() {
  const keysJson = PropertiesService.getUserProperties().getProperty('EXA_API_KEYS');
  if (!keysJson) {
    return { keys: {}, activeKeyName: null };
  }
  
  try {
    return JSON.parse(keysJson);
  } catch (e) {
    // If there's an error parsing the JSON, return empty structure
    return { keys: {}, activeKeyName: null };
  }
}

/**
 * Saves API keys data to UserProperties
 * @param {Object} keysData Object containing all keys and the active key name
 */
function saveAllApiKeys(keysData) {
  PropertiesService.getUserProperties().setProperty('EXA_API_KEYS', JSON.stringify(keysData));
}

/**
 * Saves a new API key with a name
 * @param {string} name A friendly name for the API key
 * @param {string} key The Exa API key to save
 * @param {boolean} setActive Whether to set this key as the active key
 * @return {Object} Status object with success flag and message
 */
function saveApiKey(name, key, setActive = true) {
  if (!name || !key) {
    return { 
      success: false, 
      message: 'Both name and API key are required.'
    };
  }
  
  // Get all existing keys
  const keysData = getAllApiKeys();
  
  // Check if a key with this name already exists
  if (keysData.keys[name]) {
    return { 
      success: false, 
      message: `A key named "${name}" already exists. Choose a different name.`
    };
  }
  
  // Add the new key with metadata
  const now = new Date().toISOString();
  keysData.keys[name] = {
    key: key,  // Store the actual API key
    created: now,
    lastUsed: now,
    // First few and last few characters for display, rest is masked
    displayKey: `${key.substring(0, 4)}...${key.substring(key.length - 4)}`
  };
  
  // Set as active key if requested or if it's the only key
  if (setActive || !keysData.activeKeyName) {
    keysData.activeKeyName = name;
  }
  
  // Save back to properties
  saveAllApiKeys(keysData);
  
  return { 
    success: true, 
    message: `Key "${name}" saved successfully${setActive ? ' and set as active' : ''}.`
  };
}

/**
 * Deletes an API key by name
 * @param {string} name The name of the key to delete
 * @return {Object} Status object with success flag and message
 */
function deleteApiKey(name) {
  const keysData = getAllApiKeys();
  
  if (!keysData.keys[name]) {
    return { 
      success: false, 
      message: `Key "${name}" not found.`
    };
  }
  
  // Delete the key
  delete keysData.keys[name];
  
  // If we deleted the active key, set a new active key or clear it
  if (keysData.activeKeyName === name) {
    const remainingKeys = Object.keys(keysData.keys);
    keysData.activeKeyName = remainingKeys.length > 0 ? remainingKeys[0] : null;
  }
  
  // Save the changes
  saveAllApiKeys(keysData);
  
  return { 
    success: true, 
    message: `Key "${name}" deleted successfully.`
  };
}

/**
 * Sets the active API key by name
 * @param {string} name The name of the key to set as active
 * @return {Object} Status object with success flag and message
 */
function setActiveApiKey(name) {
  const keysData = getAllApiKeys();
  
  if (!keysData.keys[name]) {
    return { 
      success: false, 
      message: `Key "${name}" not found.`
    };
  }
  
  // Set the active key
  keysData.activeKeyName = name;
  
  // Update last used timestamp
  keysData.keys[name].lastUsed = new Date().toISOString();
  
  // Save the changes
  saveAllApiKeys(keysData);
  
  return { 
    success: true, 
    message: `Key "${name}" is now active.`
  };
}

/**
 * Gets the currently active API key value for use in API calls
 * @return {string|null} The active API key or null if no key is set
 */
function getApiKey() {
  const keysData = getAllApiKeys();
  
  if (!keysData.activeKeyName || !keysData.keys[keysData.activeKeyName]) {
    return null;
  }
  
  // Update last used timestamp
  keysData.keys[keysData.activeKeyName].lastUsed = new Date().toISOString();
  saveAllApiKeys(keysData);
  
  return keysData.keys[keysData.activeKeyName].key;
}

/**
 * Gets information about the active key for display in UI
 * @return {Object} Object with active key info or null if no active key
 */
function getActiveKeyInfo() {
  const keysData = getAllApiKeys();
  
  if (!keysData.activeKeyName || !keysData.keys[keysData.activeKeyName]) {
    return null;
  }
  
  const activeKey = keysData.keys[keysData.activeKeyName];
  return {
    name: keysData.activeKeyName,
    displayKey: activeKey.displayKey,
    created: activeKey.created,
    lastUsed: activeKey.lastUsed
  };
}

/**
 * Helper function that formats API keys data for display in the UI
 * @return {Object} Object with activeKey and keys properties
 */
function getAllApiKeysForUI() {
  const keysData = getAllApiKeys();
  const result = {
    keys: {},
    activeKey: null
  };
  
  // Process each key to create the UI representation
  if (keysData && keysData.keys) {
    Object.entries(keysData.keys).forEach(([name, keyData]) => {
      result.keys[name] = {
        displayKey: keyData.displayKey,
        created: keyData.created,
        lastUsed: keyData.lastUsed
      };
    });
  }
  
  // Set the active key info
  if (keysData.activeKeyName && keysData.keys[keysData.activeKeyName]) {
    const activeKey = keysData.keys[keysData.activeKeyName];
    result.activeKey = {
      name: keysData.activeKeyName,
      displayKey: activeKey.displayKey,
      created: activeKey.created,
      lastUsed: activeKey.lastUsed
    };
  }
  
  return result;
}

/**
 * Queries the Exa /answer endpoint to provide an AI-generated answer based on search results.
 * Allows adding prefix/suffix text and optionally includes source citations.
 * By default, extracts and returns only the core answer text before any inline citations like " ([Source](URL)...)".
 *
 * @param {string} prompt The main question or prompt to send to Exa. Can be a cell reference.
 * @param {string} [prefix=""] Optional. Text to add before the main prompt.
 * @param {string} [suffix=""] Optional. Text to add after the main prompt.
 * @param {boolean} [includeCitations=FALSE] Optional. If TRUE, returns the full answer string as received from the API (including inline citations) AND appends any additional citations from the API's 'citations' array. Defaults to FALSE (extracts core answer only).
 * @return {string} The core answer, the full answer with citations, or an error message.
 * @customfunction
 */
function EXA_ANSWER(prompt, prefix, suffix, includeCitations) {
  const apiKey = getApiKey();
  if (!apiKey) return "❌ No API key set. Please use 'Set API Key' in the menu.";

  // --- Parameter Validation and Processing ---
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === "") {
    return "❌ Please provide a valid prompt/question.";
  }

  const finalPrompt = `${prefix || ''} ${prompt} ${suffix || ''}`.trim();
  // Explicitly check if includeCitations is exactly TRUE.
  const shouldShowFullAnswerWithCitations = includeCitations === true;

  // --- API Call ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/answer", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ query: finalPrompt }),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    // --- Response Handling ---
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);

      if (result && typeof result.answer === 'string') {
        const fullAnswerFromApi = result.answer;
        let finalOutput = fullAnswerFromApi; // Default to the full answer

        if (!shouldShowFullAnswerWithCitations) {
          // --- Default Behavior: Extract Core Answer ---
          // Find the start of the citation block pattern " (["
          const citationStartIndex = fullAnswerFromApi.indexOf(" ([");

          if (citationStartIndex !== -1) {
            // If the pattern is found, extract the text before it
            finalOutput = fullAnswerFromApi.substring(0, citationStartIndex).trim();
          } else {
            // If the pattern isn't found, assume no inline citations; use the full answer
            finalOutput = fullAnswerFromApi;
          }
          // Now finalOutput contains only the core answer part (hopefully)

        } else {
          // --- Include Citations Behavior ---
          // Use the full answer from API, and append from the citations array if present
           finalOutput = fullAnswerFromApi; // Start with the full API answer

           if (result.citations && Array.isArray(result.citations) && result.citations.length > 0) {
              const formattedCitations = result.citations.map(citation => {
                  const title = citation.title || 'Source';
                  const url = citation.url;
                  return url ? `([${title}](${url}))` : null;
              })
              .filter(c => c !== null)
              .join(', ');

              if (formattedCitations.length > 0) {
                  const separator = (finalOutput.slice(-1) !== ' ' && finalOutput.slice(-1) !== '\n') ? ' ' : '';
                  finalOutput += separator + formattedCitations; // Append the extra citations
              }
           }
           // If includeCitations is true but no citations array, finalOutput remains the fullAnswerFromApi
        }

        return finalOutput; // Return the processed output

      } else {
        return "❌ API returned a valid response, but no 'answer' field (string) was found.";
      }
    } else if (responseCode === 401) {
      return "❌ API Error: Invalid API Key.";
    } else { // Handle other errors
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return errorMessage;
    }
  } catch (e) {
    Logger.log(`EXA_ANSWER Error: ${e} for prompt: ${finalPrompt}`);
    return `❌ Script Error: ${e.message}`;
  }
}

/**
 * Retrieves the text content of a given URL using the Exa /contents endpoint.
 *
 * @param {string} url The full URL (including http/https) to fetch content from.
 * @return {string} The main text content of the URL or an error message.
 * @customfunction
 */
function EXA_CONTENTS(url) {
  const apiKey = getApiKey();
  if (!apiKey) return "❌ No API key set. Please use 'Set API Key' in the menu.";

  // Basic URL validation
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
      return "❌ Please provide a valid URL starting with http or https.";
  }

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/contents", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ urls: [url] }), // Exa's /contents endpoint expects an array of URLs
      headers: { "x-api-key": apiKey }, // Corrected header based on Exa Docs
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
        const result = JSON.parse(responseBody);
        // The response structure for /contents typically wraps results in a 'results' array
        const contentData = result.results && result.results[0];
        if (contentData) {
            // Prioritize text content based on Exa's common response structure
            return contentData.text || contentData.highlights || "No relevant content found in response.";
        } else {
            return "❌ API returned successfully, but no content data found for this URL.";
        }
    } else if (responseCode === 401) {
        return "❌ API Error: Invalid API Key. Please check your key in the menu.";
    } else {
        // Try to parse error message from API if possible, otherwise return generic error
        let errorMessage = `❌ API Error: Received status code ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return errorMessage;
    }
  } catch (e) {
    return `❌ Script Error: ${e.message}`;
  }
}

/**
 * Finds URLs similar to the input URL using the Exa /findSimilar endpoint, with optional filters.
 * Returns a vertical list of similar URLs.
 *
 * @param {string} url The URL to find similar links for (must include http/https).
 * @param {number} [numResults=1] Optional. The maximum number of similar URLs to return (1-10). Defaults to 1.
 * @param {string} [includeDomainsStr=""] Optional. Comma-separated list of domains to restrict results to (e.g., "example.com,anotherexample.org").
 * @param {string} [excludeDomainsStr=""] Optional. Comma-separated list of domains to exclude from results (e.g., "exclude.net,badsite.co").
 * @param {string} [includeTextStr=""] Optional. A phrase that MUST be present in the content of result pages.
 * @param {string} [excludeTextStr=""] Optional. A phrase that MUST NOT be present in the content of result pages.
 * @return {string[][]} A vertical array of similar URLs, or a single cell error message.
 * @customfunction
 */
function EXA_FINDSIMILAR(url, numResults, includeDomainsStr, excludeDomainsStr, includeTextStr, excludeTextStr) {
  const apiKey = getApiKey();
  if (!apiKey) return [["❌ No API key set. Please use 'Set API Key' in the menu."]];

  // --- Parameter Validation and Processing ---
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
    return [["❌ Please provide a valid URL starting with http or https."]];
  }

  // Validate and set numResults (sensible default and limits)
  const count = (typeof numResults === 'number' && numResults >= 1 && numResults <= 10)
                ? Math.floor(numResults)
                : 1; // Default to 1 if invalid, NaN, or outside 1-10 range

  // Process domain lists (comma-separated string to array)
  const processDomains = (domainStr) => {
    if (typeof domainStr === 'string' && domainStr.trim() !== '') {
      return domainStr.split(',').map(d => d.trim()).filter(d => d.length > 0);
    }
    return null; // Return null if empty or not a string
  };

  const includeDomains = processDomains(includeDomainsStr);
  const excludeDomains = processDomains(excludeDomainsStr);

  // Process text filters (use the string directly if provided)
  const includeText = (typeof includeTextStr === 'string' && includeTextStr.trim() !== '') ? [includeTextStr.trim()] : null;
  const excludeText = (typeof excludeTextStr === 'string' && excludeTextStr.trim() !== '') ? [excludeTextStr.trim()] : null;
  // Note: Exa API docs mention limit of 1 string, 5 words for text filters. We send as array[1].

  // --- Build API Payload ---
  const payload = {
    url: url,
    numResults: count,
    excludeSourceDomain: true // Good default to avoid getting the input URL back
  };

  if (includeDomains && includeDomains.length > 0) {
    payload.includeDomains = includeDomains;
  }
  if (excludeDomains && excludeDomains.length > 0) {
    payload.excludeDomains = excludeDomains;
  }
  if (includeText && includeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.includeText = includeText.slice(0, 1);
  }
  if (excludeText && excludeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.excludeText = excludeText.slice(0, 1);
  }

  // --- API Call and Response Handling ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/findSimilar", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        return result.results.map(item => [item.url || "N/A"]); // Map URLs, provide fallback
      } else {
        return [["✅ No similar URLs found matching the criteria."]];
      }
    } else if (responseCode === 401) {
      return [["❌ API Error: Invalid API Key."]];
    } else if (responseCode === 400) {
        // Handle potential bad request errors (e.g., invalid filters)
        let errorMessage = `❌ API Error (Bad Request): Status ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return [[errorMessage]];
    }
    else { // Handle other errors
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    // Catch script execution errors (e.g., network issues)
    Logger.log(`EXA_FINDSIMILAR Error: ${e} for payload: ${JSON.stringify(payload)}`); // Log for debugging
    return [[`❌ Script Error: ${e.message}`]];
  }
}

/**
 * Searches the web using the Exa /search endpoint based on a query.
 * Returns a vertical list of result URLs.
 *
 * @param {string} query The search query.
 * @param {number} [numResults=1] Optional. The maximum number of result URLs to return. Defaults to 1.
 * @param {string} [searchType="auto"] Optional. The type of search ('auto', 'neural', 'keyword'). Defaults to 'auto'.
 * @param {string} [prefix=""] Optional. Text to add before the main query.
 * @param {string} [suffix=""] Optional. Text to add after the main query.
 * @return {string[]} An array of result URLs or a single cell error message.
 * @customfunction
 */
function EXA_SEARCH(query, numResults, searchType, prefix, suffix) {
  const apiKey = getApiKey();
  if (!apiKey) return [["❌ No API key set. Please use 'Set API Key' in the menu."]];

  if (!query || typeof query !== 'string') {
    return [["❌ Please provide a valid search query."]];
  }

  // Process the query with optional prefix and suffix
  const finalQuery = `${prefix || ''} ${query} ${suffix || ''}`.trim();

  const count = (typeof numResults === 'number' && numResults > 0 && numResults <= 10) ? Math.floor(numResults) : 1; // Default to 1, max 10 for free tier
  const type = (searchType && ['auto', 'neural', 'keyword'].includes(searchType)) ? searchType : 'auto'; // Default to 'auto'

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/search", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        query: finalQuery,
        numResults: count,
        type: type,
        useAutoprompt: (type !== 'keyword') // Enable autoprompt for neural/auto by default
      }),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

     if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        // Map results to a 2D array for vertical spill
        return result.results.map(item => [item.url]);
      } else {
        return [["✅ API returned successfully, but no search results found."]];
      }
    } else if (responseCode === 401) {
      return [["❌ API Error: Invalid API Key. Please check your key in the menu."]];
    } else {
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    return [[`❌ Script Error: ${e.message}`]];
  }
}

/**
 * Backend function called by the sidebar's Function Builder to test an EXA function.
 *
 * @param {string} functionName The name of the EXA function to test (e.g., "EXA_ANSWER").
 * @param {object} args An object containing the arguments for the function, keyed by parameter name (e.g., { prompt: "Test prompt", includeCitations: false }).
 * @return {any} The result from the called EXA function, or an error string.
 */
function testExaFunction(functionName, args) {
  Logger.log(`testExaFunction called with: ${functionName}, Args: ${JSON.stringify(args)}`);

  // Ensure API key exists before proceeding
  const apiKey = getApiKey();
  if (!apiKey) return "❌ No API key set. Please use 'Set API Key' in the menu.";

  try {
    switch (functionName) {
      case "EXA_ANSWER":
        if (!args || typeof args.prompt !== 'string' || args.prompt.trim() === "") {
          return "❌ Missing required argument: prompt";
        }
        // Call EXA_ANSWER with arguments from the args object
        return EXA_ANSWER(
          args.prompt,
          args.prefix || "", // Default optional args to empty string or expected default
          args.suffix || "",
          args.includeCitations || false
        );

      case "EXA_CONTENTS":
        if (!args || typeof args.url !== 'string' || !args.url.startsWith('http')) {
           return "❌ Missing or invalid required argument: url (must start with http/https)";
        }
        // Call EXA_CONTENTS
        return EXA_CONTENTS(args.url);

      case "EXA_FINDSIMILAR":
        if (!args || typeof args.url !== 'string' || !args.url.startsWith('http')) {
          return [["❌ Missing or invalid required argument: url (must start with http/https)"]]; // Return as array for consistency
        }
        // Call EXA_FINDSIMILAR, passing potentially null values for optional args
        return EXA_FINDSIMILAR(
          args.url,
          args.numResults, // EXA_FINDSIMILAR handles undefined/null internally
          args.includeDomainsStr,
          args.excludeDomainsStr,
          args.includeTextStr,
          args.excludeTextStr
        );

       case "EXA_SEARCH":
         if (!args || typeof args.query !== 'string' || args.query.trim() === "") {
            return [["❌ Missing required argument: query"]]; // Return as array for consistency
         }
         // Call EXA_SEARCH
         return EXA_SEARCH(
           args.query,
           args.numResults, // EXA_SEARCH handles undefined/null internally
           args.searchType,
           args.prefix,
           args.suffix
         );

      default:
        return `❌ Unknown function: ${functionName}`;
    }
  } catch (e) {
    Logger.log(`Error in testExaFunction (${functionName}): ${e} - Args: ${JSON.stringify(args)}`);
    // Return a user-friendly script error message
    return `❌ Script Error during test: ${e.message}`;
  }
}

/**
 * Processes batch operations on selected cells containing Exa functions.
 * Supported operations: refresh (recalculate cells), clear (clear values)
 * 
 * @param {string} operation The operation to perform ('refresh' or 'clear')
 * @return {Object} Result object with success flag and message
 */
function processBatchOperation(operation) {
  try {
    // Get active spreadsheet and current selection
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const selection = ss.getActiveSheet().getActiveRange();
    
    if (!selection) {
      return { 
        success: false, 
        message: 'No cells selected. Please select cells containing Exa functions.'
      };
    }
    
    // Get formulas from the selection
    const formulas = selection.getFormulas();
    
    // Track counts for reporting
    let totalProcessed = 0;
    let exaCellsAffected = 0;
    
    // Process each cell in the selection
    for (let row = 0; row < formulas.length; row++) {
      for (let col = 0; col < formulas[row].length; col++) {
        const formula = formulas[row][col];
        totalProcessed++;
        
        // Check if this is an Exa formula
        if (formula && formula.toUpperCase().match(/^=EXA_/)) {
          const cell = selection.getCell(row + 1, col + 1);
          
          if (operation === 'refresh') {
            // Store the formula, clear the cell, then reapply to force recalculation
            const tempFormula = formula;
            // First set to a temporary value (empty string) then back to the original formula
            // This forces Google Sheets to recalculate the formula
            cell.setFormula('');
            SpreadsheetApp.flush();
            cell.setFormula(tempFormula);
            exaCellsAffected++;
          } 
          else if (operation === 'clear') {
            // Just clear the content, leaving the formula
            cell.clearContent();
            exaCellsAffected++;
          }
        }
      }
    }
    
    // Force recalculation when refreshing
    if (operation === 'refresh' && exaCellsAffected > 0) {
      SpreadsheetApp.flush();
    }
    
    // Generate result message
    let message;
    if (exaCellsAffected === 0) {
      message = `No Exa functions found in the ${totalProcessed} selected cell(s).`;
      return { success: false, message: message };
    } else {
      const actionText = operation === 'refresh' ? 'refreshed' : 'cleared';
      message = `Successfully ${actionText} ${exaCellsAffected} cell(s) with Exa functions.`;
      return { success: true, message: message };
    }
    
  } catch (e) {
    Logger.log(`Error in processBatchOperation (${operation}): ${e}`);
    return { 
      success: false, 
      message: `Operation failed: ${e.message}`
    };
  }
}
