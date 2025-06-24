# Exa for Google Sheets

## Description

This Google Apps Script integration brings the power of Exa API directly into Google Sheets, allowing you to search the web, query information, find similar content, and extract website content without leaving your spreadsheet.

## Features

### Custom Sheet Functions
* **EXA_ANSWER:** Query the Exa AI to get answers to questions based on web search results
* **EXA_SEARCH:** Search the web and retrieve a list of relevant URLs
* **EXA_CONTENTS:** Extract the text content from specific URLs
* **EXA_FINDSIMILAR:** Find URLs similar to a provided reference URL

### Enhanced Sidebar Interface
* **API Key Management:** Securely save your Exa API key
* **Documentation:** Built-in reference for all Exa functions and their parameters
* **Function Builder:** Interactive interface to build and test Exa functions
* **Batch Operations:** Refresh or clear multiple Exa function cells at once

## Setup & Installation

1.  **Open your Google Sheet:** Go to the Google Sheet where you want to use this script.
2.  **Open the Script Editor:** Click on "Extensions" > "Apps Script".
3.  **Copy the Code:**
    *   Copy the contents of `Code.gs` from this project and paste it into the `Code.gs` file in the Apps Script editor. Overwrite any existing template code.
    *   Create a new HTML file in the editor (File > New > HTML file). Name it `Sidebar.html` (ensure the name matches exactly, including capitalization).
    *   Copy the contents of the `Sidebar.html` file from this project and paste it into the newly created `Sidebar.html` file in the editor.
4.  **Save the Project:** Click the floppy disk icon (Save project) and give your script project a name (e.g., "Exa for Sheets").
5.  **Refresh Your Sheet:** Close the Apps Script editor tab and refresh your Google Sheet browser tab.
6.  **Authorize the Script:**
    *   After refreshing, a new custom menu item "Exa AI" should appear. Click on it, then select "Open Sidebar".
    *   A dialog box will pop up asking for authorization. Review the permissions requested (it will need access to external services, script properties, and the current spreadsheet) and click "Allow". You might need to go through an "Advanced" > "Go to (project name)" flow if Google warns it's an unverified app.
7.  **Set Your API Key:**
    *   Once authorized, the sidebar will open with the API Key tab active.
    *   Go to [https://exa.ai/](https://exa.ai/) to get your API key if you don't have one.
    *   Paste your Exa API key into the input field in the sidebar and click "Save API Key".
    *   You should see a success message "✅ API Key saved!".

## Using the Sidebar

The sidebar offers four main tabs:

### 1. API Key
* Set or update your Exa API key
* View status of key operations

### 2. Documentation
* Comprehensive documentation for all Exa functions
* Parameter descriptions and return value information
* Quick reference for function syntax

### 3. Function Builder
* Interactive interface to construct Exa functions
* Test functions before adding them to your sheet
* View results directly in the sidebar

### 4. Batch Operations
* Refresh selected cells containing Exa functions
* Clear results from selected Exa function cells
* Status updates for batch operations

## Using Exa Functions

### EXA_ANSWER
```
=EXA_ANSWER(prompt, [prefix], [suffix], [includeCitations])
```
Generates an AI answer based on search results for your query.

**Parameters:**
* `prompt` (required): The main question or prompt
* `prefix` (optional): Text to add before the main prompt
* `suffix` (optional): Text to add after the main prompt
* `includeCitations` (optional): If TRUE, includes source citations (Default: FALSE)

### EXA_SEARCH
```
=EXA_SEARCH(query, [numResults], [searchType], [prefix], [suffix])
```
Searches the web and returns a vertical list of URLs.

**Parameters:**
* `query` (required): The search query
* `numResults` (optional): Number of results to return (1-10, Default: 5)
* `searchType` (optional): "auto", "neural", or "keyword" (Default: "auto")
* `prefix` (optional): Text to add before the main query
* `suffix` (optional): Text to add after the main query

### EXA_CONTENTS
```
=EXA_CONTENTS(url)
```
Retrieves the text content from a specified URL.

**Parameters:**
* `url` (required): The full URL to extract content from (must start with http/https)

### EXA_FINDSIMILAR
```
=EXA_FINDSIMILAR(url, [numResults], [includeDomainsStr], [excludeDomainsStr], [includeTextStr], [excludeTextStr])
```
Finds URLs similar to the input URL.

**Parameters:**
* `url` (required): The reference URL to find similar content
* `numResults` (optional): Number of results to return (1-10, Default: 5)
* `includeDomainsStr` (optional): Comma-separated list of domains to include
* `excludeDomainsStr` (optional): Comma-separated list of domains to exclude
* `includeTextStr` (optional): Phrase that must appear in results
* `excludeTextStr` (optional): Phrase that must not appear in results

## Batch Operations

To use batch operations:

1. Select cells containing Exa functions in your sheet
2. Open the sidebar and navigate to the "Batch" tab
3. Choose an operation:
   * **Refresh Selected Cells:** Re-executes the Exa functions in selected cells
   * **Clear Selected Results:** Clears the results while preserving the formulas

## Notes

* Exa API requests count against your Exa usage quota
* For best performance, avoid excessive function calls in large sheets
* Functions will automatically refresh when their inputs change or when the sheet is reopened

---

## Privacy & Security

* Your Exa API key is stored securely using Google Apps Script's User Properties service
* The key is only accessible to your Google account
* No data is stored outside of your Google account and the Exa API