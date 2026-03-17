const DOMO_CONFIG = {
  // Google Sheets Target
  TARGET_WORKBOOK_ID: 'YOUR_SPREADSHEET_ID_HERE', // Masked Spreadsheet ID
  TARGET_SHEET_NAME: 'SYS_Item_Descriptions',     // Masked from 
  
  // Domo API Details
  DATASET_ID: 'YOUR_DOMO_DATASET_ID_HERE',        // Masked 
  API_BASE_URL: 'https://api.domo.com/v1/datasets/query/execute/',
  
  // Query Filter Parameters (Masked business logic)
  FILTERS: {
    LOCALE_KEY: 'DEFAULT_LOCALE', // Masked from
    SYSTEM_ID: 'ERP_PRD',         // Masked from 
    SITE_CODE: 'SITE_001'         // Masked from 
  }
};

/**
 * Main function to execute the data sync process.
 */
function syncDescriptionData() {
  // Define standard headers
  const header = [['ITEM_ID', 'ITEM_DESCRIPTION']];
  
  // Fetch data from external API
  const externalData = executeDataLakeQuery();
  
  if (!externalData || externalData.length === 0) {
    console.warn("No data returned from Domo query. Sync aborted.");
    return;
  }
  
  const finalOutput = header.concat(externalData);
  
  try {
    const targetWorkbook = SpreadsheetApp.openById(DOMO_CONFIG.TARGET_WORKBOOK_ID);
    const targetSheet = targetWorkbook.getSheetByName(DOMO_CONFIG.TARGET_SHEET_NAME);
    
    if (!targetSheet) {
      throw new Error(`Target sheet "${DOMO_CONFIG.TARGET_SHEET_NAME}" not found.`);
    }

    targetSheet.clear();
    targetSheet.getRange(1, 1, finalOutput.length, finalOutput[0].length).setValues(finalOutput);
    console.log(`Successfully synced ${externalData.length} records.`);
    
  } catch (error) {
    console.error(`Error writing to Google Sheets: ${error.message}`);
  }
}

/**
 * Executes a SQL query against the Domo dataset via REST API.
 * @returns {Array<Array<string>>} 2D array of query results.
 */
function executeDataLakeQuery() {
  const accessToken = getAccessToken(); // Assumes getAccessToken() is defined elsewhere
  const url = DOMO_CONFIG.API_BASE_URL + DOMO_CONFIG.DATASET_ID;

  // Masked SQL Query using template literals
  // Replaced highly specific SAP/ERP column names with generic equivalents
  const sqlQuery = `
    SELECT DISTINCT 
      "ITEM_ID", 
      "ITEM_DESCRIPTION" 
    FROM dataset_table 
    WHERE "LOCALE_KEY" = '${DOMO_CONFIG.FILTERS.LOCALE_KEY}' 
      AND "SYSTEM_ID" = '${DOMO_CONFIG.FILTERS.SYSTEM_ID}' 
      AND "SITE_CODE" = '${DOMO_CONFIG.FILTERS.SITE_CODE}' 
  `;

  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + accessToken
    },
    'payload': JSON.stringify({
      'sql': sqlQuery
    }),
    'muteHttpExceptions': true 
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(response.getContentText());
      
      // Return the rows array directly so it can be concatenated cleanly
      return data.rows || [];

    } else {
      console.error(`Domo API Error: Received HTTP ${responseCode}`);
      console.error(`Response Body: ${response.getContentText()}`);
      return [];
    }
  } catch (error) {
    console.error(`Network Exception during Domo API call: ${error.message}`);
    return [];
  }
}
