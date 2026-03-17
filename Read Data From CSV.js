const SYNC_CONFIG = {
  // SECURITY: Replace with your actual token or use PropertiesService.getUserProperties()
  API_TOKEN: "YOUR_SMARTSHEET_ACCESS_TOKEN", 
  
  //  FILE RESOURCES
  SOURCE_CSV_ID: "YOUR_DRIVE_FILE_ID_HERE",
  LOG_SPREADSHEET_URL: "https://docs.google.com/spreadsheets/d/YOUR_LOG_SHEET_ID/edit",

  //  TARGET SYSTEMS
  RESOURCES: {
    SYSTEM_A: {
      ID: "TARGET_SHEET_ID_ALPHA",
      MARKER_COL: "Done-4" // Status toggle column
    },
    SYSTEM_B: {
      ID: "TARGET_SHEET_ID_BETA",
      MARKER_COL: "Done-3"
    }
  },

  // LOGGING SCHEMA
  LOG_FIELDS: [
    "Description", "Version", "Engineer_In_Charge", "Local_Desc", 
    "Class_1", "Origin", "Appr_Date", "Category", "Family", "Class_2", 
    "Fill_Qty", "Fill_Unit"
  ]
};

/**
 * Main function to sync local data back to the remote API.
 */
function finalSyncToSmartsheetApi() {
  const token = SYNC_CONFIG.API_TOKEN;
  
  // Define time boundary (e.g., process records from the last 24 hours)
  const now = new Date();
  const syncWindowStart = new Date(now);
  syncWindowStart.setDate(now.getDate() - 1);
  syncWindowStart.setHours(0, 0, 0, 0); 
  
  // Fetch Column IDs dynamically to ensure robustness
  const colIdA = getSmartsheetColumnId(token, SYNC_CONFIG.RESOURCES.SYSTEM_A.ID, SYNC_CONFIG.RESOURCES.SYSTEM_A.MARKER_COL);
  const colIdB = getSmartsheetColumnId(token, SYNC_CONFIG.RESOURCES.SYSTEM_B.ID, SYNC_CONFIG.RESOURCES.SYSTEM_B.MARKER_COL);
  
  if (!colIdA || !colIdB) {
    console.error("Critical: Could not resolve Smartsheet Column IDs. Aborting.");
    return;
  }

  // Get log sheet reference
  let logSheet;
  try {
    logSheet = SpreadsheetApp.openByUrl(SYNC_CONFIG.LOG_SPREADSHEET_URL).getSheets()[0];
  } catch(e) { /* Silent fail if log sheet inaccessible */ }

  // Load and Parse Source CSV
  let csvData;
  try {
    const file = DriveApp.getFileById(SYNC_CONFIG.SOURCE_CSV_ID);
    const csvContent = file.getBlob().getDataAsString("UTF-8");
    csvData = Utilities.parseCsv(csvContent);
  } catch (e) {
    console.error("Failed to read source CSV file.");
    return;
  }

  const headers = csvData[0];
  const fieldIndices = SYNC_CONFIG.LOG_FIELDS.map(name => headers.indexOf(name));
  const timeIdx = headers.indexOf("Maintenance_Time");
  const idIdx = headers.indexOf("row_id"); 
  const matIdx = headers.indexOf("Material Code");

  // Iterate through records (skipping header)
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    
    // Clean and validate Row ID (handles scientific notation from exports)
    let rowId = normalizeRowId(row[idIdx]);
    const matCode = String(row[matIdx] || "").trim();
    const timeStr = row[timeIdx];

    if (!rowId || rowId === "0" || !timeStr) continue;

    // Filter by timestamp to avoid re-processing old data
    const rowTime = new Date(timeStr);
    if (isNaN(rowTime.getTime()) || rowTime < syncWindowStart) continue; 

    // Determine target sheet based on business logic (e.g., prefix 'P' for PDI system)
    const isSpecialType = matCode.toUpperCase().startsWith("P");
    const targetSheetId = isSpecialType ? SYNC_CONFIG.RESOURCES.SYSTEM_A.ID : SYNC_CONFIG.RESOURCES.SYSTEM_B.ID;
    const targetColId = isSpecialType ? colIdA : colIdB;

    // Execute API Update
    const updateSuccess = updateSmartsheetRow(token, targetSheetId, rowId, targetColId);

    // Audit Log Entry
    if (updateSuccess && logSheet) {
      const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
      const mappedValues = fieldIndices.map(idx => (idx !== -1 ? row[idx] : ""));
      const auditRow = [rowId, timestamp, matCode, ...mappedValues];
      logSheet.appendRow(auditRow); 
    }
  }
}

/** 🛠️ UTILITY FUNCTIONS **/

/**
 * Fixes row IDs that might be corrupted by scientific notation during CSV conversion.
 */
function normalizeRowId(val) {
  if (!val) return "";
  let str = String(val).trim();
  if (str.toLowerCase().includes('e')) {
    return Number(str).toLocaleString('fullwide', {useGrouping:false});
  }
  return str;
}

/**
 * Updates a single cell in Smartsheet (usually to toggle a status checkbox).
 */
function updateSmartsheetRow(token, sheetId, rowId, colId) {
  const url = `https://api.smartsheet.com/2.0/sheets/${sheetId}/rows`;
  const payload = [{ "id": rowId, "cells": [{ "columnId": colId, "value": true }] }];
  const options = {
    "method": "put",
    "headers": { 
      "Authorization": "Bearer " + token, 
      "Content-Type": "application/json" 
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch(e) { return false; }
}

/**
 * Fetches the specific Column ID from Smartsheet based on its Display Name.
 */
function getSmartsheetColumnId(token, sheetId, colName) {
  try {
    const url = `https://api.smartsheet.com/2.0/sheets/${sheetId}?include=columns`;
    const res = UrlFetchApp.fetch(url, { "headers": { "Authorization": "Bearer " + token } });
    const data = JSON.parse(res.getContentText());
    const match = data.columns.find(c => c.title === colName);
    return match ? match.id : null;
  } catch (e) { return null; }
}
