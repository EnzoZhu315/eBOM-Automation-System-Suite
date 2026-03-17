function saveSheetAsExcel() {
  // --- 1. CONFIGURATION (MASKED) ---
  const EXPORT_CONFIG = {
    SOURCE_SHEET: "PENDING_MAINTENANCE", // 原: "待维护"
    FOLDER_ID: "YOUR_GOOGLE_DRIVE_FOLDER_ID", // 替换为占位符
    SUMMARY_SHEET: "Summary",
    LOG_CELL: "B14" // 对应原代码中的 (14, 2)
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(EXPORT_CONFIG.SOURCE_SHEET);

  if (!sourceSheet) {
    console.error(`Error: Sheet "${EXPORT_CONFIG.SOURCE_SHEET}" not found.`);
    return;
  }

  // Create a temporary spreadsheet for the export process
  // This avoids exporting all sheets in the current file
  const tempSpreadsheet = SpreadsheetApp.create(`TempExport_${ss.getName()}`);
  const tempId = tempSpreadsheet.getId();
  
  try {
    // Copy the source data to the temporary file
    sourceSheet.copyTo(tempSpreadsheet).setName(EXPORT_CONFIG.SOURCE_SHEET);

    // Remove the default initial sheet
    const defaultSheet = tempSpreadsheet.getSheetByName('Sheet1');
    if (defaultSheet) tempSpreadsheet.deleteSheet(defaultSheet);

    // Construct the export URL for Excel format
    const url = `https://docs.google.com/spreadsheets/d/${tempId}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    
    const options = {
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    // Fetch the file blob
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error("Failed to fetch Excel export from Google API.");
    }
    const blob = response.getBlob();

    // Access the destination folder
    const folder = DriveApp.getFolderById(EXPORT_CONFIG.FOLDER_ID);
    
    // Define the output filename
    const fileName = `${ss.getName()}_Export_Data.xlsx`;

    // --- Clean up existing versions with the same name ---
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setTrashed(true); 
      console.log(`Archived previous version: ${file.getName()}`);
    }

    // Save the new Excel file
    folder.createFile(blob.setName(fileName));

    console.log(` Success! "${fileName}" has been saved to the secure folder.`);

    // Log the success timestamp to the summary dashboard
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const summarySheet = ss.getSheetByName(EXPORT_CONFIG.SUMMARY_SHEET);
    if (summarySheet) {
      summarySheet.getRange(EXPORT_CONFIG.LOG_CELL).setValue(now);
    }

  } catch (e) {
    console.error(`Export failed: ${e.message}`);
  } finally {
    // CRITICAL: Always delete the temporary spreadsheet from Drive
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}
