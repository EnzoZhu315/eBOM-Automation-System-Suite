const REPORT_CONFIG = {
  //  SECURITY: Replace with your actual IDs and specific recipients
  SPREADSHEET_URL: "https://docs.google.com/spreadsheets/d/YOUR_LOG_SHEET_ID_HERE/edit",
  
  // Professional placeholder for internal distribution list
  RECIPIENTS: "key_stakeholder@example.com,lead_engineer@example.com,automation_admin@example.com",
  
  // Sheet Names
  SHEET_SUCCESS: "Maintenance_Logs", // Original: 汇总表
  SHEET_PENDING: "Combined_Output_Advanced",
  
  // Column Index Configurations
  COL_LOG_DATE: 16,        // Q Column
  COL_VALIDATION_H: 7,     // H Column
  COL_VALIDATION_I: 8,     // I Column
  COL_VALIDATION_R: 17,    // R Column
  
  // Visual Styles
  COLOR_SUCCESS: "#27AE60", // Green
  COLOR_WARNING: "#E74C3C"  // Red
};

function sendDailyIntegratedReport() {
  const ss = SpreadsheetApp.openByUrl(REPORT_CONFIG.SPREADSHEET_URL);
  const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  
  // --- PART 1: Fetch Processed Items (Success) ---
  const sheetMaintained = ss.getSheetByName(REPORT_CONFIG.SHEET_SUCCESS);
  let maintainedRows = [];
  let headersMaintained = [];
  
  if (sheetMaintained) {
    const data = sheetMaintained.getDataRange().getValues();
    headersMaintained = data[0];
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][REPORT_CONFIG.COL_LOG_DATE];
      if (isDateMatchingToday_(rowDate, todayStr)) {
        maintainedRows.push(data[i]);
      }
    }
  }

  // --- PART 2: Fetch Pending/Exception Items ---
  const sheetUnmaintained = ss.getSheetByName(REPORT_CONFIG.SHEET_PENDING);
  let unmaintainedRows = [];
  let headersUnmaintained = [];
  
  if (sheetUnmaintained) {
    const data = sheetUnmaintained.getDataRange().getValues();
    headersUnmaintained = data[0];
    for (let i = 1; i < data.length; i++) {
      const valH = String(data[i][REPORT_CONFIG.COL_VALIDATION_H]).trim();
      const valI = String(data[i][REPORT_CONFIG.COL_VALIDATION_I]).trim();
      const valR = String(data[i][REPORT_CONFIG.COL_VALIDATION_R]).trim();
      
      // Business Logic: If crucial fields are empty or file is missing
      if (valH === "" || valI === "" || valR === "Not Found") {
        unmaintainedRows.push(data[i]);
      }
    }
  }

  // --- PART 3: Verification ---
  if (maintainedRows.length === 0 && unmaintainedRows.length === 0) {
    console.log("No new updates or exceptions today. Skipping report.");
    return;
  }

  // --- PART 4: Build HTML Email Body ---
  const emailTitle = `EBOM Automation Maintenance Report (${todayStr})`;
  let htmlBody = `<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">`;
  htmlBody += `<h2>System Maintenance Summary - ${todayStr}</h2>`;

  // Table 1: Successfully Processed
  htmlBody += `<h3> Successfully Processed Items (${maintainedRows.length})</h3>`;
  if (maintainedRows.length > 0) {
    htmlBody += createTableHtml_(REPORT_CONFIG.COLOR_SUCCESS, headersMaintained, maintainedRows);
  } else {
    htmlBody += `<p style="color: #666;">No maintenance activity recorded today.</p>`;
  }

  htmlBody += `<br><hr style="border: 0; border-top: 1px solid #eee;"><br>`;

  // Table 2: Pending Exceptions
  htmlBody += `<h3> Pending Actions Required (${unmaintainedRows.length})</h3>`;
  htmlBody += `<p>Please review the following items for missing descriptions or source file mapping errors:</p>`;
  if (unmaintainedRows.length > 0) {
    htmlBody += createTableHtml_(REPORT_CONFIG.COLOR_WARNING, headersUnmaintained, unmaintainedRows);
  } else {
    htmlBody += `<p style="color: #27AE60;">Zero exceptions detected in current queue.</p>`;
  }

  htmlBody += `<br><p style="font-size: 11px; color: #888;"><i>* This is an automated system notification. Please do not reply.</i></p></div>`;

  // --- PART 5: Dispatch ---
  MailApp.sendEmail({
    to: REPORT_CONFIG.RECIPIENTS,
    subject: emailTitle,
    htmlBody: htmlBody
  });
  
  console.log(`Report sent. Processed: ${maintainedRows.length}, Exceptions: ${unmaintainedRows.length}`);
}

/**
 * Helper: Create HTML Table with Zebra Striping
 */
function createTableHtml_(headerColor, headers, rows) {
  let tableHtml = `<table border="1" style="border-collapse:collapse; font-size:11px; text-align:left; width:100%; border: 1px solid #ddd;">`;
  
  tableHtml += `<tr style="background-color: ${headerColor}; color: white; font-weight: bold;">`;
  headers.forEach(h => tableHtml += `<th style="padding:8px; border:1px solid #ddd;">${h}</th>`);
  tableHtml += `</tr>`;
  
  rows.forEach((row, index) => {
    const bgColor = index % 2 === 0 ? "#ffffff" : "#fcfcfc";
    tableHtml += `<tr style="background-color: ${bgColor};">`;
    row.forEach(cell => {
      let val = (cell instanceof Date) ? Utilities.formatDate(cell, "GMT+8", "yyyy-MM-dd HH:mm:ss") : cell;
      tableHtml += `<td style="padding:8px; border:1px solid #ddd;">${val}</td>`;
    });
    tableHtml += `</tr>`;
  });
  
  tableHtml += `</table>`;
  return tableHtml;
}

/**
 * Helper: Match Date String
 */
function isDateMatchingToday_(rowDate, todayStr) {
  if (rowDate instanceof Date) {
    return Utilities.formatDate(rowDate, "GMT+8", "yyyy-MM-dd") === todayStr;
  } else if (rowDate && typeof rowDate === "string") {
    return rowDate.includes(todayStr);
  }
  return false;
}
