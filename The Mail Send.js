const MAIL_CONFIG = {
  //  SECURITY: Replace with your actual log sheet URL and recipient list
  LOG_SHEET_URL: "https://docs.google.com/spreadsheets/d/YOUR_LOG_SHEET_ID_HERE/edit",
  TARGET_SHEET_NAME: "Maintenance_Logs", // Masked from '汇总表'
  
  // List of email addresses separated by commas
  RECIPIENTS: "stakeholder_01@example.com,stakeholder_02@example.com", 
  
  // The index of the timestamp column (Q column is index 16)
  DATE_COL_IDX: 16,
  
  SYSTEM_TAG: "BOM_Auto_Sync"
};

/**
 * Sends a daily digest email containing today's maintenance logs.
 */
function sendDailySummaryEmail() {
  try {
    const ss = SpreadsheetApp.openByUrl(MAIL_CONFIG.LOG_SHEET_URL);
    const sheet = ss.getSheetByName(MAIL_CONFIG.TARGET_SHEET_NAME);
    
    if (!sheet) {
      console.error(`Error: Sheet "${MAIL_CONFIG.TARGET_SHEET_NAME}" not found.`);
      return;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const timeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    let todayRows = [];

    // --- Filter records matching today's date ---
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = row[MAIL_CONFIG.DATE_COL_IDX];
      
      let isToday = false;
      if (rowDate instanceof Date) {
        const formattedRowDate = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
        if (formattedRowDate === todayStr) isToday = true;
      } else if (rowDate && typeof rowDate === "string") {
        if (rowDate.includes(todayStr)) isToday = true;
      }

      if (isToday) todayRows.push(row);
    }

    // --- Skip if no new records found ---
    if (todayRows.length === 0) {
      console.log(`No new logs for ${todayStr}. Skipping email.`);
      return;
    }

    // --- Build HTML Email Content ---
    const subject = `[${MAIL_CONFIG.SYSTEM_TAG}] Daily Maintenance Report - ${todayStr}`;
    
    let htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <p>This is an automated summary of today's synchronization activities:</p>
        <table border="1" style="border-collapse:collapse; font-size:12px; width:100%; border:1px solid #ddd;">
          <thead style="background-color: #f8f9fa;">
            <tr>
    `;

    // Add Table Headers
    headers.forEach(h => {
      htmlBody += `<th style="padding:8px; border:1px solid #ddd; text-align:left;">${h}</th>`;
    });
    htmlBody += `</tr></thead><tbody>`;
    
    // Add Table Rows
    todayRows.forEach(row => {
      htmlBody += `<tr>`;
      row.forEach(cell => {
        let displayValue = cell;
        if (cell instanceof Date) {
          displayValue = Utilities.formatDate(cell, timeZone, "yyyy-MM-dd HH:mm:ss");
        }
        htmlBody += `<td style="padding:8px; border:1px solid #ddd;">${displayValue}</td>`;
      });
      htmlBody += `</tr>`;
    });

    htmlBody += `
          </tbody>
        </table>
        <p style="font-size:11px; color: #888; margin-top:20px;">
          <i>* This is an automated system message. Please do not reply directly.</i>
        </p>
      </div>
    `;

    // --- Send the Email ---
    MailApp.sendEmail({
      to: MAIL_CONFIG.RECIPIENTS,
      subject: subject,
      htmlBody: htmlBody
    });
    
    console.log(`Summary email dispatched to recipients. Records: ${todayRows.length}`);

  } catch (e) {
    console.error(`Email Module Error: ${e.message}`);
  }
}
