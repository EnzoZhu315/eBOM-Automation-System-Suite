function runDailyAutomationWorkflow() {
  console.log(" [Stage 1/3] Initiating Smartsheet API Sync (finalSyncToSmartsheetApi)...");
  try {
    // 1. Execute synchronization logic (Filtering & API push)
    finalSyncToSmartsheetApi();
    console.log(" Stage 1 Complete: Remote data synchronized.");
  } catch (e) {
    console.error(" Stage 1 Failed: Sync interrupted. Process terminated to prevent data inconsistency. Error: " + e.message);
    return; // Stop execution if critical sync fails
  }

  console.log(" [Stage 2/3] Initiating Data Aggregation (updateSummaryTableOverwrite)...");
  try {
    // 2. Execute local spreadsheet aggregation and cleanup
    updateSummaryTableOverwrite();
    console.log(" Stage 2 Complete: Local summary tables refreshed.");
  } catch (e) {
    console.error(" Stage 2 Failed: Aggregation error. Process terminated. Error: " + e.message);
    return; 
  }

  // --- SAFETY BUFFER ---
  // Pause for 3 seconds to ensure Google Sheets background calculation & IO buffer are cleared
  console.log("⏳ Waiting for Google Sheets background engine to settle...");
  Utilities.sleep(3000);

  console.log(" [Stage 3/3] Generating Integrated Maintenance Report (sendDailyIntegratedReport)...");
  try {
    // 3. Dispatch the final professional HTML report to stakeholders
    sendDailyIntegratedReport();
    console.log(" Stage 3 Complete: Integrated report dispatched. Workflow successful.");
  } catch (e) {
    console.error(" Stage 3 Failed: Email reporting failed. Error: " + e.message);
  }
}
