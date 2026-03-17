const FILTER_CONFIG = {
  // Sheet Names
  SOURCE_SHEET: 'Consolidated_Data_Source', // 
  DEST_SHEET: 'PENDING_MAINTENANCE',        // 
  
  // Target schema replacing highly specific business column names
  TARGET_COLUMNS: [
    'Item_Code', 'General_Description', 'Revision', 'Assigned_Engineer',
    'Local_Description', 'Classification_1', 'Origin_Country', 'Approval_Date',
    'Category_Group', 'Family_Group', 'Classification_2', 'Metric_Value',
    'Metric_Unit', 'Physical_File_Path', 'Record_UID'
  ],
  
  // Columns that are permitted to have empty/null values
  OPTIONAL_COLUMNS: ['Revision', 'Approval_Date'],
  
  // Specific string values that should be treated as invalid/empty
  ERROR_STATES: ['Not Found']
};

function filterAndSelectFinalData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(FILTER_CONFIG.SOURCE_SHEET);
    
    if (!sourceSheet) {
      throw new Error(`Source sheet "${FILTER_CONFIG.SOURCE_SHEET}" not found.`);
    }

    console.log(`Reading dataset from "${FILTER_CONFIG.SOURCE_SHEET}"...`);
    const allData = sourceSheet.getDataRange().getValues();
    const headers = allData.shift(); // Extract and remove header row

    // Map column names to indices for O(1) lookup
    const headerMap = createHeaderMap_(headers);

    // Validate that all required target columns exist in the source data
    const validTargetCols = FILTER_CONFIG.TARGET_COLUMNS.filter(col => col !== '');
    validTargetCols.forEach(colName => {
      if (headerMap[colName] === undefined) {
        throw new Error(`Schema mismatch: Required column "${colName}" is missing in the source data.`);
      }
    });
    
    // Determine which columns must be strictly checked for blanks
    const strictCheckCols = validTargetCols.filter(
      colName => !FILTER_CONFIG.OPTIONAL_COLUMNS.includes(colName)
    );
    
    // Map column names to their respective array indices
    const checkIndexes = strictCheckCols.map(name => headerMap[name]);
    const selectIndexes = validTargetCols.map(name => headerMap[name]);

    console.log('Filtering records...');
    const filteredData = [];

    // Process each row
    allData.forEach((row) => {
      // Validate row: Ensure all strictly required columns have valid values
      const isRowValid = checkIndexes.every(index => {
        const cellValue = row[index];
        return (
          cellValue !== null && 
          cellValue !== undefined && 
          cellValue !== '' && 
          !FILTER_CONFIG.ERROR_STATES.includes(cellValue)
        );
      });

      // If valid, project the row to only include the target columns
      if (isRowValid) {
        const selectedRowData = selectIndexes.map(index => row[index]);
        filteredData.push(selectedRowData);
      }
    });

    console.log(`Filtering complete. Identified ${filteredData.length} valid records.`);

    // --- Render to Destination Sheet ---
    let destSheet = ss.getSheetByName(FILTER_CONFIG.DEST_SHEET);
    if (!destSheet) {
      destSheet = ss.insertSheet(FILTER_CONFIG.DEST_SHEET);
    }
    destSheet.clear();

    // Construct final dataset with headers
    const finalOutput = [FILTER_CONFIG.TARGET_COLUMNS, ...filteredData];
    
    // Safety check before writing to sheet to avoid dimension mismatch errors
    if (finalOutput.length > 0 && finalOutput[0].length > 0) {
      const numRows = finalOutput.length;
      const numCols = FILTER_CONFIG.TARGET_COLUMNS.length; 

      destSheet.getRange(1, 1, numRows, numCols).setValues(finalOutput);
      destSheet.autoResizeColumns(1, numCols);
    }

    console.log(`Success! Data projected and written to "${FILTER_CONFIG.DEST_SHEET}".`);
    
    // Trigger downstream export process
    if (typeof saveSheetAsExcel === 'function') {
        saveSheetAsExcel(); 
    }

  } catch (e) {
    console.error(`Runtime Exception in filtering process: ${e.message} \nStack: ${e.stack}`);
  }
}

/**
 * Utility: Maps an array of headers to an object of { HeaderName: Index }
 * @param {Array<string>} headers - Array of header strings
 * @returns {Object} - Header index map
 */
function createHeaderMap_(headers) {
    return headers.reduce((map, header, index) => {
        map[header] = index;
        return map;
    }, {});
}
