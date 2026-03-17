// --- 1. GLOBAL CONFIGURATION ---
const CONFIG = {
  // API & Security
  AUTH_TOKEN: "YOUR_SMARTSHEET_ACCESS_TOKEN", // ⚠️ TOKENS
  
  // External Resource IDs
  SMARTSHEET_IDS: {
    PDI_SOURCE: 'YOUR_PDI_SHEET_ID_HERE',
    COPACK_SOURCE: 'YOUR_COPACK_SHEET_ID_HERE'
  },
  DRIVE_FOLDER_ID: "YOUR_GOOGLE_DRIVE_FOLDER_ID",

  // Sheet Names in Google Spreadsheet
  SHEETS: {
    DESTINATION: 'Consolidated_Report',
    FILE_LOOKUP: 'sys_file_mapping',
    DESC_MAPPING: 'sys_material_desc',
    TEMPLATE: 'sys_process_template',
    SUMMARY: 'Summary'
  },

  // Physical Path Settings (for local drive mapping)
  PATH_CONFIG: {
    PREFIX: 'S:\\Corporate_Shared_Drives\\Project_BOM_Root',
    SUB_FOLDER_REPACK: 'Site_Repack_Guidelines',
    SUB_FOLDER_ARTWORK: 'Site_BOM_Artwork'
  },

  COLUMN_LIMIT: 20
};

/**
 * Main execution engine.
 */
function runAdvancedDataCombination() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Step 1: Update Internal File Index ---
    console.log("Refreshing file index from Drive...");
    updateLocalFileIndex(); 

    // --- Step 2 & 3: Fetch and process Source Data ---
    // Note: Assuming get_copackebom_smartsheet() etc. are defined elsewhere using CONFIG.SMARTSHEET_IDS
    console.log("Fetching remote data streams...");
    const transformedCopackData = processCopackTable1(); 
    const transformedPdiData = processPdiTable1();
    
    // Internal data maintenance tasks
    syncExternalData(); 

    // --- Step 4: Merge Data Streams ---
    console.log("Merging data streams...");
    const header = [
      'Status', 'Material Code', 'Description', 'Version', 
      'Lead Engineer', 'Change Category', 'Ref Doc', 'Type', 'uid'
    ];
    let combinedData = [header, ...transformedCopackData, ...transformedPdiData];

    if (combinedData.length <= 1) {
      console.warn('No active records found. Terminating process.');
      return;
    }
    
    // --- Step 5: Enrichment - Local Descriptions ---
    console.log("Enriching data with localized descriptions...");
    combinedData = addLocalizedDescriptions(ss, combinedData);

    // --- Step 6: Enrichment - Template Mapping ---
    console.log("Applying process templates...");
    combinedData = applyTemplateJoin(ss, combinedData);
    
    // --- Step 7: Enrichment - File Path Resolution ---
    console.log("Resolving physical file paths...");
    combinedData = resolveFilePaths(ss, combinedData);

    // --- Step 8: Final Output & Formatting ---
    renderFinalOutput(ss, combinedData);
    
    // Log completion
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const summarySheet = ss.getSheetByName(CONFIG.SHEETS.SUMMARY);
    if (summarySheet) summarySheet.getRange(10, 2).setValue(timestamp);

    // Optional Export
    finalizeAndExport();

    console.log('✅ Process completed successfully.');
  } catch (e) {
    console.error(`Process Error: ${e.message} \nStack: ${e.stack}`);
  }
}

/**
 * Resolves descriptions using a Map for O(n) performance.
 */
function addLocalizedDescriptions(ss, data) {
    const descSheet = ss.getSheetByName(CONFIG.SHEETS.DESC_MAPPING);
    if (!descSheet) throw new Error("Description mapping sheet missing.");
    
    const descData = descSheet.getRange(2, 1, descSheet.getLastRow() - 1, 2).getValues();
    const descMap = new Map(descData.map(row => [String(row[0]).trim(), row[1]]));

    const headers = data[0];
    const codeIdx = headers.indexOf('Material Code');
    
    headers.push('localized_desc');

    for (let i = 1; i < data.length; i++) {
        const code = String(data[i][codeIdx]).trim();
        data[i].push(descMap.get(code) || 'N/A');
    }
    return data;
}

/**
 * Maps physical paths based on Material Type.
 */
function resolveFilePaths(ss, data) {
    const fileIndexSheet = ss.getSheetByName(CONFIG.SHEETS.FILE_LOOKUP);
    if (!fileIndexSheet) throw new Error("File index sheet missing.");
    
    const fileData = fileIndexSheet.getDataRange().getValues();
    const fileHeaders = fileData.shift();
    const fileHeaderMap = createHeaderMap_(fileHeaders);

    const dataHeaders = data[0];
    const dataHeaderMap = createHeaderMap_(dataHeaders);

    data[0].push('Physical_File_Path'); 

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const mCode = row[dataHeaderMap['Material Code']];
        const mType = row[dataHeaderMap['Type']];
        const version = row[dataHeaderMap['Version']];
        
        let path = 'NOT_FOUND';
        if (mType === 'SKU') {
            path = searchPath_(mCode, version, fileData, fileHeaderMap, CONFIG.PATH_CONFIG.SUB_FOLDER_REPACK);
        } else {
            path = searchPath_(mCode, null, fileData, fileHeaderMap, CONFIG.PATH_CONFIG.SUB_FOLDER_ARTWORK);
        }
        row.push(path);
    }
    return data;
}

/**
 * Generic path search logic.
 */
function searchPath_(identifier, version, fileEntries, headerMap, subFolder) {
  if (!identifier) return '';
  const idUpper = String(identifier).toUpperCase();
  const verUpper = version ? String(version).toUpperCase() : null;
  
  const matches = fileEntries.filter(row => {
    const nameMatch = String(row[headerMap['File_Name']]).toUpperCase().includes(idUpper);
    const verMatch = verUpper ? String(row[headerMap['File_Name']]).toUpperCase().includes(verUpper) : true;
    const folderMatch = String(row[headerMap['Folder_Path']]).includes(subFolder);
    return nameMatch && verMatch && folderMatch;
  });

  if (matches.length === 0) return 'Path Error: File not in index';

  // Return the latest version found
  matches.sort((a, b) => new Date(b[headerMap['Last_Updated']]) - new Date(a[headerMap['Last_Updated']]));
  
  return [CONFIG.PATH_CONFIG.PREFIX, matches[0][headerMap['Folder_Path']], matches[0][headerMap['File_Name']]].join('\\');
}

/**
 * Moves uid/row_id to the last column and writes to sheet.
 */
function renderFinalOutput(ss, combinedData) {
    let destSheet = ss.getSheetByName(CONFIG.SHEETS.DESTINATION) || ss.insertSheet(CONFIG.SHEETS.DESTINATION);
    destSheet.clear();

    const header = combinedData[0];
    const uidIdx = header.indexOf('uid');

    if (uidIdx !== -1) {
      combinedData = combinedData.map(row => {
        const uidValue = row.splice(uidIdx, 1)[0]; 
        row.push(uidValue);
        return row;
      });
    }

    destSheet.getRange(1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
    destSheet.autoResizeColumns(1, combinedData[0].length);
}

// Internal Helper
function createHeaderMap_(headerRow) {
  return headerRow.reduce((map, header, i) => {
    map[header] = i;
    return map;
  }, {});
}
