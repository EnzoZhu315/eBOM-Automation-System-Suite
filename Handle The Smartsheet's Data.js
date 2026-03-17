/**
 * Configuration object to store business logic constants.
 * In a real-world scenario, these could be fetched from ScriptProperties.
 */
const APP_CONFIG = {
  SHEETS: {
    BOM: 'CoPack_eBOM',
    PDI: 'pdi_initial'
  },
  TARGET_PLANT: 'SITE_CODE_001', // Masked from 'CN12'
  STATUS: {
    NEW: 'STATUS_A',
    CHANGE: 'STATUS_B',
    IGNORE: 'No change'
  },
  PROCUREMENT: {
    OUTSOURCE: 'TYPE_EXT',
    CATEGORY_P: 'CAT_P'
  }
};

/**
 * Processes Bill of Materials (BOM) data with privacy-first filtering.
 */
function processBomData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.BOM);
  
  if (!sheet) return [];

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  const headers = allData[0];
  const dataRows = allData.slice(1);

  // Map indices using generic internal keys
  const getIdx = (name) => headers.indexOf(name);
  
  const col = {
    status: getIdx('Status'),
    id: getIdx('SKU'),
    desc: getIdx('Description'),
    version: getIdx('Version'),
    owner: getIdx('PKG Owner'),
    changeLog: getIdx('Key Change Dropdown'), 
    refDoc: getIdx('ECN if any'),
    gate1: getIdx('Done-1'),
    location: getIdx('Co-pack center'),
    timestamp: getIdx('Initiate Date'),
    gate3: getIdx('Done-3'),
    row_uid: getIdx('row_id')
  };

  // Validate if critical columns exist
  if (Object.values(col).includes(-1)) {
    Logger.log('Required headers missing.');
    return [];
  }

  return dataRows.filter(row => {
    const isNew = (
      row[col.status] === APP_CONFIG.STATUS.NEW &&
      row[col.gate3] === '' &&
      row[col.location] === APP_CONFIG.TARGET_PLANT &&
      row[col.version] !== APP_CONFIG.STATUS.IGNORE
    );

    const isChange = (
      row[col.status] === APP_CONFIG.STATUS.CHANGE &&
      row[col.gate1] !== '' &&
      row[col.location] === APP_CONFIG.TARGET_PLANT &&
      row[col.version] !== APP_CONFIG.STATUS.IGNORE &&
      row[col.gate3] === '' 
    );
    
    return isNew || isChange;
  }).map(row => [
    row[col.status],
    row[col.id],
    row[col.desc],
    row[col.version],
    row[col.owner],
    row[col.changeLog],
    row[col.refDoc],
    'DATA_TYPE_LABEL', 
    row[col.row_uid]
  ]);
}

/**
 * Processes Procurement and Delivery Index (PDI) data.
 */
function processPdiData() {
  const pdiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.PDI);
  if (!pdiSheet) return [];

  const allData = pdiSheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  const headers = allData[0];
  const dataRows = allData.slice(1);

  const getIdx = (name) => headers.indexOf(name);
  
  const col = {
    priority: getIdx('Standard/Urgent'),
    pType: getIdx('Procurement Type'),
    dept: getIdx('Plant/Department'),
    cat: getIdx('Code Category'),
    gate4: getIdx('Done-4'),
    mCode: getIdx('Material Code'),
    desc: getIdx('Description'),
    user: getIdx('Initial Person'),
    mType: getIdx('Material Type'),
    row_uid: getIdx('row_id')
  };

  if (Object.values(col).includes(-1)) return [];
   
  return dataRows.filter(row => {
    const pVal = String(row[col.priority]);
    return (
      pVal !== '' && 
      pVal !== 'CANCELLED' &&
      row[col.pType] === APP_CONFIG.PROCUREMENT.OUTSOURCE && 
      String(row[col.dept]).includes(APP_CONFIG.TARGET_PLANT) &&
      row[col.cat] === APP_CONFIG.PROCUREMENT.CATEGORY_P && 
      row[col.gate4] == ''
    );
  }).map(row => [
    '', // Placeholder
    row[col.mCode],
    row[col.desc],
    '', // Placeholder
    row[col.user],
    '', // Placeholder
    '', // Placeholder
    row[col.mType],
    row[col.row_uid]
  ]);
}
