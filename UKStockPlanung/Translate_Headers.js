function translateHeadersFromExternalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let headerSheet;
  try {
    headerSheet = SpreadsheetApp.openById(CONFIG.TRANS_SHEET_ID).getSheetByName(CONFIG.TAB_TRANS_HEADERS);
  } catch (e) { return; }
  
  if (!headerSheet) return;
  const headerData = headerSheet.getDataRange().getValues();
  const headerMap = {};
  
  for (let i = 0; i < headerData.length; i++) {
    const orig = String(headerData[i][0] || "").trim();
    const trans = String(headerData[i][1] || "").trim();
    if (orig && trans) headerMap[orig] = trans;
  }
  
  ss.getSheets().forEach(sheet => {
    if (CONFIG.IGNORED_SHEETS.includes(sheet.getName())) return;
    const maxCol = sheet.getLastColumn();
    if (maxCol === 0) return;
    
    const range = sheet.getRange(1, 1, 1, maxCol);
    const headers = range.getValues()[0];
    let hasChanges = false;
    
    for (let c = 0; c < headers.length; c++) {
      const txt = String(headers[c] || "").trim();
      if (headerMap[txt]) {
        headers[c] = headerMap[txt];
        hasChanges = true;
      }
    }
    if (hasChanges) range.setValues([headers]);
  });
}
