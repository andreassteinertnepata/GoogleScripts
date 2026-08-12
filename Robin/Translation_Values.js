function translateValuesFromExternalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let valuesSheet;
  try {
    valuesSheet = SpreadsheetApp.openById(CONFIG.TRANS_SHEET_ID).getSheetByName(CONFIG.TAB_TRANS_VALUES);
  } catch (e) { return; }
  
  if (!valuesSheet) return;
  
  const valuesData = valuesSheet.getDataRange().getValues();
  const translationMap = {};
  
  for (let i = 0; i < valuesData.length; i++) {
    const headerName = String(valuesData[i][0] || "").trim();
    const originalValue = String(valuesData[i][1] || "").trim();
    const translatedValue = valuesData[i][2];
    
    if (headerName && originalValue) {
      if (!translationMap[headerName]) translationMap[headerName] = {};
      translationMap[headerName][originalValue] = translatedValue;
    }
  }
  
  ss.getSheets().forEach(function(targetSheet) {
    if (CONFIG.IGNORED_SHEETS.includes(targetSheet.getName())) return;
    
    const targetRange = targetSheet.getDataRange();
    const targetData = targetRange.getValues();
    if (targetData.length <= 1) return; 
    
    const headers = targetData[0];
    let hasChanges = false;
    
    for (let col = 0; col < headers.length; col++) {
      const currentHeader = String(headers[col] || "").trim();
      if (translationMap[currentHeader]) {
        for (let row = 1; row < targetData.length; row++) {
          const cellValue = String(targetData[row][col] || "").trim();
          if (translationMap[currentHeader][cellValue] !== undefined) {
            targetData[row][col] = translationMap[currentHeader][cellValue];
            hasChanges = true;
          }
        }
      }
    }
    if (hasChanges) targetRange.setValues(targetData);
  });
}

