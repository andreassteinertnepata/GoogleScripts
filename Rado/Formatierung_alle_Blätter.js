function formatGoogleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colorMapSheet = ss.getSheetByName(CONFIG.SHEET_COLOR_MAPPING);
  let colorMap = {};
  
  if (colorMapSheet) {
    const mapData = colorMapSheet.getDataRange().getValues();
    for (let i = 1; i < mapData.length; i++) {
      if (mapData[i][0] && mapData[i][1]) {
        colorMap[mapData[i][0]] = mapData[i][1].toString().replace("#", "");
      }
    }
  }

  ss.getSheets().forEach(function(sheet) {
    const sheetName = sheet.getName();
    if (CONFIG.IGNORED_SHEETS.includes(sheetName)) return;
    
    sheet.setFrozenRows(1);
    const maxRow = sheet.getLastRow();
    const maxCol = sheet.getLastColumn();
    if (maxRow === 0 || maxCol === 0) return; 

    const fullRange = sheet.getRange(1, 1, maxRow, maxCol);
    if (sheet.getFilter()) sheet.getFilter().remove();
    fullRange.createFilter();
    
    if (colorMap[sheetName]) sheet.setTabColor(colorMap[sheetName]);
    
    fullRange.setBorder(true, true, true, true, true, true, CONFIG.FORMAT_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
    
    const headerRange = sheet.getRange(1, 1, 1, maxCol);
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(CONFIG.HEADER_FONT_SIZE);
    headerRange.setBackground(CONFIG.FORMAT_HEADER_BG); 
    headerRange.setFontColor(CONFIG.FORMAT_HEADER_TEXT); 
    headerRange.setHorizontalAlignment("left");
    
    if (maxRow > 1) {
      const dataRange = sheet.getRange(2, 1, maxRow - 1, maxCol);
      const backgrounds = [];
      for (let r = 2; r <= maxRow; r++) {
        const rowColor = (r % 2 === 0) ? CONFIG.FORMAT_ROW_EVEN : CONFIG.FORMAT_ROW_ODD;
        backgrounds.push(new Array(maxCol).fill(rowColor));
      }
      dataRange.setBackgrounds(backgrounds);
    }
    
    sheet.autoResizeColumns(1, maxCol);
    for (let col = 1; col <= maxCol; col++) {
      let targetWidth = sheet.getColumnWidth(col) + CONFIG.COL_WIDTH_BUFFER;
      if (targetWidth > CONFIG.MAX_COL_WIDTH) targetWidth = CONFIG.MAX_COL_WIDTH;
      sheet.setColumnWidth(col, targetWidth);
    }
  });
}