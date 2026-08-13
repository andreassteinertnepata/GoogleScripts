/**
 * Hochoptimierte Formatierungsfunktion
 * Reduziert die Ausführungszeit extrem durch Minimierung von API-Calls.
 */
function formatGoogleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Farbmuster (Mapping) in einem einzigen Aufruf laden
  const colorMapSheet = ss.getSheetByName(CONFIG.SHEET_COLOR_MAPPING);
  const colorMap = {};
  
  if (colorMapSheet) {
    const mapData = colorMapSheet.getDataRange().getValues();
    for (let i = 1; i < mapData.length; i++) {
      if (mapData[i][0] && mapData[i][1]) {
        colorMap[mapData[i][0]] = mapData[i][1].toString().replace("#", "");
      }
    }
  }

  const sheets = ss.getSheets();
  const ignoredSheets = new Set(CONFIG.IGNORED_SHEETS || []);

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (ignoredSheets.has(sheetName)) return;

    const maxRow = sheet.getLastRow();
    const maxCol = sheet.getLastColumn();
    if (maxRow === 0 || maxCol === 0) return;

    // 1. Zeile fixieren
    sheet.setFrozenRows(1);

    // Filter bereinigen & neu anlegen
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    
    const fullRange = sheet.getRange(1, 1, maxRow, maxCol);
    fullRange.createFilter();

    // Tab-Farbe setzen
    if (colorMap[sheetName]) {
      sheet.setTabColor(colorMap[sheetName]);
    }

    // Rahmen ziehen
    fullRange.setBorder(
      true, true, true, true, true, true, 
      CONFIG.FORMAT_BORDER_COLOR || "#000000", 
      SpreadsheetApp.BorderStyle.SOLID
    );

    // Header stylen
    const headerRange = sheet.getRange(1, 1, 1, maxCol);
    headerRange.setFontWeight("bold")
               .setFontSize(CONFIG.HEADER_FONT_SIZE || 11)
               .setBackground(CONFIG.FORMAT_HEADER_BG || "#3bb7c4")
               .setFontColor(CONFIG.FORMAT_HEADER_TEXT || "#000000")
               .setHorizontalAlignment("left");

    // Zebra-Streifen (Datenbereich) effizient erstellen
    if (maxRow > 1) {
      const dataRange = sheet.getRange(2, 1, maxRow - 1, maxCol);
      const evenColor = CONFIG.FORMAT_ROW_EVEN || "#f4f8f9";
      const oddColor = CONFIG.FORMAT_ROW_ODD || "#ffffff";
      
      const backgrounds = new Array(maxRow - 1);
      for (let r = 0; r < maxRow - 1; r++) {
        // r = 0 entspricht Zeile 2 (Gerade)
        const rowColor = (r % 2 === 0) ? evenColor : oddColor;
        backgrounds[r] = new Array(maxCol).fill(rowColor);
      }
      dataRange.setBackgrounds(backgrounds);
    }

    // SPALTENBREITEN OPTIMIEREN (Massiver Geschwindigkeitsgewinn)
    // 1. Auto-Resize für das gesamte Sheet in einem Rutsch
    sheet.autoResizeColumns(1, maxCol);

    // 2. Breiten in einem Durchlauf anpassen
    const buffer = CONFIG.COL_WIDTH_BUFFER || 20;
    const maxWidth = CONFIG.MAX_COL_WIDTH || 250;

    for (let col = 1; col <= maxCol; col++) {
      const currentWidth = sheet.getColumnWidth(col);
      const targetWidth = Math.min(currentWidth + buffer, maxWidth);
      if (targetWidth !== currentWidth) {
        sheet.setColumnWidth(col, targetWidth);
      }
    }
  });


}