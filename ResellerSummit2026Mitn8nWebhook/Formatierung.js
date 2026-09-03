function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const lastColumn = dataRange.getLastColumn();
  const lastRow = dataRange.getLastRow();

  if (lastRow === 0 || lastColumn === 0) return;

  // A. Header-Zeile formatieren (Zeile 1)
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  headerRange
    .setBackground('#46bdc6')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // B. Datenbereich formatieren (ab Zeile 2)
  if (lastRow > 1) {
    const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    bodyRange
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // C. Auto-Filter setzen (entfernt vorherige Filter, um Fehler zu vermeiden)
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  dataRange.createFilter();

  // D. Spaltenbreiten automatisch anpassen
  for (let col = 1; col <= lastColumn; col++) {
    sheet.autoResizeColumn(col);
  }

  
}