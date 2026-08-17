function Mobile_importierenUndAktualisieren() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MOBILE);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MOBILE);
  }

  // Ordner abrufen mit Wiederholungsversuchen
  var folder = null;
  for (var i = 0; i < 3; i++) {
    try {
      folder = DriveApp.getFolderById(CONFIG.MOBILE_FOLDER_ID);
      break;
    } catch (e) {
      Utilities.sleep(1000);
    }
  }

  if (!folder) {
    throw new Error("Ordner konnte nicht geladen werden. Bitte prüfen Sie Ihre Zugriffsrechte.");
  }

  // Neueste Datei finden
  var files = folder.getFiles();
  var latestFile = null;
  var latestDate = new Date(0);

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(CONFIG.MOBILE_FILE_PREFIX) === 0) {
      if (file.getLastUpdated() > latestDate) {
        latestDate = file.getLastUpdated();
        latestFile = file;
      }
    }
  }

  if (!latestFile) {
    Logger.log('Keine Datei gefunden, die mit "' + CONFIG.MOBILE_FILE_PREFIX + '" beginnt.');
    return;
  }

  Logger.log('Neueste Datei gefunden: ' + latestFile.getName());

  // Datum der letzten Änderung formatieren
  var fileLastModified = latestFile.getLastUpdated();
  var formattedDate = Utilities.formatDate(fileLastModified, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm:ss");

  // Daten auslesen und CSV parsen
  var fileData = latestFile.getBlob().getDataAsString();
  var delimiter = fileData.indexOf(';') !== -1 ? ';' : ',';
  var csvData = Utilities.parseCsv(fileData, delimiter);

  if (csvData.length === 0) {
    Logger.log('Die Datei ist leer.');
    return;
  }

  // Indizes der gewünschten Spalten ermitteln
  var headers = csvData[0];
  var colIndices = [];
  
  for (var c = 0; c < CONFIG.MOBILE_DESIRED_COLUMNS.length; c++) {
    var targetCol = CONFIG.MOBILE_DESIRED_COLUMNS[c];
    var foundIndex = headers.indexOf(targetCol);
    
    if (foundIndex === -1) {
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].trim() === targetCol) {
          foundIndex = h;
          break;
        }
      }
    }
    
    colIndices.push(foundIndex);
  }

  var skuIndexInDesired = CONFIG.MOBILE_DESIRED_COLUMNS.indexOf("SKU");

  // Spalten extrahieren
  var filteredData = [];
  filteredData.push(CONFIG.MOBILE_DESIRED_COLUMNS);

  for (var r = 1; r < csvData.length; r++) {
    var row = csvData[r];
    var newRow = [];
    
    for (var k = 0; k < colIndices.length; k++) {
      var colIdx = colIndices[k];
      var cellValue = (colIdx !== -1 && row[colIdx] !== undefined) ? row[colIdx] : "";
      
      if (k === skuIndexInDesired && cellValue !== "") {
        cellValue = cellValue.toString();
      }
      
      newRow.push(cellValue);
    }
    filteredData.push(newRow);
  }

  // Zielblatt leeren & Daten eintragen
  sheet.clearContents();
  sheet.clearFormats();

  var numRows = filteredData.length;
  var numCols = CONFIG.MOBILE_DESIRED_COLUMNS.length;
  var targetRange = sheet.getRange(1, 1, numRows, numCols);

  var skuColNum = skuIndexInDesired + 1;
  if (skuColNum > 0 && numRows > 1) {
    sheet.getRange(2, skuColNum, numRows - 1, 1).setNumberFormat('@');
  }

  targetRange.setValues(filteredData);

  // Zeitstempel in Zeile 1 hinter die Daten schreiben
  var lastUpdateCol = numCols + 1;
  sheet.getRange(1, lastUpdateCol).setValue("last update");
  sheet.getRange(1, lastUpdateCol + 1).setValue(formattedDate);

  Logger.log('Erfolgreich ' + (numRows - 1) + ' Zeilen importiert!');
}