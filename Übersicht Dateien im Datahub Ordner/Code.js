function listFilesInFolder() {
  // Trage HIER die ID deines Google Drive Ordners ein
  var folderId = '1tMH5rKOuCWaWE-ryTVd0q83qWd6TwYU5'; 
  var targetSheetName = 'Update Sheet'; // Name des Zielblattes
  var dataSheetName = 'DATA'; // Name des Referenzblattes
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(targetSheetName);
  var dataSheet = spreadsheet.getSheetByName(dataSheetName);
  
  // Prüfen, ob das Blatt "Update Sheet" existiert. Wenn nicht, wird es erstellt.
  if (!sheet) {
    sheet = spreadsheet.insertSheet(targetSheetName);
  }
  
  // WICHTIG: Wenn auf dem Blatt bereits ein Filter aktiv ist, muss dieser erst entfernt werden, 
  // da 'sheet.clear()' den Filterbereich sonst durcheinanderbringen kann.
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  
  // Löscht den bisherigen Inhalt, Formate und Filter des Blattes
  sheet.clear(); 
  
  // Kopfzeilen für die Daten erstellen
  var headers = ["Dateiname", "Link zur Datei", "Erstellt am", "Zuletzt bearbeitet"];
  sheet.appendRow(headers);
  
  // --- Formatierung der Kopfzeile (A1:D1) ---
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold")
             .setBackground("#e0e0e0") // Leicht grauer Hintergrund
             .setBorder(true, true, true, true, null, null); // Rahmen um die Kopfzeile
             
  // Erste Zeile fixieren (damit sie beim Scrollen immer sichtbar bleibt)
  sheet.setFrozenRows(1);
  
  // --- Zeitstempel der Aktualisierung einfügen ---
  var currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  var timestampCell = sheet.getRange(1, 6); // Zelle F1
  timestampCell.setValue("Letzte Aktualisierung: " + currentTime);
  timestampCell.setFontWeight("bold"); 
  
  // --- NEU: Dateinamen aus dem Blatt "DATA" auslesen ---
  var existingFileNames = new Set();
  if (dataSheet) {
    var lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      // Holt alle Einträge aus Spalte A (Dateiname) im Blatt "DATA"
      var names = dataSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (names[i][0]) {
          // Dateinamen speichern (trim() entfernt überflüssige Leerzeichen)
          existingFileNames.add(names[i][0].toString().trim());
        }
      }
    }
  }
  
  // Ordner abrufen
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  var data = [];
  var backgrounds = []; // NEU: Array für die Hintergrundfarben der Zeilen
  
  // Alle Dateien im Ordner durchgehen
  while (files.hasNext()) {
    var file = files.next();
    
    // Aktuell nur Google Sheets auflisten
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var fileName = file.getName();
      
      data.push([
        fileName,
        file.getUrl(),
        file.getDateCreated(),
        file.getLastUpdated()
      ]);
      
      // --- NEU: Prüfen, ob der Dateiname in "DATA" fehlt ---
      if (!existingFileNames.has(fileName.trim())) {
        // Fehlend: Die 4 Zellen dieser Zeile z.B. hellrot markieren (#f8cecc)
        backgrounds.push(["#f8cecc", "#f8cecc", "#f8cecc", "#f8cecc"]); 
      } else {
        // Vorhanden: Kein Hintergrund (weiß / null)
        backgrounds.push([null, null, null, null]);
      }
    }
  }
  
  // Daten in das Sheet schreiben, falls Dateien gefunden wurden
  if (data.length > 0) {
    // Ab Zeile 2 einfügen
    var dataRange = sheet.getRange(2, 1, data.length, headers.length);
    dataRange.setValues(data);
    
    // --- NEU: Hintergrundfarben in einem Rutsch anwenden ---
    dataRange.setBackgrounds(backgrounds);
    
    // --- Datumsspalten (C und D) richtig formatieren ---
    sheet.getRange(2, 3, data.length, 2).setNumberFormat("dd.MM.yyyy HH:mm");
    
    // --- Filter über die Datenbereiche (A bis D) setzen ---
    sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
    
    // Die Spaltenbreite automatisch anpassen (Spalte 1 bis 6)
    sheet.autoResizeColumns(1, 6);
  } else {
    sheet.appendRow(["Keine Google Sheets in diesem Ordner gefunden.", "", "", ""]);
  }
}