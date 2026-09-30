function listFilesInFolder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var infoSheetName = 'Info';
  var infoSheet = spreadsheet.getSheetByName(infoSheetName);
  
  // --- NEU: Wenn das Info-Blatt noch nicht existiert, wird es automatisch erstellt ---
  if (!infoSheet) {
    createInfoSheet(spreadsheet, infoSheetName);
    SpreadsheetApp.getUi().alert(
      "Das Blatt 'Info' wurde erfolgreich erstellt!\n\n" +
      "Bitte prüfe dort die angelegten Variablen (Ordner-ID, Blattnamen) und führe das Skript danach noch einmal aus."
    );
    return; // Skript bricht hier beim ersten Mal ab, damit du die Variablen ansehen kannst
  }
  
  // --- 1. Variablen direkt aus dem "Info"-Blatt auslesen ---
  var folderId = infoSheet.getRange("B4").getValue();
  var targetSheetName = infoSheet.getRange("B5").getValue();
  var dataSheetName = infoSheet.getRange("B6").getValue();
  var missingColor = infoSheet.getRange("B7").getValue();
  
  var sheet = spreadsheet.getSheetByName(targetSheetName);
  var dataSheet = spreadsheet.getSheetByName(dataSheetName);
  
  // Prüfen, ob das Zielblatt existiert. Wenn nicht, wird es erstellt.
  if (!sheet) {
    sheet = spreadsheet.insertSheet(targetSheetName);
  }
  
  // Filter entfernen und Zielblatt leeren
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  sheet.clear(); 
  
  // Kopfzeilen für die Daten erstellen
  var headers = ["Dateiname", "Link zur Datei", "Erstellt am", "Zuletzt bearbeitet"];
  sheet.appendRow(headers);
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold")
             .setBackground("#e0e0e0") 
             .setBorder(true, true, true, true, null, null); 
             
  sheet.setFrozenRows(1);
  
  // Zeitstempel der Aktualisierung einfügen
  var currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  var timestampCell = sheet.getRange(1, 6); // Zelle F1
  timestampCell.setValue("Letzte Aktualisierung: " + currentTime);
  timestampCell.setFontWeight("bold"); 
  
  // --- 2. Referenzdaten aus "DATA" (bzw. dem in der Variable definierten Blatt) einlesen ---
  var existingFileNames = new Set();
  if (dataSheet) {
    var lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      var names = dataSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (names[i][0]) {
          existingFileNames.add(names[i][0].toString().trim());
        }
      }
    }
  }
  
  // --- 3. Ordner abrufen ---
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  var data = [];
  var backgrounds = []; 
  
  // Alle Dateien im Ordner durchgehen
  while (files.hasNext()) {
    var file = files.next();
    
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var fileName = file.getName();
      
      data.push([
        fileName,
        file.getUrl(),
        file.getDateCreated(),
        file.getLastUpdated()
      ]);
      
      // Prüfen, ob der Dateiname in "DATA" fehlt
      if (!existingFileNames.has(fileName.trim())) {
        // Fehlend: Die 4 Zellen dieser Zeile mit der Variablen-Farbe markieren
        backgrounds.push([missingColor, missingColor, missingColor, missingColor]); 
      } else {
        // Vorhanden: Kein Hintergrund
        backgrounds.push([null, null, null, null]);
      }
    }
  }
  
  // --- 4. Daten in das Sheet schreiben ---
  if (data.length > 0) {
    var dataRange = sheet.getRange(2, 1, data.length, headers.length);
    dataRange.setValues(data);
    dataRange.setBackgrounds(backgrounds);
    
    sheet.getRange(2, 3, data.length, 2).setNumberFormat("dd.MM.yyyy HH:mm");
    sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
    sheet.autoResizeColumns(1, 6);
  } else {
    sheet.appendRow(["Keine Google Sheets in diesem Ordner gefunden.", "", "", ""]);
  }
}

// =========================================================================
// HILFSFUNKTION: Baut das "Info"-Blatt inkl. Variablen und Erklärung auf
// =========================================================================
function createInfoSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.insertSheet(sheetName, 0); // Fügt das Blatt ganz vorne (Position 0) ein
  
  // Titel
  sheet.getRange("A1").setValue("Skript Info & Variablen").setFontWeight("bold").setFontSize(14);
  
  // Variablen-Kopfbereich
  sheet.getRange("A3:B3").setValues([["Variable", "Aktueller Wert"]]).setFontWeight("bold").setBackground("#e0e0e0");
  
  // Die Variablen selbst (werden ab jetzt vom Skript hier ausgelesen)
  sheet.getRange("A4:B7").setValues([
    ["Ordner-ID (Google Drive):", "1tMH5rKOuCWaWE-ryTVd0q83qWd6TwYU5"],
    ["Name des Zielblattes:", "Update Sheet"],
    ["Name des Referenzblattes:", "DATA"],
    ["Warnfarbe (Hex-Code):", "#f8cecc"]
  ]);
  
  // Eingabefelder für die Variablen gelb hinterlegen zur Verdeutlichung
  sheet.getRange("B4:B7").setBackground("#fff2cc").setBorder(true, true, true, true, null, null);
  
  // Erklärungs-Titel
  sheet.getRange("A10").setValue("Was passiert in diesem Skript?").setFontWeight("bold").setFontSize(12);
  
  // Detaillierte Schritt-für-Schritt-Erklärung
  var beschreibung = [
    ["1. Variablen einlesen:", "Das Skript holt sich die oben definierten Variablen (Ordner-ID, Blattnamen, Farbe) dynamisch aus diesem Blatt."],
    ["2. Referenzdaten laden:", "Es liest alle vorhandenen Dateinamen aus Spalte A des Blattes aus, das in Zelle B6 definiert ist (z.B. 'DATA')."],
    ["3. Drive Ordner scannen:", "Es durchsucht den Google Drive Ordner (Zelle B4) nach Google Sheets Dateien."],
    ["4. Abgleich:", "Jede im Ordner gefundene Datei wird mit der Liste aus dem Referenzblatt ('DATA') abgeglichen."],
    ["5. Markierung:", "Fehlt eine Datei im Referenzblatt, wird sie in der neuen Übersicht mit der Farbe aus Zelle B7 markiert."],
    ["6. Ausgabe:", "Die Ergebnisse werden im Zielblatt (Zelle B5) übersichtlich inkl. Zeitstempel ausgegeben."]
  ];
  
  sheet.getRange("A12:B17").setValues(beschreibung);
  sheet.getRange("A12:A17").setFontWeight("bold"); // Spalte A der Beschreibung fett
  
  // Spaltenbreiten anpassen für gute Lesbarkeit
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 600);
}