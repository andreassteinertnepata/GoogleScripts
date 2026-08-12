function translateSheetNamesFromExternalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Externe Übersetzungstabelle per ID öffnen
  const EXTERNAL_SPREADSHEET_ID = "1urqEUeu1TdGBU-MA_lnUuGzuiWHosOix4qYaB6mrACQ";
  const SHEET_NAMES_TAB = "Sheets"; // Tabellenblatt mit den Blattnamen-Übersetzungen
  
  let sheetTranslationTab;
  try {
    const externalSs = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
    sheetTranslationTab = externalSs.getSheetByName(SHEET_NAMES_TAB);
  } catch (e) {
    Logger.log("Fehler beim Zugriff auf das externe Übersetzungs-Sheet: " + e.message);
    return;
  }
  
  if (!sheetTranslationTab) {
    Logger.log(`Das Blatt '${SHEET_NAMES_TAB}' wurde im externen Sheet nicht gefunden.`);
    return;
  }
  
  // 2. Übersetzungen für Blattnamen in eine Map laden (Spalte A -> Spalte B)
  const sheetData = sheetTranslationTab.getDataRange().getValues();
  const nameTranslationMap = {};
  
  for (let i = 0; i < sheetData.length; i++) {
    const originalName = String(sheetData[i][0] || "").trim();    // Spalte A: Such-Blattname
    const translatedName = String(sheetData[i][1] || "").trim();  // Spalte B: Übersetzung
    
    if (originalName && translatedName) {
      nameTranslationMap[originalName] = translatedName;
    }
  }
  
  // 3. Alle Tabellenblätter in der aktuellen Datei durchlaufen und umbenennen
  const sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    const currentName = sheet.getName().trim();
    
    // Steuerungs-/Systemblätter nicht umbenennen
    if (currentName === "Values" || currentName === "Header" || currentName === "Sheet" || currentName === "ColorMapping") {
      return;
    }
    
    if (nameTranslationMap[currentName]) {
      const newName = nameTranslationMap[currentName];
      
      // Nur umbenennen, wenn sich der Name wirklich unterscheidet und der Zielname noch nicht vergeben ist
      if (currentName !== newName && !ss.getSheetByName(newName)) {
        sheet.setName(newName);
      }
    }
  });
}