// ==========================================
// 5. MAIN.GS - Hauptausführung & Auto-Close Toast
// ==========================================

function mainSalesAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ss.toast("Lade Archivdaten der letzten 24 Monate...", "Datahub API", 8);
    
    // 1. Daten per GraphQL abrufen & Rohdatenblatt schreiben
    fetchAndSaveRawData();
    
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    
    ss.toast("Analysiere Produkte, Vertreter & Länder...", "Datahub API", 8);
    
    // 2. Daten verarbeiten, Farbmatrix aufbauen & Blätter befüllen
    analyzeAdvancedSalesData();
    
    ss.toast("Vertriebsanalyse erfolgreich aktualisiert!", "Fertig", 5);

  } catch (error) {
    Logger.log("Fehler in mainSalesAnalysis: " + error.toString());
    ss.toast("Fehler bei der Ausführung: " + error.message, "Fehler", 10);
  }
}