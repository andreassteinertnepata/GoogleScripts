
/**
 * HAUPTFUNKTION: Steuert den gesamten Ablauf nacheinander
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast("Update running...", "Please wait", -1);

    // Schritt 1: Secabo Stock & GraphQL Daten verarbeiten
  
    importArchivKomplett();
    // Schritt 2: Mobile CSV Datei aus Drive importieren
    Mobile_importierenUndAktualisieren();
    
    fetchArticles();
    fetchStuecklisten();
    translateHeadersFromExternalSheet();
    translateValuesFromExternalSheet();
    // Schritt 3: Alle Tabellenblätter nach Corporate Design formatieren
    formatGoogleSheet();
    // WICHTIG: Erwirkt das sofortige Anwenden aller Änderungen und Neuberechnen aller Formeln
    SpreadsheetApp.flush();

 
    ss.toast("Everything is up to date", "Success", 5);

  } catch (e) {
    Logger.log("FEHLER in main(): " + e.toString());
    SpreadsheetApp.getUi().alert("Fehler bei der Ausführung: " + e.toString());
  }
}