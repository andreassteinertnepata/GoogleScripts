function MainFunction() {
  // 1. Daten per API holen und ins Blatt schreiben
  importGraphQLData();
  
  // Zwingt Google Sheets, alle Formeln (VLOOKUP in K) sofort fertig zu berechnen
  SpreadsheetApp.flush();
  Utilities.sleep(1000); // Wartet sicherheitshalber 1 Sekunde
  
}