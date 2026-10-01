// ==========================================
// FILE: Main.gs
// ==========================================
function checkOldVorkasseAndNotify() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  safeToast("Prüfe alte Vorkasse-Vorgänge...", "Start");

  // 1. Abfrage an Datahub v2
  const { belegeNachVertreter, detailsRows, totalFoundDocs } = fetchAlteVorkasseVorgaenge(startTime);

  // 2. Tabellenblatt "Auftragsdetails" aktualisieren
  writeDetailsSheet(ss, detailsRows);

  if (totalFoundDocs === 0) {
    safeToast("Keine alten Vorkasse-Belege gefunden.", "Fertig");
    Logger.log("Keine passenden Belege gefunden.");
    return;
  }

  // 3. E-Mails versenden & protokollieren
  sendVorkasseMails(ss, belegeNachVertreter, totalFoundDocs);
}