/**
 * Main.js - Hauptsteuerung
 */

function updateUmsatzziele() {
  const startTime = Date.now();
  
  try {
    safeToast("Ist-Umsätze & Verkaufszahlen der letzten 12M abfragen...", "Start", 5);
    
    // 1. Verkaufszahlen und Umsätze aus Datahub v2 abfragen
    const stats = fetchSalesData(startTime);

    // 2. Artikel aus Katalog 1 laden & Verkaufszahlen pro Betreuer im Blatt 'Artikel' aufschlüsseln
    safeToast("Katalog 1 Artikel & Stückzahlen laden...", "Artikel", 5);
    const articles = fetchArticles(stats);

    // 3. Konfigurationsblätter aktualisieren (Durchschnittswerte berechnen, manuelle Ziele behalten)
    safeToast("Ziel-Konfigurationen synchronisieren...", "Konfiguration", 5);
    const targets = syncAndGetTargetsConfig(stats);
    
    // articles explizit übergeben
    const productTargets = syncAndGetProductTargetsConfig(articles, stats);
    const keyAccountTargets = syncAndGetKeyAccountsConfig(stats);

    // 4. Vertreter-Dashboards rendern & Farbskala anwenden
    safeToast("Vertriebs-Dashboards rendern...", "Dashboards", 5);
    renderDashboards(stats, targets, productTargets, keyAccountTargets);

    // 5. Info-Blatt aktualisieren
    safeToast("Info-Blatt aktualisieren...", "Info", 5);
    generateInfoSheet(stats);

    safeToast("Umsatz-, Produkt- & Key-Account-Dashboards erfolgreich aktualisiert!", "Erfolg", 5);

  } catch (error) {
    console.error("Fehler bei updateUmsatzziele: ", error);
    safeToast("Fehler bei der Aktualisierung: " + error.message, "Fehler", 10);
  }
}

function safeToast(message, title, timeout) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title || "Info", timeout || 5);
  } catch (e) {
    console.log(`[Toast] ${title}: ${message}`);
  }
}