/**
 * Info.gs - Generiert das Informations- und Dokumentationsblatt "Info"
 */

function generateInfoSheet(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Info";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const now = new Date();
  const dateFrom12M = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  const timeProgress = getTimeProgress();

  sheet.getRange("A1").setValue("System-Informationen & Dokumentation").setFontSize(16).setFontWeight("bold");

  sheet.getRange("A3:B3").setValues([["System-Status & Auswertungszeitraum", "Wert"]])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

  const statusRows = [
    ["Letzte Aktualisierung", Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss")],
    ["Startdatum Auswertung (12M)", Utilities.formatDate(dateFrom12M, Session.getScriptTimeZone(), "dd.MM.yyyy")],
    ["Enddatum Auswertung (Heute)", Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yyyy")],
    ["Monatsfortschritt (Heute)", timeProgress.monthProgress],
    ["Jahresfortschritt (Heute)", timeProgress.yearProgress]
  ];

  sheet.getRange(4, 1, statusRows.length, 2).setValues(statusRows);
  sheet.getRange("B7:B8").setNumberFormat("0.0%");

  let currRow = 11;
  sheet.getRange(currRow, 1, 1, 2).setValues([["Konfigurierte Variable (Config.gs)", "Einstellung / Wert"]])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
  currRow++;

  const configRows = [
    ["Ziel-Katalog", CONFIG.TARGET_CATALOG],
    ["Mindest-EK für Artikel (€)", CONFIG.MIN_EK],
    ["KeyAccount Mindest-Umsatz (€)", CONFIG.KEYACCOUNT_MIN_UMSATZ],
    ["Farbe Übererfüllung (>100%)", CONFIG.COLOR_OVER_100],
    ["Farbe Im Soll (90-99%)", CONFIG.COLOR_90_TO_99],
    ["Farbe Warnung (50-89%)", CONFIG.COLOR_50_TO_89],
    ["Farbe Rückstand (<50%)", CONFIG.COLOR_UNDER_50]
  ];

  sheet.getRange(currRow, 1, configRows.length, 2).setValues(configRows);
  currRow += configRows.length + 2;

  sheet.getRange(currRow, 1, 1, 3).setValues([["Vertreter-ID", "Mitarbeiter Name", "Dashboard Tab"]])
       .setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
  currRow++;

  const repRows = [];
  Object.keys(CONFIG.SALES_REPS).forEach(function(id) {
    repRows.push([id, CONFIG.SALES_REPS[id].name, CONFIG.SALES_REPS[id].tab]);
  });

  sheet.getRange(currRow, 1, repRows.length, 3).setValues(repRows);
  sheet.getRange(currRow, 1, repRows.length, 1).setNumberFormat("@");
  currRow += repRows.length + 2;

  sheet.getRange(currRow, 1, 1, 2).setValues([["Prozess-Schritt", "Beschreibung der Hintergrund-Abläufe"]])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
  currRow++;

  const guideRows = [
    ["1. Datahub v2 Abruf", "Lädt alle Belege der letzten 12 Monate (tblVorgangArchiv für Ist-Umsätze, tblVorgang für offene Angebote/Aufträge) via GraphQL mit Paginierung."],
    ["2. Kaufmännische Filterung", "Entfernt Testbelege (fldSel14=true), ignoriert Stornos (fldStorniertKz=true) und dreht das Vorzeichen der Menge bei Gutschriften/Korrekturen (90, 123, 156)."],
    ["3. Artikel-Stammdaten", "Lädt alle Artikel aus Katalog 1 mit EK Roh >= 50 € und schlüsselt die 12M-Verkaufsstückzahlen pro Betreuer in separaten Spalten im Blatt 'Artikel' auf."],
    ["4. Ziel-Konfigurationen", "Erstellt & synchronisiert 'Ziele_Konfiguration', 'Produktziele_Konfiguration' und 'KeyAccounts_Konfiguration' mit 12M-Durchschnittswerten. Manuell geänderte Monatsziele bleiben unverändert erhalten."],
    ["5. Pro-Rata Zeitrechnung", "Ermittelt den exakten Zeitfortschritt im Monat (Tag/Tage des Monats) und Jahr (Tag/Tage des Jahres) zur Berechnung des Soll-Wertes bis 'heute'."],
    ["6. Dashboard-Rendering", "Erzeugt für jeden Vertreter ein eigenes Dashboard mit pro-rata Zielerreichung heute in % und dynamischer 4-Farben-Skala."]
  ];

  sheet.getRange(currRow, 1, guideRows.length, 2).setValues(guideRows);
  currRow += guideRows.length + 2;

  sheet.getRange(currRow, 1, 1, 3).setValues([["Spaltenüberschrift (Header)", "Tabellenblatt", "Erklärung & Formel / Datenherkunft"]])
       .setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
  currRow++;

  const headerDocs = [
    ["Durchschnitt 12M (€ / Stk.)", "Konfigurationen", "Errechneter Monatsdurchschnitt der letzten 12 Monate (Summe 12M / 12). Dient als Vorschlag und Orientierungswert."],
    ["Ziel Monat (€ / Stk.)", "Konfigurationen", "Das manuell vorgegebene Monatsziel für Umsatz, Produktstückzahlen oder Key Accounts."],
    ["Soll Monat heute (€ / Stk.)", "Dashboards", "Ziel Monat × Monatsfortschritt (aktueller Tag / Gesamttage des laufenden Monats)."],
    ["Ist MTD (€ / Stk.)", "Dashboards", "Actual Month-To-Date: Bereits fakturierter Umsatz / Stückzahl im laufenden Kalendermonat."],
    ["Status Monat heute (%)", "Dashboards", "Zielerreichung heute in % (Ist MTD / Soll Monat heute). Dynamisch eingefärbt."],
    ["Ziel Jahr (€ / Stk.)", "Dashboards", "Ziel Monat × 12. Das hochgerechnete Jahresziel."],
    ["Soll Jahr heute (€ / Stk.)", "Dashboards", "Ziel Jahr × Jahresfortschritt (aktueller Tag / Gesamttage des laufenden Jahres)."],
    ["Ist YTD (€ / Stk.)", "Dashboards", "Actual Year-To-Date: Fakturierter Umsatz / Stückzahl im laufenden Kalenderjahr."],
    ["Status Jahr heute (%)", "Dashboards", "Zielerreichung heute in % (Ist YTD / Soll Jahr heute). Dynamisch eingefärbt."],
    ["Prognose Offen (€ / Stk.)", "Dashboards", "Volumen der offenen, noch nicht fakturierten Aufträge (tblVorgang) in der Pipeline."],
    ["Menge Gesamt (12M)", "Artikel", "Gesamte verkaufte Stückzahl aller Vertriebler für diesen Artikel in den letzten 12 Monaten."],
    ["Stk [Mitarbeiter Name]", "Artikel", "Vom jeweiligen Mitarbeiter in den letzten 12 Monaten verkaufte Stückzahl dieses Artikels."]
  ];

  sheet.getRange(currRow, 1, headerDocs.length, 3).setValues(headerDocs);

  sheet.getRange("A:C").setVerticalAlignment("top");
  sheet.getRange("B:B").setWrap(true);
  sheet.getRange("C:C").setWrap(true);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(3, 460);
}