/**
 * ZENTRALE KONFIGURATION & ZUGANGSDATEN
 */
const API_URL = 'https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql';
const API_TOKEN = 'e12Bfv!@Ss#asrpPFjucm8a8';

// =========================================================================
// ⚙️ VARIABLEN ZUR STEUERUNG
// =========================================================================
const CATALOG_ID = "19";                   // Katalog-ID für NSY Produkte
const PURCHASE_PRICE_FIELD = "fldEkMitt";  // Mögliche Werte: fldEkMitt, fldEkRoh, fldEkKalk
const MIN_PURCHASE_PRICE = 0.00;           // Mindest-EK (0 = alle anzeigen)
const MAX_EXECUTION_TIME_MS = 5.5 * 60 * 1000; // 5,5 Minuten Zeitbremse
// =========================================================================

/**
 * Erstellt das Menü in der Google Sheets Oberfläche (nur wenn UI vorhanden).
 */
function onOpen() {
  if (isUiAvailable()) {
    SpreadsheetApp.getUi()
      .createMenu('Nepata Datahub')
      .addItem('NSY Produkte','fetchCatalogProducts')
      .addToUi();
  }
}

/**
 * Hilfsfunktion: Prüft, ob eine Benutzeroberfläche (UI) verfügbar ist.
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Zeigt ein Pop-up-Fenster nur an, wenn ein Nutzer im Sheet aktiv ist.
 */
function safeAlert(message) {
  if (isUiAvailable()) {
    SpreadsheetApp.getUi().alert(message);
  } else {
    console.log("[TRIGGER-RUN] " + message);
  }
}

/**
 * Zentrale Funktion für Toasts (nur wenn UI vorhanden).
 */
function safeToast(message, title) {
  const toastTitle = title || "Hinweis";
  console.log("[" + toastTitle + "] " + message);
  
  if (isUiAvailable()) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, toastTitle);
  }
}

/**
 * Hauptabfrage für NSY Produkte aus Katalog 19 (Trigger-sicher)
 */
function fetchCatalogProducts() {
  const startTime = Date.now();
  safeToast("Abfrage für Katalog " + CATALOG_ID + " gestartet...", "Status");

  let hasNextPage = true;
  let cursor = null;
  let allRows = [];
  
  let countTotalCatalogItems = 0;
  let countItemsWithStock = 0;

  // Kopfzeile ("Item name" für fldKuBez1)
  const headers = [
    "Artikelnummer", 
    "Artikelname", 
    "Item name",
    "Lagernummer", 
    "stock level", 
    "VK0 Preis",
    "EK (" + PURCHASE_PRICE_FIELD + ")"
  ];

  try {
    while (hasNextPage) {
      if (Date.now() - startTime > MAX_EXECUTION_TIME_MS) {
        safeToast("Zeitlimit erreicht. Bisher geladene Daten werden geschrieben.", "Warnung");
        break;
      }

      // GraphQL-Query mit orderBy nach fldEkMitt (desc: true)
      const query = 'query ($cursor: String) { tblArtikel { conRead(first: 100, after: $cursor, orderBy: [{field: fldEkMitt, desc: true}], fastFilter: { eq: [{field: fldKatalog}, {value: {string: "' + CATALOG_ID + '"}}] }) { pageInfo { hasNextPage endCursor } edges { node { fldArtNr fldKuBez1 fldKuBez3 fldVk0Preis fldKatalog fldEkMitt fldEkRoh fldEkKalk rowsArtikelLager { fldLagNr fldMge } } } } } }';

      const payload = JSON.stringify({ query: query, variables: { cursor: cursor } });
      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'X-API-Token': API_TOKEN },
        payload: payload,
        muteHttpExceptions: true
      };

      let response;
      try {
        response = UrlFetchApp.fetch(API_URL, options);
      } catch (e) {
        console.error("Network Error:", e);
        safeToast("Netzwerkfehler beim Abruf der Daten.", "Fehler");
        break;
      }

      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 401) {
        console.error("401 Unauthorized:", responseText);
        safeToast("Authentifizierungsfehler (HTTP 401). Bitte API-Zugang prüfen.", "Fehler");
        break;
      }

      let json;
      try {
        json = JSON.parse(responseText);
      } catch (e) {
        console.error("JSON Parse Error. HTTP Code:", responseCode, "Response:", responseText);
        safeToast("Ungültige API-Antwort (kein JSON). Siehe Konsole.", "Fehler");
        break;
      }

      if (json.errors) {
        console.error("GraphQL Errors:", json.errors);
        safeToast("GraphQL Fehler aufgetreten. Siehe Konsolen-Log.", "Fehler");
        break;
      }

      if (!json.data || !json.data.tblArtikel || !json.data.tblArtikel.conRead) {
        console.error("Unexpected API Structure:", json);
        safeToast("API lieferte nicht die erwartete Datenstruktur.", "Fehler");
        break;
      }

      const connection = json.data.tblArtikel.conRead;
      const edges = connection.edges || [];

      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node;
        countTotalCatalogItems++; 

        const articleNumber = String(node.fldArtNr || "");
        const description1 = node.fldKuBez1 || "";
        const description3 = node.fldKuBez3 || "";
        const vk0Price = node.fldVk0Preis || 0.0;
        const purchasePrice = node[PURCHASE_PRICE_FIELD] || 0.0;

        const stockRows = node.rowsArtikelLager || [];
        let hasPushedStock = false;
        
        for (let j = 0; j < stockRows.length; j++) {
          const stockItem = stockRows[j];
          const stockQuantity = Math.round(stockItem.fldMge || 0);
          
          // Filter: Nur Zeilen mit Bestand > 0 und EK >= Mindest-EK
          if (stockQuantity > 0 && purchasePrice >= MIN_PURCHASE_PRICE) {
            allRows.push({
              articleNumber: articleNumber,
              description1: description1,
              description3: description3,
              warehouseNr: String(stockItem.fldLagNr || ""),
              stockQuantity: stockQuantity,
              vk0Price: vk0Price,
              purchasePrice: purchasePrice
            });
            hasPushedStock = true;
          }
        }
        
        if (hasPushedStock) {
          countItemsWithStock++;
        }
      }

      hasNextPage = connection.pageInfo.hasNextPage;
      cursor = connection.pageInfo.endCursor;
    }

    // Clientseitige Sortierung nach EK (absteigend)
    allRows.sort((a, b) => b.purchasePrice - a.purchasePrice);

    // In zweidimensionales Ausgabearray umwandeln
    const outputRows = [headers];
    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];
      outputRows.push([
        row.articleNumber,
        row.description1,
        row.description3,
        row.warehouseNr,
        row.stockQuantity,
        row.vk0Price,
        row.purchasePrice
      ]);
    }

    // Ergebnisse ins Sheet schreiben
    if (outputRows.length > 1) {
      writeToSheet(outputRows);
      safeAlert("Suchlauf beendet!\n\nGefundene Artikel im Katalog '" + CATALOG_ID + "': " + countTotalCatalogItems + "\nDavon mit Bestand > 0 und EK >= " + MIN_PURCHASE_PRICE + ": " + countItemsWithStock);
    } else {
      safeToast("Keine passenden Artikel gefunden.", "Hinweis");
    }

  } catch (err) {
    console.error("Fehler bei der Ausführung:", err);
    safeToast("Ausführungsfehler: " + err.message, "Fehler");
  }
}

/**
 * Schreibt die Ergebnisse ins Tabellenblatt und aktiviert den Autofilter.
 */
function writeToSheet(dataRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "NSY Kat" + CATALOG_ID + ")";
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clear();
  
  const lastRow = dataRows.length;
  const lastCol = dataRows[0].length;
  
  if (lastRow > 1) {
    // Textformatierung (@) für Spalte 1 (Artikelnummer) und Spalte 4 (Lagernummer)
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("@");
    sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat("@");
    
    // Ganzzahlformat ("0") für Spalte 5 (stock level)
    sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat("0");
  }

  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setValues(dataRows);

  // Kopfzeilen-Formatierung
  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  sheet.autoResizeColumns(1, lastCol);

  // Autofilter im Tabellenblatt einrichten
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  sheet.getDataRange().createFilter();
  
  // Währungsformatierung (#,##0.00 €) für Spalte 6 (VK0 Preis) und Spalte 7 (EK)
  if (lastRow > 1) {
    sheet.getRange(2, 6, lastRow - 1, 2).setNumberFormat("#,##0.00 €");
  }
  safeToast("Daten erfolgreich im Sheet aktualisiert.", "Erfolg");
}