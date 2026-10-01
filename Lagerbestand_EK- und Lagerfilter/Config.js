// ==========================================
// CENTRAL CONFIGURATION
// ==========================================
const CONFIG = {
  apiUrl: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  token: "e12Bfv!@Ss#asrpPFjucm8a8",
  targetSheetName: "Lager_Pivot_EK_Filter",
  infoSheetName: "Info",
  
  // Feste Lagernummern, die als eigene Spalten dargestellt werden sollen
  targetWarehouses: ["1", "2", "200", "13"],
  
  // Filter: Mindest-Einkaufspreis (EK). Gilt für fldEkRoh
  minEkThreshold: 300,

  // Internationales Vertreterteam
  salesReps: [
    { code: "28", name: "Alexis Fonte", lang: "Deutsch" },
    { code: "36", name: "Imad Nassef", lang: "Deutsch" },
    { code: "43", name: "Andreas Steinert", lang: "Deutsch" },
    { code: "46", name: "Stefanie Binder", lang: "Deutsch" },
    { code: "56", name: "Robin Carter Browne", lang: "Englisch" },
    { code: "59", name: "Aneta Nedyalkova", lang: "Deutsch" },
    { code: "60", name: "Rado Kabakov", lang: "Englisch" }
  ]
};

// ==========================================
// GRAPHQL QUERY
// ==========================================
const STOCK_PIVOT_QUERY = `
query GetHighValueStockPivoted(
  $cursor: String, 
  $minEk: Float!
) {
  tblArtikel {
    conRead(
      first: 100,
      after: $cursor,
      fastFilter: {
        ge: [{ field: fldEkRoh }, { value: { float: $minEk } }]
      }
    ) {
      edges {
        node {
          fldArtNr
          fldKuBez1
          fldEkRoh
          fldEkKalk
          fldVk0Preis
          fldLiefNr
          fldKatalog
          lblKatalog
          rowsArtikelLager {
            fldLagNr
            fldMge
            fldKBstMge
          }
          rowsArtikelLieferanten {
            fldAdrNr
            fldStdKz
            rowAdresse {
              rowsAnschriften {
                fldNa1
                fldNa2
              }
            }
          }
        }
      }
      pageInfo {
        hasNextPage
        endCursor
      }
    }
  }
}
`;

/**
 * Hauptfunktion: Lädt Artikel ab einem bestimmten EK, pivotiert die Läger als Spalten,
 * ergänzt Katalognummer & Katalogname, VK-Preis sowie Lieferantendaten, filtert leere Bestände aus
 * und erzeugt/aktualisiert das Info-Blatt.
 */
function importLagerBestandPivot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.targetSheetName);

  // 1. Zielblatt zurücksetzen oder neu anlegen
  if (sheet) {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.targetSheetName);
  }

  safeToast(ss, `Lade Bestände für EK >= ${CONFIG.minEkThreshold}€...`, "Bitte warten", -1);

  // 2. Dynamische Tabellenspalten (Header) aufbauen
  const headers = [
    "Article_No", 
    "Description", 
    "Catalog_No",    // Katalognummer (fldKatalog)
    "Catalog_Name",  // Katalogbezeichnung / Klartext (lblKatalog)
    "EK_Roh", 
    "EK_Kalk",
    "VK_0_Price",
    "Supplier_No",
    "Supplier_Name"
  ];
  
  // Pivot-Spalten pro Lager hinzufügen
  CONFIG.targetWarehouses.forEach(lagNr => {
    headers.push(`Lager_${lagNr}_Stock`);
    headers.push(`Lager_${lagNr}_Order`);
    headers.push(`Lager_${lagNr}_Available`);
  });

  // Abschluss-Summenspalten
  headers.push("Total_Stock", "Total_Order", "Total_Available");

  const rowsToWrite = [headers];
  let hasNextPage = true;
  let cursor = null;
  let pageCount = 0;
  
  const startTime = Date.now();
  const MAX_TIME_MS = 5.5 * 60 * 1000; // 5,5 Minuten Timeout-Sperre

  // 3. API Loop mit Pagination
  while (hasNextPage) {
    pageCount++;

    // Zeitbremse prüfen (Timeout abfangen)
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit von 5,5 Min. erreicht. Zwischenstand wird gespeichert.");
      safeToast(ss, "Zeitlimit erreicht! Zwischenstand wird gespeichert...", "Achtung", 5);
      break;
    }

    const payload = {
      query: STOCK_PIVOT_QUERY,
      variables: {
        cursor: cursor,
        minEk: CONFIG.minEkThreshold
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.apiUrl, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Fehler: " + JSON.stringify(json.errors || response.getContentText()));
        safeToast(ss, "Fehler bei der GraphQL API Abfrage.", "Fehler");
        return;
      }

      const conRead = json.data.tblArtikel?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const warehouses = node.rowsArtikelLager || [];

        // Basis-Daten des Artikels
        const artNr = String(node.fldArtNr || "");
        const desc = String(node.fldKuBez1 || "");
        const catalogNo = String(node.fldKatalog || "");
        const catalogName = String(node.lblKatalog || "");
        const ekRoh = node.fldEkRoh || 0;
        const ekKalk = node.fldEkKalk || 0;
        const vkPreis = node.fldVk0Preis || 0;

        // Lieferantendaten ermitteln
        let liefNr = String(node.fldLiefNr || "").trim();
        let liefName = "";
        const artLiefList = node.rowsArtikelLieferanten || [];

        // Suche nach Standard-Lieferant (fldStdKz === true), matching Adressnummer oder den ersten Eintrag
        const stdLief = artLiefList.find(l => l.fldStdKz === true) || 
                        artLiefList.find(l => String(l.fldAdrNr || "").trim() === liefNr) || 
                        artLiefList[0];

        if (stdLief) {
          if (!liefNr && stdLief.fldAdrNr) {
            liefNr = String(stdLief.fldAdrNr).trim();
          }
          const anschriften = stdLief.rowAdresse?.rowsAnschriften || [];
          if (anschriften.length > 0) {
            const na1 = String(anschriften[0].fldNa1 || "").trim();
            const na2 = String(anschriften[0].fldNa2 || "").trim();
            liefName = na2 ? `${na1} ${na2}` : na1;
          }
        }

        // Struktur für Lagerdaten & Summen
        const lagData = {};
        let totalStock = 0;
        let totalOrder = 0;
        let totalAvailable = 0;

        CONFIG.targetWarehouses.forEach(w => {
          lagData[w] = { stock: 0, order: 0, available: 0 };
        });

        // Werte aus den Lagern konsolidieren
        warehouses.forEach(lager => {
          const lagNr = String(lager.fldLagNr || "").trim();
          
          if (CONFIG.targetWarehouses.includes(lagNr)) {
            const stockQty = lager.fldMge || 0;
            const orderQty = lager.fldKBstMge || 0;
            const availQty = stockQty - orderQty;
            
            lagData[lagNr].stock += stockQty;
            lagData[lagNr].order += orderQty;
            lagData[lagNr].available += availQty;

            totalStock += stockQty;
            totalOrder += orderQty;
            totalAvailable += availQty;
          }
        });

        // FILTER: Wenn über alle gewählten Läger weder Bestand noch Bestellung existiert, überspringen
        if (totalStock === 0 && totalOrder === 0) {
          return;
        }

        // Zeilen-Array zusammenbauen
        const row = [
          artNr, 
          desc, 
          catalogNo,
          catalogName,
          ekRoh, 
          ekKalk, 
          vkPreis, 
          liefNr, 
          liefName
        ];
        
        CONFIG.targetWarehouses.forEach(lagNr => {
          row.push(
            lagData[lagNr].stock,
            lagData[lagNr].order,
            lagData[lagNr].available
          );
        });

        // Summen am Ende anfügen
        row.push(totalStock, totalOrder, totalAvailable);

        rowsToWrite.push(row);
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Netzwerk-/Skriptfehler: " + e.toString());
      safeToast(ss, "Fehler: " + e.message, "Fehler");
      break;
    }
  }

  // 4. Keine Daten gefunden
  if (rowsToWrite.length <= 1) {
    sheet.getRange(1, 1).setValue("Keine Bestände für die gewählten Kriterien gefunden (oder alles war 0).");
    createOrUpdateInfoSheet(ss, 0);
    safeToast(ss, "Abfrage beendet - Keine Daten gefunden.", "Hinweis");
    return;
  }

  // 5. In Google Sheet schreiben & Formatieren
  const totalRows = rowsToWrite.length;
  const totalCols = headers.length;

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(rowsToWrite);

  if (totalRows > 1) {
    // Textformate für Identifikatoren (Spalten A bis D)
    sheet.getRange(2, 1, totalRows - 1, 4).setNumberFormat("@");
    
    // EK- und VK-Preise formatieren (Spalten E, F, G)
    sheet.getRange(2, 5, totalRows - 1, 3).setNumberFormat("#,##0.00 €");

    // Textformate für Lieferantennummer & Name (Spalten H & I)
    sheet.getRange(2, 8, totalRows - 1, 2).setNumberFormat("@");
    
    // Alle Lagermengen (Einzel-Läger + Summen) formatieren (ab Spalte J bis Ende)
    const qtyColsCount = totalCols - 9;
    if (qtyColsCount > 0) {
      sheet.getRange(2, 10, totalRows - 1, qtyColsCount).setNumberFormat("#,##0");
    }
  }

  // Header-Styling (Dunkelblau)
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  // Summenspalten am Ende hervorheben
  const sumColsStartIndex = totalCols - 2;
  sheet.getRange(1, sumColsStartIndex, totalRows, 3).setBackground("#eaf2f8");
  sheet.getRange(1, sumColsStartIndex, 1, 3).setBackground("#1a252f");

  // Rahmen ziehen & Auto-Filter setzen
  range.setBorder(true, true, true, true, true, true);
  range.createFilter();
  
  // Fixieren der ersten Spalten & Zeile
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Spaltenbreiten anpassen
  sheet.autoResizeColumns(1, totalCols);

  // 6. Info-Tabellenblatt mit Variablen & Ausführungsstatus aktualisieren
  createOrUpdateInfoSheet(ss, totalRows - 1);

  safeToast(ss, `Erfolgreich ${totalRows - 1} pivotierte Artikel geladen!`, "Erfolg", 5);
}

/**
 * Erstellt oder aktualisiert das Info-Tabellenblatt mit den zentralen Skriptvariablen.
 */
function createOrUpdateInfoSheet(ss, recordsCount) {
  const infoSheetName = CONFIG.infoSheetName || "Info";
  let infoSheet = ss.getSheetByName(infoSheetName);

  if (infoSheet) {
    infoSheet.clear();
  } else {
    infoSheet = ss.insertSheet(infoSheetName);
  }

  const infoData = [
    ["Parameter / Variable", "Aktueller Wert", "Beschreibung"],
    ["targetSheetName", CONFIG.targetSheetName, "Name des Ziel-Tabellenblatts für den Datenexport"],
    ["minEkThreshold", `${CONFIG.minEkThreshold} €`, "Mindesteinkaufspreis (fldEkRoh) für den Server-Filter"],
    ["targetWarehouses", CONFIG.targetWarehouses.join(", "), "Ausgewählte Lagernummern für die Spaltenpivotierung"],
    ["apiUrl", CONFIG.apiUrl, "Nepata Datahub GraphQL Endpoint"],
    ["Letzte Aktualisierung", new Date().toLocaleString("de-DE"), "Zeitstempel des letzten erfolgreichen Abrufs"],
    ["Gefundene Artikelzeilen", recordsCount, "Anzahl exportierter Artikel mit relevantem Bestand/Bestellungen"]
  ];

  const totalRows = infoData.length;
  const range = infoSheet.getRange(1, 1, totalRows, 3);
  range.setValues(infoData);

  // Formatierung des Info-Blatts
  infoSheet.getRange("A:C").setNumberFormat("@");
  
  const headerRange = infoSheet.getRange(1, 1, 1, 3);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  infoSheet.getRange(2, 1, totalRows - 1, 1).setFontWeight("bold");
  range.setBorder(true, true, true, true, true, true);
  
  infoSheet.autoResizeColumns(1, 3);
}

/**
 * Hilfsfunktion für sichere Toasts
 */
function safeToast(ss, message, title, timeoutSeconds = 3) {
  try {
    ss.toast(message, title, timeoutSeconds);
  } catch(e) {
    Logger.log(`Toast failed: ${title} - ${message}`);
  }
}