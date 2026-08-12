/**
 * HAUPTFUNKTION: Steuert den gesamten Ablauf mit Toasts (Status-Meldungen unten rechts)
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Info-Meldung anzeigen
    ss.toast("Starte Stock- und Lagerdaten-Update...", "Bitte warten", -1);

    // Schritt 1: Stock-Daten von Secabo API importieren
    importSecaboStock();

    // Schritt 2: Lagerdaten via GraphQL abfragen, verarbeiten & Handlungsanweisung erstellen
    processStockAndWarehouseData();

    // Fertig-Meldung
    ss.toast("Stock und Lagerdaten erfolgreich aktualisiert!", "Erfolg", 5);

  } catch (e) {
    Logger.log("FEHLER in main(): " + e.toString());
    SpreadsheetApp.getUi().alert("Fehler bei der Ausführung: " + e.toString());
  }
}


/**
 * TEIL 1: Importiert die Stock-Daten von der Secabo Store API in das Blatt "Stock"
 */
function importSecaboStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ZIEL_BLATT = "Stock";
  
  ss.toast("Importiere Stock-Daten von Secabo...", "Schritt 1/2");

  let sheet = ss.getSheetByName(ZIEL_BLATT);
  if (!sheet) {
    sheet = ss.insertSheet(ZIEL_BLATT);
  } else {
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    sheet.clear();
  }

  // Formatierung festlegen
  sheet.getRange("A:A").setNumberFormat("@");

  const url = "https://www.secabo.com/store-api/nepata/018f1aad15167b07b14e9954409c9ef5/stock";
  const response = UrlFetchApp.fetch(url);
  const json = JSON.parse(response.getContentText());
  const data = json.data;
  
  // Exaktes 6-Spalten-Layout (ohne Update-Zeitstempel)
  let output = [["ID", "Artikelbezeichnung", "Available Stock", "Earliest Delivery", "Latest Delivery", "Handlungsanweisung"]];
  
  data.forEach(item => {
    output.push([
      item.id.toString(),
      "", // Artikelbezeichnung wird später durch GraphQL ergänzt
      item.availableStock || 0,
      item.deliveryDate ? item.deliveryDate.earliest : "-",
      item.deliveryDate ? item.deliveryDate.latest : "-",
      ""  // Handlungsanweisung wird später ergänzt
    ]);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}


/**
 * TEIL 2: GraphQL-Abfrage für Lagerbestände, Bezeichnungen & Handlungsanweisungen
 */
function processStockAndWarehouseData() {
  const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
  const API_TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
  
  const QUELL_BLATT = "Stock";
  const ZIEL_BLATT = "ArtikelLager";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Frage GraphQL-Lagerdaten ab...", "Schritt 2/2");

  const quellSheet = ss.getSheetByName(QUELL_BLATT);
  
  // 1. ArtikelLager-Blatt vorbereiten & leeren
  let zielSheet = ss.getSheetByName(ZIEL_BLATT);
  if (!zielSheet) {
    zielSheet = ss.insertSheet(ZIEL_BLATT);
  } else {
    const currentFilter = zielSheet.getFilter();
    if (currentFilter) currentFilter.remove();
    zielSheet.clear();
  }

  // 2. Artikelnummern aus "Stock" (Spalte A) holen
  const lastRowQuell = quellSheet.getLastRow();
  if (lastRowQuell < 2) return;

  const rawArtNrList = quellSheet.getRange(2, 1, lastRowQuell - 1, 1).getValues();
  const artikelNummern = [];
  rawArtNrList.forEach(row => {
    const artNr = String(row[0]).trim();
    if (artNr !== "" && !artikelNummern.includes(artNr)) {
      artikelNummern.push(artNr);
    }
  });

  if (artikelNummern.length === 0) return;

  // 3. GraphQL Query
  const query = `
    query GetArtikelLager($artNrs: [FilterValue!]!, $lagNrs: [FilterValue!]!, $cursor: String) {
      tblArtikelLager {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { in: { field: fldArtNr, values: $artNrs } },
              { in: { field: fldLagNr, values: $lagNrs } }
            ]
          }
        ) {
          edges {
            node {
              fldArtNr
              fldKBstMge
              fldLagNr
              fldMge
              rowArtikel {
                fldKatalog
                fldKuBez1
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

  const formattedArtNrs = artikelNummern.map(nr => ({ string: nr }));
  const erwaegteLager = ["1", "13", "2", "200", "202"];
  const formattedLagNrs = erwaegteLager.map(nr => ({ string: nr }));

  let alleErgebnisse = [];
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { artNrs: formattedArtNrs, lagNrs: formattedLagNrs, cursor: cursor }
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.errors || !json.data) {
      Logger.log("FEHLER GraphQL: " + JSON.stringify(json.errors));
      return;
    }

    const conRead = json.data.tblArtikelLager?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      if (edge.node) alleErgebnisse.push(edge.node);
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  // 4. Daten verarbeiten: Maps für Artikelnamen und Lager-Bestände
  const lagerDetailsMap = new Map();
  const artikelNamenMap = new Map();

  alleErgebnisse.forEach(row => {
    const artNr = String(row.fldArtNr || "").trim();
    const mge = Number(row.fldMge || 0);
    const lagNr = String(row.fldLagNr || "").trim();
    const bez = String(row.rowArtikel?.fldKuBez1 || "").trim();

    if (bez !== "") artikelNamenMap.set(artNr, bez);

    // Lager 1 bei der Handlungsanweisung komplett IGNORIEREN
    if (mge > 0 && lagNr !== "1") {
      if (!lagerDetailsMap.has(artNr)) {
        lagerDetailsMap.set(artNr, []);
      }
      lagerDetailsMap.get(artNr).push(`Lager ${lagNr}: ${mge} Stk.`);
    }
  });

  // 5. In "ArtikelLager" schreiben & formatieren (hier bleibt Lager 1 zur Übersicht enthalten)
  if (alleErgebnisse.length > 0) {
    const headers = ["fldArtNr", "fldKuBez1", "fldKatalog", "fldLagNr", "fldMge", "fldKBstMge"];
    const zeilen = [headers];

    alleErgebnisse.forEach(row => {
      zeilen.push([
        String(row.fldArtNr || ""),
        String(row.rowArtikel?.fldKuBez1 || ""),
        String(row.rowArtikel?.fldKatalog || ""),
        String(row.fldLagNr || ""),
        row.fldMge || 0,
        row.fldKBstMge || 0
      ]);
    });

    const totalRowsAL = zeilen.length;
    const totalColsAL = headers.length;

    zielSheet.getRange("A:A").setNumberFormat("@");
    zielSheet.getRange("B:B").setNumberFormat("@");
    zielSheet.getRange("C:C").setNumberFormat("@");
    zielSheet.getRange("D:D").setNumberFormat("@");

    const rangeAL = zielSheet.getRange(1, 1, totalRowsAL, totalColsAL);
    rangeAL.setValues(zeilen);

    if (totalRowsAL > 1) {
      zielSheet.getRange(2, 5, totalRowsAL - 1, 2).setNumberFormat("#,##0");
    }

    const headerRangeAL = zielSheet.getRange(1, 1, 1, totalColsAL);
    headerRangeAL.setBackground("#2c3e50");
    headerRangeAL.setFontColor("#ffffff");
    headerRangeAL.setFontWeight("bold");

    rangeAL.setBorder(true, true, true, true, true, true);
    rangeAL.createFilter();
    zielSheet.autoResizeColumns(1, totalColsAL);
  }

  // 6. Blatt "Stock" vervollständigen (Bezeichnungen + Handlungsanweisungen)
  const stockData = quellSheet.getRange(2, 1, lastRowQuell - 1, 3).getValues(); // Spalten A, B, C
  const bezeichnungenCol = [];
  const handlungsAnweisungenCol = [];

  stockData.forEach(row => {
    const artNr = String(row[0]).trim();
    const availableStock = Number(row[2] || 0); // Available Stock liegt in Spalte C

    bezeichnungenCol.push([artikelNamenMap.get(artNr) || ""]);

    // Handlungsanweisung erstellen (prüft nur Bestände abseits von Lager 1)
    if (availableStock === 0 && lagerDetailsMap.has(artNr)) {
      const detailsText = lagerDetailsMap.get(artNr).join(", ");
      handlungsAnweisungenCol.push([`Prüfen/Umbuchen: Menge vorhanden (${detailsText})`]);
    } else {
      handlungsAnweisungenCol.push([""]);
    }
  });

  // Bezeichnungen in Spalte B und Handlungsanweisungen in Spalte F schreiben
  quellSheet.getRange(2, 2, bezeichnungenCol.length, 1).setValues(bezeichnungenCol);
  quellSheet.getRange(2, 6, handlungsAnweisungenCol.length, 1).setValues(handlungsAnweisungenCol);

  // Layout & Design für "Stock"
  const totalColsStock = 6;
  const stockHeaderRange = quellSheet.getRange(1, 1, 1, totalColsStock);
  stockHeaderRange.setBackground("#2c3e50");
  stockHeaderRange.setFontColor("#ffffff");
  stockHeaderRange.setFontWeight("bold");

  quellSheet.getRange("A:A").setNumberFormat("@");
  quellSheet.getRange("B:B").setNumberFormat("@");
  quellSheet.getRange("C:C").setNumberFormat("#,##0");

  const stockFullRange = quellSheet.getRange(1, 1, lastRowQuell, totalColsStock);
  stockFullRange.setBorder(true, true, true, true, true, true);
  quellSheet.autoResizeColumns(1, totalColsStock);

  // NEU (Filtert nach Werten: Blendet starr nur leere Zellen aus,
  // lässt aber die Checkboxen für die Mitarbeiter aktiv!):
  
  // NEU: Nur leere Werte ("") explizit ausblenden.
  // Dadurch bleiben alle echten Texte sichtbar und die Checkboxen unter "Nach Werten filtern" aktiv!
  const filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([""]) 
    .build();

  const stockFilter = stockFullRange.createFilter();
  stockFilter.setColumnFilterCriteria(6, filterCriteria);
}