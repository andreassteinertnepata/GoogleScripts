/**
 * ZENTRALE KONFIGURATION (Variablen)
 * Wenn du hier etwas änderst, ändert es sich im Skript UND im Info-Blatt!
 */
const CONFIG = {
  // URLs & Tokens
  SECABO_API_URL: "https://www.secabo.com/store-api/nepata/018f1aad15167b07b14e9954409c9ef5/stock",
  GRAPHQL_API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  GRAPHQL_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  
  // Tabellenblätter
  SHEET_INFO: "Info",
  SHEET_STOCK: "Stock",
  SHEET_LAGER: "ArtikelLager",
  
  // Logik-Parameter
  PFLICHT_ARTIKEL: ["109.000.10"], // Artikel, die immer abgefragt werden (auch wenn sie bei Secabo fehlen)
  ERWAEGTE_LAGER: ["1", "13", "2", "200", "202"], // Läger, die via GraphQL abgefragt werden
  LAGER_IGNORIERT: "1", // Lager, das bei der Handlungsanweisung (Mangel) ignoriert wird
  CHUNK_SIZE: 50 // Blockgröße für GraphQL-Abfragen (verhindert API-Limits)
};

/**
 * HAUPTFUNKTION: Steuert den gesamten Ablauf mit Toasts
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Info-Meldung anzeigen
    ss.toast("Starte Stock- und Lagerdaten-Update...", "Bitte warten", -1);

    // Schritt 0: Info-Blatt generieren/aktualisieren
    createInfoSheet();

    // Schritt 1: Stock-Daten von Secabo API importieren
    importSecaboStock();

    // Schritt 2: Lagerdaten via GraphQL abfragen, verarbeiten & Handlungsanweisung erstellen
    processStockAndWarehouseData();

    // Fertig-Meldung
    ss.toast("Stock, Lagerdaten und Info-Blatt erfolgreich aktualisiert!", "Erfolg", 5);

  } catch (e) {
    Logger.log("FEHLER in main(): " + e.toString());
    SpreadsheetApp.getUi().alert("Fehler bei der Ausführung: " + e.toString());
  }
}

/**
 * SCHRITT 0: Generiert das Info-Blatt und beschreibt den Prozess dynamisch
 */
function createInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Erstelle Info-Blatt...", "Schritt 0/3");

  let sheet = ss.getSheetByName(CONFIG.SHEET_INFO);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_INFO, 0); // An erster Stelle einfügen
  } else {
    sheet.clear();
    // Verschiebe es an die erste Stelle, falls es das nicht ist
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }

  // --- Inhalt für das Info-Blatt definieren ---
  const infoData = [
    ["System-Information & Ablauf-Dokumentation", "", ""],
    ["", "", ""],
    ["Was macht dieses Skript?", "", ""],
    ["Dieses Tool synchronisiert den Secabo-Shop-Bestand mit den tatsächlichen ERP-Lagerbeständen (via GraphQL).", "", ""],
    ["Ziel ist es, Handlungsanweisungen (z.B. 'Umbuchen') für Artikel zu generieren, die im Shop auf 0 stehen, aber physisch in anderen Lägern noch vorhanden sind.", "", ""],
    ["", "", ""],
    ["Schritt-für-Schritt Ablauf:", "", ""],
    ["1. Import Secabo", "Das Skript lädt aktuelle Shop-Daten ('Available Stock', 'Delivery Dates') in das Blatt '" + CONFIG.SHEET_STOCK + "'.", ""],
    ["2. GraphQL Abfrage", "Alle gefundenen Artikelnummern (plus Pflicht-Artikel) werden in " + CONFIG.CHUNK_SIZE + "er-Blöcken an die Cloud-API gesendet.", ""],
    ["3. Auswertung", "Die Roh-Lagerdaten landen im Blatt '" + CONFIG.SHEET_LAGER + "'.", ""],
    ["4. Handlungsanweisung", "Im Blatt '" + CONFIG.SHEET_STOCK + "' (Spalte F) wird geprüft: Ist Bestand = 0, aber Ware in Lägern (außer Lager " + CONFIG.LAGER_IGNORIERT + ") vorhanden? -> Anweisung generieren.", ""],
    ["", "", ""],
    ["Aktuell verwendete Parameter (Dynamisch aus dem Skript):", "", ""],
    ["Parameter / Variable", "Eingestellter Wert", "Beschreibung"],
    ["Secabo API URL", CONFIG.SECABO_API_URL, "Datenquelle für Shop-Bestände"],
    ["GraphQL API URL", CONFIG.GRAPHQL_API_URL, "Datenquelle für ERP/Lager-Bestände"],
    ["Haupt-Arbeitsblatt", CONFIG.SHEET_STOCK, "Blatt für die finale Übersicht inkl. Anweisungen"],
    ["Rohdaten-Arbeitsblatt", CONFIG.SHEET_LAGER, "Blatt für die reine Lager-Datenablage"],
    ["Abgefragte Läger", CONFIG.ERWAEGTE_LAGER.join(", "), "Nur Bestände aus diesen Lägern werden geholt"],
    ["Ignoriertes Lager", "Lager " + CONFIG.LAGER_IGNORIERT, "Dieses Lager zählt nicht für die Handlungsanweisung 'Umbuchen'"],
    ["GraphQL Blockgröße", CONFIG.CHUNK_SIZE + " Artikel pro Request", "Verhindert Server-Überlastung / Timeouts"],
    ["Force-Include Artikel", CONFIG.PFLICHT_ARTIKEL.length > 0 ? CONFIG.PFLICHT_ARTIKEL.join(", ") : "Keine", "Artikel, die hartkodiert abgefragt werden, falls sie bei Secabo fehlen"]
  ];

  // Daten in das Blatt schreiben
  sheet.getRange(1, 1, infoData.length, 3).setValues(infoData);

  // --- Formatierung & Design ---
  
  // Hauptüberschrift
  sheet.getRange("A1:C1").merge().setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
  
  // Zwischenüberschriften
  const subHeaders = ["A3:C3", "A7:C7", "A13:C13"];
  subHeaders.forEach(range => {
    sheet.getRange(range).merge().setBackground("#ecf0f1").setFontWeight("bold").setFontSize(11);
  });

  // Parameter-Tabelle Header
  const paramHeaderRange = sheet.getRange("A14:C14");
  paramHeaderRange.setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold");

  // Parameter-Tabelle Rahmen (ab Zeile 14 bis Ende)
  const paramTableRange = sheet.getRange(14, 1, infoData.length - 13, 3);
  paramTableRange.setBorder(true, true, true, true, true, true);

  // Spaltenbreiten anpassen (Word Wrap für längere Texte aktivieren)
  sheet.getRange("A:C").setWrap(true).setVerticalAlignment("middle");
  sheet.setColumnWidth(1, 200); // Parameter-Name
  sheet.setColumnWidth(2, 450); // Wert
  sheet.setColumnWidth(3, 350); // Beschreibung

  // Gitternetzlinien verstecken für cleanen Look
  sheet.setHiddenGridlines(true);
}

/**
 * TEIL 1: Importiert die Stock-Daten von der Secabo Store API in das Blatt "Stock"
 */
function importSecaboStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Importiere Stock-Daten von Secabo...", "Schritt 1/3");

  let sheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_STOCK);
  } else {
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    sheet.clear();
  }

  // Formatierung festlegen
  sheet.getRange("A:A").setNumberFormat("@");

  const response = UrlFetchApp.fetch(CONFIG.SECABO_API_URL);
  const json = JSON.parse(response.getContentText());
  const data = json.data;
  
  let output = [["ID", "Artikelbezeichnung", "Available Stock", "Earliest Delivery", "Latest Delivery", "Handlungsanweisung"]];
  
  if (data && Array.isArray(data)) {
    data.forEach(item => {
      output.push([
        item.id ? item.id.toString() : "",
        "", 
        item.availableStock || 0,
        item.deliveryDate ? item.deliveryDate.earliest : "-",
        item.deliveryDate ? item.deliveryDate.latest : "-",
        ""  
      ]);
    });
  }

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

/**
 * TEIL 2: GraphQL-Abfrage für Lagerbestände, Bezeichnungen & Handlungsanweisungen
 */
function processStockAndWarehouseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Frage GraphQL-Lagerdaten ab...", "Schritt 2/3");

  const quellSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
  if (!quellSheet) return;
  
  // 1. ArtikelLager-Blatt vorbereiten & leeren
  let zielSheet = ss.getSheetByName(CONFIG.SHEET_LAGER);
  if (!zielSheet) {
    zielSheet = ss.insertSheet(CONFIG.SHEET_LAGER);
  } else {
    const currentFilter = zielSheet.getFilter();
    if (currentFilter) currentFilter.remove();
    zielSheet.clear();
  }

  // 2. Artikelnummern sammeln
  const lastRowQuell = quellSheet.getLastRow();
  const artikelNummern = [];

  if (lastRowQuell >= 2) {
    const rawArtNrList = quellSheet.getRange(2, 1, lastRowQuell - 1, 1).getValues();
    rawArtNrList.forEach(row => {
      const artNr = String(row[0]).trim();
      if (artNr !== "" && !artikelNummern.includes(artNr)) {
        artikelNummern.push(artNr);
      }
    });
  }

  // Force-Include
  CONFIG.PFLICHT_ARTIKEL.forEach(artNr => {
    if (!artikelNummern.includes(artNr)) {
      artikelNummern.push(artNr);
      quellSheet.appendRow([artNr, "Lade...", 0, "-", "-", ""]);
    }
  });

  if (artikelNummern.length === 0) return;

  const formattedLagNrs = CONFIG.ERWAEGTE_LAGER.map(nr => ({ string: nr }));
  let alleErgebnisse = [];
  
  // 3. Chunking der GraphQL-Abfrage
  for (let i = 0; i < artikelNummern.length; i += CONFIG.CHUNK_SIZE) {
    const chunk = artikelNummern.slice(i, i + CONFIG.CHUNK_SIZE);
    const formattedArtNrs = chunk.map(nr => ({ string: nr }));
    
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
            edges { node { fldArtNr fldKBstMge fldLagNr fldMge rowArtikel { fldKatalog fldKuBez1 } } }
            pageInfo { hasNextPage endCursor }
          }
        }
      }
    `;

    let hasNextPage = true;
    let cursor = null;

    while (hasNextPage) {
      const options = {
        method: "post",
        contentType: "application/json",
        headers: { "X-API-Token": CONFIG.GRAPHQL_TOKEN },
        payload: JSON.stringify({ query: query, variables: { artNrs: formattedArtNrs, lagNrs: formattedLagNrs, cursor: cursor } }),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(CONFIG.GRAPHQL_API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("FEHLER GraphQL: " + JSON.stringify(json.errors || json));
        break;
      }

      const conRead = json.data.tblArtikelLager?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => { if (edge.node) alleErgebnisse.push(edge.node); });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    }
  }

  // 4. Daten verarbeiten
  const lagerDetailsMap = new Map();
  const artikelNamenMap = new Map();

  alleErgebnisse.forEach(row => {
    const artNr = String(row.fldArtNr || "").trim();
    const mge = Number(row.fldMge || 0);
    const lagNr = String(row.fldLagNr || "").trim();
    const bez = String(row.rowArtikel?.fldKuBez1 || "").trim();

    if (bez !== "") artikelNamenMap.set(artNr, bez);

    // Ausgeschlossenes Lager ignorieren
    if (mge > 0 && lagNr !== CONFIG.LAGER_IGNORIERT) {
      if (!lagerDetailsMap.has(artNr)) lagerDetailsMap.set(artNr, []);
      lagerDetailsMap.get(artNr).push(`Lager ${lagNr}: ${mge} Stk.`);
    }
  });

  // 5. In "ArtikelLager" schreiben
  if (alleErgebnisse.length > 0) {
    const headers = ["fldArtNr", "fldKuBez1", "fldKatalog", "fldLagNr", "fldMge", "fldKBstMge"];
    const zeilen = [headers];

    alleErgebnisse.forEach(row => {
      zeilen.push([
        String(row.fldArtNr || ""), String(row.rowArtikel?.fldKuBez1 || ""), String(row.rowArtikel?.fldKatalog || ""),
        String(row.fldLagNr || ""), row.fldMge || 0, row.fldKBstMge || 0
      ]);
    });

    const totalRowsAL = zeilen.length;
    const totalColsAL = headers.length;

    zielSheet.getRange("A:D").setNumberFormat("@");
    const rangeAL = zielSheet.getRange(1, 1, totalRowsAL, totalColsAL);
    rangeAL.setValues(zeilen);

    if (totalRowsAL > 1) zielSheet.getRange(2, 5, totalRowsAL - 1, 2).setNumberFormat("#,##0");

    zielSheet.getRange(1, 1, 1, totalColsAL).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
    rangeAL.setBorder(true, true, true, true, true, true);
    rangeAL.createFilter();
    zielSheet.autoResizeColumns(1, totalColsAL);
  }

  // 6. Blatt "Stock" vervollständigen
  const newLastRowQuell = quellSheet.getLastRow();
  if (newLastRowQuell >= 2) {
    const stockData = quellSheet.getRange(2, 1, newLastRowQuell - 1, 3).getValues(); 
    const bezeichnungenCol = [];
    const handlungsAnweisungenCol = [];

    stockData.forEach(row => {
      const artNr = String(row[0]).trim();
      const availableStock = Number(row[2] || 0); 

      bezeichnungenCol.push([artikelNamenMap.get(artNr) || ""]);

      if (availableStock === 0 && lagerDetailsMap.has(artNr)) {
        const detailsText = lagerDetailsMap.get(artNr).join(", ");
        handlungsAnweisungenCol.push([`Prüfen/Umbuchen: Menge vorhanden (${detailsText})`]);
      } else {
        handlungsAnweisungenCol.push([""]);
      }
    });

    quellSheet.getRange(2, 2, bezeichnungenCol.length, 1).setValues(bezeichnungenCol);
    quellSheet.getRange(2, 6, handlungsAnweisungenCol.length, 1).setValues(handlungsAnweisungenCol);

    const stockHeaderRange = quellSheet.getRange(1, 1, 1, 6);
    stockHeaderRange.setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
    quellSheet.getRange("A:B").setNumberFormat("@");
    quellSheet.getRange("C:C").setNumberFormat("#,##0");

    const stockFullRange = quellSheet.getRange(1, 1, newLastRowQuell, 6);
    stockFullRange.setBorder(true, true, true, true, true, true);
    quellSheet.autoResizeColumns(1, 6);

    const currentStockFilter = quellSheet.getFilter();
    if (currentStockFilter) currentStockFilter.remove();
    
    const filterCriteria = SpreadsheetApp.newFilterCriteria().setHiddenValues([""]).build();
    const stockFilter = stockFullRange.createFilter();
    stockFilter.setColumnFilterCriteria(6, filterCriteria);
  }
}