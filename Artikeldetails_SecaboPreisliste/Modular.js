function fetchModularArtikel() {
  const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
  const API_TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
  const BLATT_NAME = "Modular";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);
  if (!sheet) sheet = ss.insertSheet(BLATT_NAME);
  
  // Vorherige Filter & Formate sauber entfernen
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  sheet.clear(); 
  sheet.clearFormats();

  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 330000; // Zeitbremse: 5,5 Minuten

  // =========================================================
  // PHASE 1: MODULARE ARTIKEL & LIEFERANTEN LADEN
  // =========================================================
  // Filtert nach Katalogen (0, 1, 8, 13, 18) UND ArtikelArt ungleich 0
  const queryArtikel = `
    query GetModularArtikel($cursor: String) {
      tblArtikel {
        conRead(
          first: 100, 
          after: $cursor,
          fastFilter: {
            and: [
              {
                in: {
                  field: fldKatalog,
                  values: [{ string: "1" }]
                }
              },
              {
                ne: [
                  { field: fldArtikelArt },
                  { value: { string: "0" } }
                ]
              }
            ]
          }
        ) {
          edges {
            node {
              fldArtNr
              fldKuBez1
              fldKuBez2
              fldKuBez3
              fldKuBez4
              fldArtikelArt
              lblArtikelArt
              fldGspKz
              fldKatalog
              lblKatalog
              fldBarCd
              fldLaenge
              fldBreite
              fldHoehe
              fldLaengeEinh
              fldGew
              fldSel16
              fldSel8
              fldEkKalk
              fldEkMitt
              fldEkRoh
              fldVk0Preis
              fldLiefNr
              rowsArtikelLieferanten {
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

  let hasNextPage = true;
  let cursor = null;
  let alleArtikel = [];

  ss.toast("Lade Modulare Artikel...", "Schritt 1/3");

  while (hasNextPage) {
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      ss.toast("Zeitlimit erreicht bei Artikel-Abfrage.", "Hinweis");
      break;
    }

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({ query: queryArtikel, variables: { cursor: cursor } }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("FEHLER: " + response.getContentText());
        ss.toast("Fehler beim Abrufen der Artikel!", "Fehler");
        return;
      }

      const conRead = json.data.tblArtikel.conRead || {};
      const edges = conRead.edges || [];
      edges.forEach(edge => alleArtikel.push(edge.node));

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) {
      Logger.log("Fehler: " + e.message);
      ss.toast("Fehler: " + e.message, "Fehler");
      return;
    }
  }

  if (alleArtikel.length === 0) {
    sheet.getRange(1, 1).setValue("Keine modularen Artikel gefunden.");
    return;
  }

  // =========================================================
  // PHASE 2: ABWEICHENDE PREISE LADEN
  // =========================================================
  const queryPreise = `
    query GetAbweichendePreise($cursor: String) {
      tblAbweichendeArtikeldaten {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { isNotNull: { field: fldGrp } },
              { isNull: { field: fldAdrNr } }
            ]
          }
        ) {
          edges {
            node {
              fldArtNr
              lblGrp
              fldAbwPr
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

  hasNextPage = true;
  cursor = null;
  let abweichendePreiseMap = {}; 

  ss.toast("Lade Reseller-Preise...", "Schritt 2/3");

  while (hasNextPage) {
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      ss.toast("Zeitlimit erreicht bei Preis-Abfrage.", "Hinweis");
      break;
    }

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({ query: queryPreise, variables: { cursor: cursor } }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        break;
      }

      const conRead = json.data.tblAbweichendeArtikeldaten.conRead || {};
      const edges = conRead.edges || [];
      
      edges.forEach(edge => {
        const node = edge.node;
        const artNr = node.fldArtNr;
        const categoryName = node.lblGrp || "";
        const price = node.fldAbwPr || 0;

        if (!abweichendePreiseMap[artNr]) {
          abweichendePreiseMap[artNr] = {};
        }
        abweichendePreiseMap[artNr][categoryName] = price;
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) {
      break;
    }
  }

  // =========================================================
  // PHASE 3: TABELLE ZUSAMMENBAUEN & FORMATIEREN
  // =========================================================
  ss.toast("Erstelle Tabelle...", "Schritt 3/3");

  const headers = [
    "fldArtNr", "fldKuBez1", "fldKuBez2", "fldKuBez3", "fldKuBez4", 
    "fldArtikelArt", "lblArtikelArt", "fldGspKz", "fldKatalog", "lblKatalog", 
    "fldBarCd", "fldLaenge", "fldBreite", "fldHoehe", "fldLaengeEinh", "fldGew", 
    "Ursprungsland (fldSel16)", "Zolltarifnr (fldSel8)", 
    "fldEkKalk", "fldEkMitt", "fldEkRoh", "fldVk0Preis", 
    "Lieferantennummer", "Lieferantenname 1", "Lieferantenname 2",
    "Secabo A", "Secabo B", "Secabo C"
  ];
  
  const zeilen = [headers];

  alleArtikel.forEach(artikel => {
    const artNr = artikel.fldArtNr || "";
    
    // 1. Reseller-Preise mappen
    const preiseObj = abweichendePreiseMap[artNr] || {};
    const preisA = preiseObj["Secabo Reseller A"] !== undefined ? preiseObj["Secabo Reseller A"] : "";
    const preisB = preiseObj["Secabo Reseller B"] !== undefined ? preiseObj["Secabo Reseller B"] : "";
    const preisC = preiseObj["Secabo Reseller C"] !== undefined ? preiseObj["Secabo Reseller C"] : "";

    // 2. Lieferantenname 1 & 2 auslesen
    let liefName1 = "";
    let liefName2 = "";
    const lieferanten = artikel.rowsArtikelLieferanten || [];
    
    if (lieferanten.length > 0) {
      const stdLief = lieferanten.find(l => l.fldStdKz === true) || lieferanten[0];
      if (stdLief && stdLief.rowAdresse && stdLief.rowAdresse.rowsAnschriften && stdLief.rowAdresse.rowsAnschriften.length > 0) {
        liefName1 = stdLief.rowAdresse.rowsAnschriften[0].fldNa1 || "";
        liefName2 = stdLief.rowAdresse.rowsAnschriften[0].fldNa2 || "";
      }
    }

    zeilen.push([
      artNr,
      artikel.fldKuBez1 || "",
      artikel.fldKuBez2 || "",
      artikel.fldKuBez3 || "",
      artikel.fldKuBez4 || "",
      artikel.fldArtikelArt || "",
      artikel.lblArtikelArt || artikel.fldArtikelArt || "",
      artikel.fldGspKz === true,
      artikel.fldKatalog || "",
      artikel.lblKatalog || artikel.fldKatalog || "",
      artikel.fldBarCd || "",
      artikel.fldLaenge || 0,
      artikel.fldBreite || 0,
      artikel.fldHoehe || 0,
      artikel.fldLaengeEinh || "",
      artikel.fldGew || 0,
      artikel.fldSel16 || "",
      artikel.fldSel8 || "",
      artikel.fldEkKalk || 0,
      artikel.fldEkMitt || 0,
      artikel.fldEkRoh || 0,
      artikel.fldVk0Preis || 0,
      artikel.fldLiefNr || "",
      liefName1,
      liefName2,
      preisA,
      preisB,
      preisC
    ]);
  });

  const totalRows = zeilen.length;
  const totalCols = headers.length;

  // 1. Textformate (@) für Codes und Nummern setzen
  sheet.getRange("A:A").setNumberFormat("@"); // ArtNr
  sheet.getRange("F:F").setNumberFormat("@"); // ArtikelArt
  sheet.getRange("I:I").setNumberFormat("@"); // Katalog
  sheet.getRange("K:K").setNumberFormat("@"); // Barcode
  sheet.getRange("Q:R").setNumberFormat("@"); // Ursprungsland & Zolltarif
  sheet.getRange("W:W").setNumberFormat("@"); // Lieferantennummer

  // 2. Daten eintragen
  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(zeilen);

  // 3. Zahlen- und Währungsformatierungen anwenden
  if (totalRows > 1) {
    sheet.getRange(2, 12, totalRows - 1, 3).setNumberFormat("#,##0.00"); // Maße (Spalten 12-14)
    sheet.getRange(2, 16, totalRows - 1, 1).setNumberFormat("#,##0.00"); // Gewicht (Spalte 16)
    
    sheet.getRange(2, 19, totalRows - 1, 4).setNumberFormat("#,##0.00 €"); // EK/VK Standard (Spalten 19-22)
    sheet.getRange(2, 26, totalRows - 1, 3).setNumberFormat("#,##0.00 €"); // Reseller A, B, C Preise (Spalten 26-28)
  }

  // 4. Header-Layout
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  // 5. Fixierungen (Zeile 1 & Spalte A einfrieren)
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1); 
  
  // 6. Auto-Filter setzen
  if (totalRows > 1) {
    range.createFilter();
  }

  // 7. Rahmen & Spaltenbreiten anpassen
  range.setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, totalCols);

  ss.toast("Modulare Artikel geladen & formatiert!", "Fertig");
}