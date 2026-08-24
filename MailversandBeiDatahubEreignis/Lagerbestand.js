/**
 * ==========================================
 * LAGERBESTAND.JS
 * ==========================================
 * Fragt den aktuellen Lagerbestand (stock) aus rowCalc
 * sowie die Menge in Kundenbestellungen (fldKBstMge) aus rowsArtikelLager ab
 * und aktualisiert die Spalten E und F im Blatt "Bedingungen".
 */

function lagerbestandAktualisieren() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBedingungen = ss.getSheetByName("Bedingungen");

  if (!sheetBedingungen) {
    Logger.log("Fehler: Tabellenblatt 'Bedingungen' wurde nicht gefunden!");
    return;
  }

  // 1. ARTIKELLISTE LESEN (ZEILE 8 BIS ENDE)
  const lastRow = sheetBedingungen.getLastRow();
  if (lastRow < 8) {
    Logger.log("Keine Artikel in Zeile 8ff. gefunden.");
    return;
  }

  const artikelRange = sheetBedingungen.getRange(8, 1, lastRow - 7, 1);
  const artikelDaten = artikelRange.getValues();

  const artikelZeilenMap = new Map(); // artNr -> Zeilenindex im Sheet
  const artikelNummernListe = [];

  for (let i = 0; i < artikelDaten.length; i++) {
    const artNr = String(artikelDaten[i][0] || "").trim();
    if (artNr !== "") {
      artikelZeilenMap.set(artNr, i + 8);
      if (!artikelNummernListe.includes(artNr)) {
        artikelNummernListe.push(artNr);
      }
    }
  }

  if (artikelNummernListe.length === 0) {
    Logger.log("Keine gültigen Artikelnummern in Spalte A gefunden.");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Lade Lagerbestände vom Datahub...", "Bitte warten");

  // 2. GRAPHQL QUERY FÜR STOCK UND KBSTMGE
  const query = `
    query GetLagerbestaende($artNrs: [FilterValue!]!, $cursor: String) {
      tblArtikel {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            in: {
              field: fldArtNr,
              values: $artNrs
            }
          }
        ) {
          edges {
            node {
              fldArtNr
              rowCalc {
                stock
              }
              rowsArtikelLager {
                fldLagNr
                fldKBstMge
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

  const formattedArtNrs = artikelNummernListe.map(nr => ({ string: nr }));
  const ergebnisMap = new Map(); // artNr -> { stock, kundenbestellungen }

  let hasNextPage = true;
  let cursor = null;

  // 3. API-ABFRAGE MIT PAGINATION
  while (hasNextPage) {
    const payload = {
      query: query,
      variables: {
        artNrs: formattedArtNrs,
        cursor: cursor
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Error: " + JSON.stringify(json.errors || response.getContentText()));
        break;
      }

      const conRead = json.data?.tblArtikel?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const artNr = String(node.fldArtNr || "").trim();

        // Gesamtlagerbestand aus sidecar rowCalc
        const lagerbestand = node.rowCalc?.stock || 0;

        // Kundenbestellungen über alle relevanten Lager aufsummieren
        let kundenbestellungenSumme = 0;
        const lagerList = node.rowsArtikelLager || [];
        
        lagerList.forEach(lager => {
          // Erlaubte Lager aus CONFIG berücksichtigen, falls vorhanden
          const lagNr = String(lager.fldLagNr || "").trim();
          if (!CONFIG.ERLAUBTE_LAGER || CONFIG.ERLAUBTE_LAGER.includes(lagNr)) {
            kundenbestellungenSumme += Number(lager.fldKBstMge || 0);
          }
        });

        ergebnisMap.set(artNr, {
          stock: lagerbestand,
          kundenbestellungen: kundenbestellungenSumme
        });
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler bei Datahub-Abfrage (Lagerbestand): " + e.toString());
      break;
    }
  }

  // 4. DATEN IN SPALTEN E (Lagerbestand) UND F (Kundenbestellungen) SCHREIBEN
  artikelZeilenMap.forEach((zeilenIndex, artNr) => {
    const daten = ergebnisMap.get(artNr) || { stock: 0, kundenbestellungen: 0 };

    // Spalte E (Spalte 5): Aktueller Lagerbestand
    sheetBedingungen.getRange(zeilenIndex, 5).setValue(daten.stock);
    sheetBedingungen.getRange(zeilenIndex, 5).setNumberFormat("#,##0");

    // Spalte F (Spalte 6): in Kundenbestellungen
    sheetBedingungen.getRange(zeilenIndex, 6).setValue(daten.kundenbestellungen);
    sheetBedingungen.getRange(zeilenIndex, 6).setNumberFormat("#,##0");
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("Lagerbestände und Kundenbestellungen erfolgreich aktualisiert!", "Erfolg", 5);
}