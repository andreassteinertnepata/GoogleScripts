/**
 * Test.gs - Diagnose-Funktion speziell für Artikel 109.010.11
 */

function testArticle109() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheetName = "Test_109_010_11";
  let sheet = ss.getSheetByName(testSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(testSheetName);
  } else {
    sheet.clear();
  }

  const targetArtNr = "109.010.11";
  const now = new Date();
  const dateFrom12M = new Date(now.getTime() - (400 * 24 * 60 * 60 * 1000)); // 400 Tage Puffer
  const dateFromIso = dateFrom12M.toISOString();

  const headers = [
    "Tabelle", "Beleg-Nr.", "Belegart", "Belegdatum", "Vertreter-ID", "Kunden-Nr.",
    "Artikel-Nr.", "Menge (fldMge)", "Einzelpreis (€)", "Gesamt (€)", 
    "Abrechnungspos? (fldAbrPosKz)", "Storniert? (fldStorniertKz)", "Sel14? (fldSel14)"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground("#1c4587").setFontColor("#ffffff").setFontWeight("bold");

  // Rechnungsbelegarten inklusive Vorkasse (129) und Gutschriften (123, 156)
  const allowedTypes = ["70", "105", "109", "110", "113", "115", "122", "123", "129", "154", "155", "156"];
  const formattedTypes = allowedTypes.map(function(art) { return { string: art }; });

  const query = `
    query GetTestArticle($cursor: String, $dateFrom: DateTime!, $vorgangsArten: [FilterValue!]!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldErstDat }, { value: { datetime: $dateFrom } }] },
              { in: { field: fldArt, values: $vorgangsArten } }
            ]
          }
        ) {
          pageInfo { hasNextPage endCursor }
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              fldVtrNr
              fldAdrNr
              fldStorniertKz
              fldSel14
              rowsPositions {
                fldMge
                fldEPrNt
                fldAbrPosKz
                fldArtNr
                rowArtikel {
                  fldKatalog
                  fldArtNr
                }
              }
            }
          }
        }
      }
    }
  `;

  let hasNextPage = true;
  let cursor = null;
  const rows = [];
  let sumMenge = 0;
  let sumUmsatz = 0;

  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { cursor: cursor, dateFrom: dateFromIso, vorgangsArten: formattedTypes }
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.errors) {
      Logger.log("GraphQL Error: " + JSON.stringify(json.errors));
      break;
    }

    const conRead = json.data && json.data.tblVorgangArchiv ? json.data.tblVorgangArchiv.conRead : null;
    if (!conRead || !conRead.edges) break;

    conRead.edges.forEach(function(edge) {
      const node = edge.node;
      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const belegDat = node.fldDat ? new Date(node.fldDat).toLocaleDateString("de-DE") : (node.fldErstDat ? new Date(node.fldErstDat).toLocaleDateString("de-DE") : "");
      const vtrNr = String(node.fldVtrNr || "");
      const adrNr = String(node.fldAdrNr || "");
      const storniert = node.fldStorniertKz === true;
      const sel14 = node.fldSel14 === true;

      (node.rowsPositions || []).forEach(function(pos) {
        const artNr = String(pos.fldArtNr || (pos.rowArtikel ? pos.rowArtikel.fldArtNr : "") || "").trim();
        if (artNr === targetArtNr) {
          let mge = pos.fldMge || 0;
          const epr = pos.fldEPrNt || 0;
          const abrPos = pos.fldAbrPosKz;

          const isKorrektur = CONFIG.CORRECTION_TYPES.includes(artCode) || belegNr.startsWith("123");
          if (isKorrektur) {
            mge = -Math.abs(mge);
          }

          const lineTotal = mge * epr;

          if (!storniert && !sel14 && (abrPos !== false)) {
            sumMenge += mge;
            sumUmsatz += lineTotal;
          }

          rows.push([
            "tblVorgangArchiv",
            belegNr,
            artCode,
            belegDat,
            vtrNr,
            adrNr,
            artNr,
            mge,
            epr,
            lineTotal,
            abrPos === null ? "NULL (ok)" : String(abrPos),
            storniert ? "JA (Ausschluss)" : "NEIN",
            sel14 ? "JA (Ausschluss)" : "NEIN"
          ]);
        }
      });
    });

    hasNextPage = conRead.pageInfo ? conRead.pageInfo.hasNextPage : false;
    cursor = conRead.pageInfo ? conRead.pageInfo.endCursor : null;
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 7).setNumberFormat("@");
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 9, rows.length, 2).setNumberFormat("#,##0.00 €");
    
    // Summenzeile schreiben
    const sumRow = ["GESAMTSUMME", "", "", "", "", "", targetArtNr, sumMenge, "", sumUmsatz, "", "", ""];
    sheet.appendRow(sumRow);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.getRange(lastRow, 8).setNumberFormat("#,##0");
    sheet.getRange(lastRow, 10).setNumberFormat("#,##0.00 €");

    sheet.autoResizeColumns(1, headers.length);
  } else {
    sheet.getRange(2, 1).setValue("Keine Belege für " + targetArtNr + " im Zeitraum gefunden.");
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Test beendet. Gefundene Gesamtmenge für ${targetArtNr}: ${sumMenge}`, "Test-Ergebnis", 10);
}