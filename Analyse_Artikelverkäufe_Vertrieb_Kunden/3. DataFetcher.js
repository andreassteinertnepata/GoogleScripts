// ==========================================
// 3. DATAFETCHER.GS - GraphQL Abruf & Kaufmännische Logik
// ==========================================

function fetchAndSaveRawData() {
  const activeTypes = getActiveVorgangsarten();
  const qBounds = getQuarterBoundaries();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let rawSheet = ss.getSheetByName(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    rawSheet = ss.insertSheet(CONFIG.SHEETS.RAW_DATA);
  } else {
    if (rawSheet.getFilter()) rawSheet.getFilter().remove();
    rawSheet.clear();
  }

  rawSheet.setFrozenRows(1); // FIXIERUNG HEADER

  rawSheet.getRange("D:D").setNumberFormat("@");
  rawSheet.getRange("I:I").setNumberFormat("@");
  
  rawSheet.appendRow(CONFIG.HEADERS_RAW);
  rawSheet.getRange(1, 1, 1, CONFIG.HEADERS_RAW.length)
          .setBackground("#2c3e50")
          .setFontColor("#ffffff")
          .setFontWeight("bold");

  const query = `
    query GetVertriebArchiv($cursor: String, $start24M: DateTime!, $vorgangsArten: [FilterValue!]!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $start24M } }] },
              { in: { field: fldArt, values: $vorgangsArten } }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldArt
              fldBelegNr
              fldDat
              fldStorniertKz
              fldVtrNr
              fldReNa2
              fldReNa3
              fldReLandBez
              rowAdresse {
                fldAbwArtDatGrp
              }
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                fldAbrPosKz
                rowArtikel {
                  fldKatalog
                  fldKuBez1
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
  let pageCount = 0;
  const rawRows = [];

  ss.toast("Lade Archivdaten der letzten 24 Monate...", "Datahub API", 8);

  while (hasNextPage && pageCount < CONFIG.API.MAX_PAGES) {
    pageCount++;
    const payload = {
      query: query,
      variables: {
        cursor: cursor,
        start24M: qBounds.isoStart24M,
        vorgangsArten: activeTypes
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API.TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API.URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors) {
        Logger.log("GraphQL Fehler: " + JSON.stringify(json.errors));
        break;
      }

      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      const edges = conRead.edges || [];
      if (edges.length === 0) break;

      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node || {};

        // 1. Storno-Regel: Stornierte Belege strikt ignorieren[cite: 17]
        if (node.fldStorniertKz === true) continue;

        const vtrNr = String(node.fldVtrNr || "").trim();
        const repInfo = CONFIG.REPS[vtrNr] || { name: "Unbekannt (" + vtrNr + ")" };
        
        // 2. Belegdatum (fldDat) für kaufmännischen Zeitraum[cite: 17]
        const recordDate = node.fldDat ? new Date(node.fldDat) : null;
        const quarterLabel = assignQuarterLabel(recordDate, qBounds);

        const kundenName = [node.fldReNa2, node.fldReNa3].filter(Boolean).join(" ").trim();
        const artCode = String(node.fldArt || "").trim();
        const belegNr = String(node.fldBelegNr || "").trim();
        const adrNrStr = String(node.fldAdrNr || "").trim();

        // Wiederverkäufer-Prüfung über fldAbwArtDatGrp
        const abwGrp = String(node.rowAdresse?.fldAbwArtDatGrp || "0").trim();
        const isReseller = (abwGrp !== "0" && abwGrp !== "");

        // 3. Vorzeichenlogik: Rechnungskorrekturen, Gutschriften, Erstattungen negativ berechnen[cite: 17]
        const isCreditNote = belegNr.startsWith("123") || CONFIG.CREDIT_TYPES.includes(artCode);

        const positions = node.rowsPositions || [];
        for (let j = 0; j < positions.length; j++) {
          const pos = positions[j] || {};
          if (pos.fldAbrPosKz !== true) continue;

          // 4. Katalog-Filter: Nur Katalog "1"
          const katalog = String(pos.rowArtikel?.fldKatalog || "").trim();
          if (katalog !== CONFIG.CATALOG_FILTER) continue;

          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;
          let artikelNrStr = String(pos.fldArtNr || "").trim();
          let artikelName = pos.rowArtikel?.fldKuBez1 || "";

          if (isCreditNote) {
            mge = -Math.abs(mge);
          }
          const nettoUmsatz = mge * eprNt;

          rawRows.push([
            recordDate, vtrNr, repInfo.name, adrNrStr, kundenName, node.fldReLandBez || "Unbekannt",
            belegNr, artCode, artikelNrStr, artikelName, katalog, mge, nettoUmsatz, quarterLabel, isReseller
          ]);
        }
      }
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler beim Abruf: " + e.toString());
      break;
    }
  }

  if (rawRows.length > 0) {
    rawSheet.getRange(2, 1, rawRows.length, CONFIG.HEADERS_RAW.length).setValues(rawRows);
    rawSheet.getRange(2, 1, rawRows.length, 1).setNumberFormat("yyyy-mm-dd");
    rawSheet.getRange(2, 12, rawRows.length, 1).setNumberFormat("#,##0");
    rawSheet.getRange(2, 13, rawRows.length, 1).setNumberFormat("#,##0.00 €");
    rawSheet.autoResizeColumns(1, CONFIG.HEADERS_RAW.length);
  }
}