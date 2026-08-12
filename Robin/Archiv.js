function importArchivKomplett() {
  const HEADERS = [
    "Adress-Nr.", "Vorgangsart", "Auftrags-Nr.", "Beleg-Nr.", "Datum", "Re-Name 2", "Re-Name 3", 
    "Rechnungsland", "Lieferland", "Zahlungsart", "Artikel-Nr.", "Artikel-Bez 1", "Artikel-Bez 3", 
    "Menge", "EP Netto", "Gesamt Netto"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVE);
  let isNewSheet = false; 
  if (!sheet) { sheet = ss.insertSheet(CONFIG.SHEET_ARCHIVE); isNewSheet = true; }
  const currentFilter = sheet.getFilter();
  if (currentFilter) currentFilter.remove();
  
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns()).clearContent();
    sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns()).clearFormat();
  }
  if (isNewSheet) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  }

  const vorgangsartenMap = new Map();
  const archivArten = [];
  try {
    const externalSs = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    const katalogSheet = externalSs.getSheetByName(CONFIG.TAB_VORGANGSARTEN);
    if (katalogSheet) {
      const katalogData = katalogSheet.getDataRange().getValues();
      for (let i = 1; i < katalogData.length; i++) {
        const code = String(katalogData[i][0]).trim();
        const archivWert = String(katalogData[i][2]).trim().toLowerCase();
        if (code !== "" && (archivWert === "archiv" || archivWert === "x" || archivWert === "ja" || katalogData[i][2] === true)) {
          vorgangsartenMap.set(code, String(katalogData[i][1]).trim());
          archivArten.push(code); 
        }
      }
    }
  } catch (e) { return; }

  if (archivArten.length === 0) return;
  const vor90Tagen = new Date(new Date().getTime() - (90 * 24 * 60 * 60 * 1000));
  const START_DATUM = vor90Tagen.toISOString();
  const formattedVorgangsArten = archivArten.map(art => ({ string: art }));

  const query = `
    query GetArchivRobin($vorgangsArten: [FilterValue!]!, $cursor: String, $startDatum: DateTime!, $vtrNr: String!) {
      tblVorgangArchiv {
        conRead(first: 100, after: $cursor, fastFilter: { and: [ { ge: [{ field: fldErstDat }, { value: { datetime: $startDatum } }] }, { in: { field: fldArt, values: $vorgangsArten } }, { eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }] } ] }) {
          edges { node { fldAdrNr fldArt fldAuftrNr fldBelegNr fldDat fldReNa2 fldReNa3 fldReLandBez fldLiLandBez fldZahlBed rowsPositions { fldArtNr fldMge fldEPrNt fldAbrPosKz rowArtikel { fldKuBez1 fldKuBez3 } } } }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  let allRows = [];
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const options = {
      method: "post", contentType: "application/json", headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({ query: query, variables: { vorgangsArten: formattedVorgangsArten, cursor: cursor, startDatum: START_DATUM, vtrNr: CONFIG.VERTRETER_NR } }),
      muteHttpExceptions: true
    };
    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());
      if (json.errors) break;
      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      (conRead.edges || []).forEach(edge => {
        const node = edge.node || {};
        const artCode = String(node.fldArt || "").trim();
        const belegNr = String(node.fldBelegNr || "").trim();
        (node.rowsPositions || []).forEach(pos => {
          if (pos.fldAbrPosKz !== true) return;
          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;
          if (artCode === "123" || belegNr.startsWith("123")) { eprNt = Math.abs(eprNt); mge = -Math.abs(mge); }
          
          allRows.push([
            String(node.fldAdrNr || ""), vorgangsartenMap.get(artCode) || "", String(node.fldAuftrNr || ""), belegNr, node.fldDat ? new Date(node.fldDat) : "",
            String(node.fldReNa2 || ""), String(node.fldReNa3 || ""), String(node.fldReLandBez || ""), String(node.fldLiLandBez || ""), String(node.fldZahlBed || ""),
            String(pos.fldArtNr || ""), String(pos.rowArtikel?.fldKuBez1 || ""), String(pos.rowArtikel?.fldKuBez3 || ""), mge, eprNt, mge * eprNt
          ]);
        });
      });
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) { break; }
  }

  if (allRows.length > 0) {
    sheet.getRange("A:D").setNumberFormat("@"); 
    sheet.getRange("K:K").setNumberFormat("@");
    sheet.getRange(2, 1, allRows.length, HEADERS.length).setValues(allRows);
    sheet.getRange(2, 5, allRows.length, 1).setNumberFormat("dd.mm.yyyy");
    sheet.getRange(2, 14, allRows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 15, allRows.length, 2).setNumberFormat("#,##0.00 €");
  }
}