function fetchArticles() {
  const HEADERS = ["fldArtNr", "fldKuBez1", "fldKuBez3", "fldKatalog", "fldVk0Preis", "Lager_Menge", "Lager_KundenBestMenge", "WE_LiefDat", "WE_OffeneMenge"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_ARTICLES);
  let isNewSheet = false;
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_ARTICLES);
    isNewSheet = true;
  }
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

  const formattedKataloge = CONFIG.ARTIKEL_KATALOGE.map(k => ({ string: k }));
  const query = `
    query GetArticlesWithLagerAndWE($kataloge: [FilterValue!]!, $cursor: String) {
      tblArtikel {
        conRead(first: 100, after: $cursor, fastFilter: { in: { field: fldKatalog, values: $kataloge } }) {
          edges { node { fldArtNr fldKuBez1 fldKuBez3 fldKatalog fldVk0Preis rowsArtikelLager { fldArtNr fldLagNr fldMge fldKBstMge } rowsLieferantenbestelleingang { fldLiefDat fldOMge } } }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  let hasNextPage = true;
  let cursor = null;
  let alleArtikel = [];

  while (hasNextPage) {
    const options = {
      method: "post", contentType: "application/json", headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({ query: query, variables: { kataloge: formattedKataloge, cursor: cursor } }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());
      if (json.errors) return;
      const conRead = json.data.tblArtikel?.conRead || {};
      (conRead.edges || []).forEach(edge => { if (edge.node) alleArtikel.push(edge.node); });
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) { break; }
  }

  const zeilen = [];
  alleArtikel.forEach(art => {
    const kopf = [ String(art.fldArtNr || ""), String(art.fldKuBez1 || ""), String(art.fldKuBez3 || ""), String(art.fldKatalog || ""), art.fldVk0Preis || 0 ];
    let gesamtLagerMenge = 0;
    let gesamtKBstMge = 0;

    (art.rowsArtikelLager || []).forEach(l => {
      if (CONFIG.ERLAUBTE_LAGER.includes(String(l.fldLagNr || "").trim())) {
        gesamtLagerMenge += (l.fldMge || 0);
        gesamtKBstMge += (l.fldKBstMge || 0);
      }
    });

    const weList = art.rowsLieferantenbestelleingang || [];
    if (weList.length > 0) {
      weList.forEach(we => {
        zeilen.push([ ...kopf, gesamtLagerMenge, gesamtKBstMge, we.fldLiefDat ? new Date(we.fldLiefDat) : "", we.fldOMge || 0 ]);
      });
    } else {
      zeilen.push([ ...kopf, gesamtLagerMenge, gesamtKBstMge, "", 0 ]);
    }
  });

  if (zeilen.length > 0) {
    sheet.getRange("A:D").setNumberFormat("@");
    sheet.getRange(2, 1, zeilen.length, HEADERS.length).setValues(zeilen);
    sheet.getRange(2, 5, zeilen.length, 1).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 6, zeilen.length, 2).setNumberFormat("#,##0");
    sheet.getRange(2, 8, zeilen.length, 1).setNumberFormat("dd.mm.yyyy");
    sheet.getRange(2, 9, zeilen.length, 1).setNumberFormat("#,##0");
    // --- NEU: Sortierung nach Spalte 8 (H) aufsteigend ---
    sheet.getRange(2, 1, zeilen.length, HEADERS.length).sort({column: 8, ascending: true});
  }
}
