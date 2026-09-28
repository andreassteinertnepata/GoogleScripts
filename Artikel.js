/**
 * Artikel.js - Abfrage der Artikel aus Katalog 1 (EK >= 50 €) inklusive Verkaufsstückzahlen
 */

function fetchArticles(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_ARTIKEL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_ARTIKEL);
  }

  const baseHeaders = ["Artikel-Nr.", "Bezeichnung 1", "Katalog", "EK Roh (€)", "VK 0 (€)", "Menge Gesamt (12M)"];
  const repIds = Object.keys(CONFIG.SALES_REPS);
  const repHeaders = repIds.map(function(id) {
    return `Stk ${CONFIG.SALES_REPS[id].name}`;
  });

  const fullHeaders = baseHeaders.concat(repHeaders);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders])
       .setBackground("#1c4587").setFontColor("#ffffff").setFontWeight("bold");

  const query = `
    query GetCatalogArticles($cursor: String) {
      tblArtikel {
        conRead(
          first: ${CONFIG.PAGE_SIZE},
          after: $cursor,
          fastFilter: {
            eq: [{ field: fldKatalog }, { value: { string: "${CONFIG.TARGET_CATALOG}" } }]
          }
        ) {
          pageInfo { hasNextPage endCursor }
          edges {
            node {
              fldArtNr
              fldKuBez1
              fldKatalog
              fldEkRoh
              fldVk0Preis
            }
          }
        }
      }
    }
  `;

  let hasNextPage = true;
  let cursor = null;
  const articlesList = [];
  const rows = [];

  const totalArticleStats = (stats && stats._totalArticles) ? stats._totalArticles : {};

  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { cursor: cursor }
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.errors) break;

    const tableData = (json.data && json.data.tblArtikel) ? json.data.tblArtikel : null;
    const conRead = (tableData && tableData.conRead) ? tableData.conRead : null;
    if (!conRead || !conRead.edges) break;

    conRead.edges.forEach(function(edge) {
      const art = edge.node;
      const artNr = String(art.fldArtNr || "").trim();

      if (CONFIG.EXCLUDED_ARTICLES.includes(artNr)) return;

      const ekRoh = art.fldEkRoh || 0;
      if (ekRoh < CONFIG.MIN_EK) return;

      articlesList.push(art);

      // Gesamte Verkaufsmenge (12M) über ALLE Dokumente (inkl. Vorkasse 129 & unzugeordnete)
      const totArtStat = totalArticleStats[artNr];
      const total12M = totArtStat ? (totArtStat.stueck12M || 0) : 0;

      const repQtys = [];
      repIds.forEach(function(id) {
        const repArtData = (stats && stats[id] && stats[id].articles && stats[id].articles[artNr]) ? stats[id].articles[artNr] : null;
        const qty = repArtData ? (repArtData.stueck12M || 0) : 0;
        repQtys.push(qty);
      });

      const row = [
        artNr,
        String(art.fldKuBez1 || ""),
        String(art.fldKatalog || ""),
        ekRoh,
        art.fldVk0Preis || 0,
        total12M
      ].concat(repQtys);

      rows.push(row);
    });

    hasNextPage = conRead.pageInfo ? conRead.pageInfo.hasNextPage : false;
    cursor = conRead.pageInfo ? conRead.pageInfo.endCursor : null;
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, fullHeaders.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 4, rows.length, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 6, rows.length, repIds.length + 1).setNumberFormat("#,##0");
    sheet.autoResizeColumns(1, fullHeaders.length);
  }

  return articlesList;
}