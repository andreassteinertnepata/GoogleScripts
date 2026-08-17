/**
 * Optimized fetchArticles function
 * Filters server-side via GraphQL fastFilter for Catalogs AND Standard Articles (fldArtikelArt = "0").
 * Safely handles Google Apps Script 5-minute timeout protection.
 */
function fetchArticles() {
  const HEADERS = [
    "fldArtNr", "fldKuBez1", "fldKuBez3", "fldKatalog", "fldVk0Preis", 
    "Lager_Menge", "Lager_KundenBestMenge", "WE_LiefDat", "WE_OffeneMenge"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = CONFIG.SHEET_ARTICLES || "Articles";
  let sheet = ss.getSheetByName(targetSheetName);
  let isNewSheet = false;

  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
    isNewSheet = true;
  }

  // 1. Remove existing filter & clear contents starting row 2
  const currentFilter = sheet.getFilter();
  if (currentFilter) currentFilter.remove();

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, Math.max(maxCols, HEADERS.length)).clearContent();
  }

  if (isNewSheet) {
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setValues([HEADERS])
         .setBackground("#2c3e50")
         .setFontColor("#ffffff")
         .setFontWeight("bold");
  }

  // 2. Prepare GraphQL fastFilter Catalog Values
  const catalogList = CONFIG.ARTIKEL_KATALOGE || ["0", "1", "7", "8", "13", "18"];
  const formattedKataloge = catalogList.map(k => ({ string: String(k) }));

  // GraphQL Query: Filter by Catalog AND Standard Article Type ("0")
  const query = `
    query GetArticlesWithLagerAndWE($kataloge: [FilterValue!]!, $cursor: String) {
      tblArtikel {
        conRead(
          first: 100, 
          after: $cursor, 
          fastFilter: { 
            and: [
              { in: { field: fldKatalog, values: $kataloge } },
              { eq: [{ field: fldArtikelArt }, { value: { string: "0" } }] }
            ]
          }
        ) {
          edges { 
            node { 
              fldArtNr 
              fldKuBez1 
              fldKuBez3 
              fldKatalog 
              fldVk0Preis 
              rowsArtikelLager { 
                fldArtNr 
                fldLagNr 
                fldMge 
                fldKBstMge 
              } 
              rowsLieferantenbestelleingang { 
                fldLiefDat 
                fldOMge 
              } 
            } 
          }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  let hasNextPage = true;
  let cursor = null;
  const zeilen = [];

  // Timeout Protection (5 minutes)
  const startTime = Date.now();
  const MAX_TIME_MS = 5 * 60 * 1000;


  // 3. API Fetch & Inline Processing Loop
  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit fast erreicht. Bisher geladene Artikel werden geschrieben.");
      SpreadsheetApp.getActiveSpreadsheet().toast("Zeitlimit erreicht! Zwischenstand wird gespeichert...", "Achtung", 5);
      break;
    }

    const payload = {
      query: query,
      variables: { kataloge: formattedKataloge, cursor: cursor }
    };

    // WICHTIG: Verwende jetzt CONFIG.NEPATA_API_TOKEN und CONFIG.NEPATA_GRAPHQL_URL
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.NEPATA_API_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.NEPATA_GRAPHQL_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Error in fetchArticles: " + JSON.stringify(json.errors || response.getContentText()));
        break;
      }

      const conRead = json.data.tblArtikel?.conRead || {};
      const edges = conRead.edges || [];

      // Inline-Entpackung der Knoten
      edges.forEach(edge => {
        const art = edge.node;
        if (!art) return;

        const kopf = [
          String(art.fldArtNr || ""), 
          String(art.fldKuBez1 || ""), 
          String(art.fldKuBez3 || ""), 
          String(art.fldKatalog || ""), 
          art.fldVk0Preis || 0
        ];

        let gesamtLagerMenge = 0;
        let gesamtKBstMge = 0;

        // Lagerbestand für erlaubte Lager aus Config summieren
        const erlaubteLager = CONFIG.ERLAUBTE_LAGER || ["1", "2", "200", "13"];
        (art.rowsArtikelLager || []).forEach(l => {
          if (erlaubteLager.includes(String(l.fldLagNr || "").trim())) {
            gesamtLagerMenge += (l.fldMge || 0);
            gesamtKBstMge += (l.fldKBstMge || 0);
          }
        });

        // Offene Wareneingänge auflösen
        const weList = art.rowsLieferantenbestelleingang || [];
        if (weList.length > 0) {
          weList.forEach(we => {
            zeilen.push([
              ...kopf, 
              gesamtLagerMenge, 
              gesamtKBstMge, 
              we.fldLiefDat ? new Date(we.fldLiefDat) : "", 
              we.fldOMge || 0
            ]);
          });
        } else {
          zeilen.push([...kopf, gesamtLagerMenge, gesamtKBstMge, "", 0]);
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler im Fetch-Loop von fetchArticles: " + e.toString());
      break;
    }
  }

  // 4. In Sheet schreiben, Formate zuweisen & Sortieren
  if (zeilen.length > 0) {
    sheet.getRange("A:D").setNumberFormat("@");

    sheet.getRange(2, 1, zeilen.length, HEADERS.length).setValues(zeilen);

    sheet.getRange(2, 5, zeilen.length, 1).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 6, zeilen.length, 2).setNumberFormat("#,##0");
    sheet.getRange(2, 8, zeilen.length, 1).setNumberFormat("dd.mm.yyyy");
    sheet.getRange(2, 9, zeilen.length, 1).setNumberFormat("#,##0");

    sheet.getRange(2, 1, zeilen.length, HEADERS.length).sort({ column: 8, ascending: true });
    
    Logger.log("Erfolgreich " + zeilen.length + " Artikel-Zeilen geladen.");
  } else {
    Logger.log("Keine Artikel mit den angegebenen Filterkriterien gefunden.");
  }
}