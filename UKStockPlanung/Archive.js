function importArchivKomplett() {
  const HEADERS = [
    "Artikel-Nr.", "Artikel-Bez 1", "Artikel-Bez 3", 
    "Menge 30 Tage", "Menge 60 Tage", "Menge 90 Tage", "Menge 180 Tage"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVE);
  let isNewSheet = false; 
  
  if (!sheet) { 
    sheet = ss.insertSheet(CONFIG.SHEET_ARCHIVE); 
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
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setValues([HEADERS])
         .setBackground("#2c3e50")
         .setFontColor("#ffffff")
         .setFontWeight("bold");
  }

  // 1. Vorgangsarten laden
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
  } catch (e) { 
    Logger.log("Fehler beim Lesen der externen Vorgangsarten: " + e.toString());
    return; 
  }

  if (archivArten.length === 0) {
    Logger.log("Keine Archiv-Vorgangsarten im Master-Sheet gefunden.");
    return;
  }

  // 2. Datumsparameter festlegen (Maximaler Zeitraum: 180 Tage)
  const jetzt = new Date();
  const vor180Tagen = new Date(jetzt.getTime() - (180 * 24 * 60 * 60 * 1000));
  const START_DATUM = vor180Tagen.toISOString();
  const formattedVorgangsArten = archivArten.map(art => ({ string: art }));

  const query = `
    query GetArchivRobin($vorgangsArten: [FilterValue!]!, $cursor: String, $startDatum: DateTime!, $vtrNr: String!) {
      tblVorgangArchiv {
        conRead(
          first: 100, 
          after: $cursor, 
          fastFilter: { 
            and: [ 
              { ge: [{ field: fldErstDat }, { value: { datetime: $startDatum } }] }, 
              { in: { field: fldArt, values: $vorgangsArten } }, 
              { eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }] },
              { eq: [{ field: fldLiLandBez }, { value: { string: "Großbritannien" } }] }
            ] 
          }
        ) {
          edges { node { fldArt fldBelegNr fldDat rowsPositions { fldArtNr fldMge fldAbrPosKz rowArtikel { fldKuBez1 fldKuBez3 } } } }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  // Map zur Aggregation pro Artikel: Key = artNr -> Value = { bez1, bez3, mge30, mge60, mge90, mge180 }
  const artikelMap = new Map();

  let hasNextPage = true;
  let cursor = null;



  while (hasNextPage) {
    const options = {
      method: "post", 
      contentType: "application/json", 
      headers: { "X-API-Token": CONFIG.NEPATA_API_TOKEN },
      payload: JSON.stringify({ 
        query: query, 
        variables: { 
          vorgangsArten: formattedVorgangsArten, 
          cursor: cursor, 
          startDatum: START_DATUM, 
          vtrNr: CONFIG.VERTRETER_NR 
        } 
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.NEPATA_GRAPHQL_URL, options);
      const json = JSON.parse(response.getContentText());
      
      if (json.errors) {
        Logger.log("GraphQL Error in importArchivKomplett: " + JSON.stringify(json.errors));
        break;
      }

      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      (conRead.edges || []).forEach(edge => {
        const node = edge.node || {};
        const artCode = String(node.fldArt || "").trim();
        const belegNr = String(node.fldBelegNr || "").trim();
        const belegDatum = node.fldDat ? new Date(node.fldDat) : null;

        if (!belegDatum) return;

        // Alter des Belegs in Tagen ermitteln
        const alterInTagen = (jetzt.getTime() - belegDatum.getTime()) / (1000 * 60 * 60 * 24);

        (node.rowsPositions || []).forEach(pos => {
          if (pos.fldAbrPosKz !== true) return;

          const artNr = String(pos.fldArtNr || "").trim();
          if (!artNr) return;

          let mge = pos.fldMge || 0;
          if (artCode === "123" || belegNr.startsWith("123")) { 
            mge = -Math.abs(mge); 
          }

          // Eintrag in Map initialisieren falls noch nicht vorhanden
          if (!artikelMap.has(artNr)) {
            artikelMap.set(artNr, {
              bez1: String(pos.rowArtikel?.fldKuBez1 || ""),
              bez3: String(pos.rowArtikel?.fldKuBez3 || ""),
              mge30: 0,
              mge60: 0,
              mge90: 0,
              mge180: 0
            });
          }

          const item = artikelMap.get(artNr);

          // Kumulativ zu den passenden Zeitfenstern addieren
          if (alterInTagen <= 30) item.mge30 += mge;
          if (alterInTagen <= 60) item.mge60 += mge;
          if (alterInTagen <= 90) item.mge90 += mge;
          if (alterInTagen <= 180) item.mge180 += mge;
        });
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) { 
      Logger.log("Fehler im Fetch-Loop von importArchivKomplett: " + e.toString());
      break; 
    }
  }

  // 3. Aufbereitete Daten in Matrix für Sheet-Export umwandeln
  const allRows = [];
  artikelMap.forEach((data, artNr) => {
    allRows.push([
      artNr,
      data.bez1,
      data.bez3,
      data.mge30,
      data.mge60,
      data.mge90,
      data.mge180
    ]);
  });

  // 4. Daten eintragen & formatieren
  if (allRows.length > 0) {
    sheet.getRange("A:A").setNumberFormat("@"); 
    sheet.getRange(2, 1, allRows.length, HEADERS.length).setValues(allRows);
    
    // Mengen (Spalten D bis G / 4 bis 7) als ganze Zahl formatieren
    sheet.getRange(2, 4, allRows.length, 4).setNumberFormat("#,##0");
    
    Logger.log("Erfolgreich " + allRows.length + " Artikel-Summenzeilen importiert.");
  } else {
    Logger.log("Keine Archiv-Daten für Großbritannien in den letzten 180 Tagen gefunden.");
  }
}