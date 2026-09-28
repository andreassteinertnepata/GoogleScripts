function fetchVorgaenge() {
  const targetSheetName = CONFIG.SHEET_OPEN_ORDERS || "Open Orders"; 
  const repNo = String(CONFIG.VERTRETER_NR).trim();

  // Die Übersetzungsfunktion fängt diese Header später ab, wenn isEnglish wahr ist
  const HEADERS_DE = [
    "fldBelegNr", "fldAuftrNr", "fldArt", "Vorgangsart_Klartext", "fldVtrNr", "fldAdrNr", "fldDat", "fldLiefDat", "fldGspKz",
    "fldReNa2", "fldReNa3", "fldReLandBez", "fldLiNa2", "fldLiNa3", "fldLiStr", "fldLiPLZ", "fldLiOrt", "fldLiLandBez",
    "fldArtNr", "fldKuBez1", "fldKuBez3", "fldMge", "fldEPrNt", "Position Gesamtpreis Netto", "fldAusLagNr", "fldAbrPosKz"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vorgangsartenMap = new Map();
  
  try {
    const externalSs = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    const katalogSheet = externalSs.getSheetByName(CONFIG.TAB_VORGANGSARTEN);
    if (katalogSheet) {
      const katalogDaten = katalogSheet.getRange(2, 1, katalogSheet.getLastRow() - 1, 3).getValues();
      katalogDaten.forEach(row => {
        const code = String(row[0] || "").trim();
        const klartext = String(row[1] || "").trim();
        const kategorie = String(row[2] || "").trim(); 
        if (code !== "" && kategorie === "Vorgänge") vorgangsartenMap.set(code, klartext);
      });
    }
  } catch (e) { Logger.log("Error external sheet: " + e.message); }

  let vorgangsArten = Array.from(vorgangsartenMap.keys());
  if (vorgangsArten.length === 0) return;
  const formattedVorgangsArten = vorgangsArten.map(art => ({ string: art }));

  let zielSheet = ss.getSheetByName(targetSheetName);
  let isNewSheet = false;
  if (!zielSheet) { zielSheet = ss.insertSheet(targetSheetName); isNewSheet = true; } 
  if (zielSheet.getFilter()) zielSheet.getFilter().remove();
  
  const maxRows = zielSheet.getMaxRows();
  if (maxRows > 1) {
    zielSheet.getRange(2, 1, maxRows - 1, zielSheet.getMaxColumns()).clearContent();
    zielSheet.getRange(2, 1, maxRows - 1, zielSheet.getMaxColumns()).clearFormat();
  }

  if (isNewSheet) {
    zielSheet.getRange(1, 1, 1, HEADERS_DE.length).setValues([HEADERS_DE])
             .setBackground(CONFIG.FORMAT_HEADER_BG).setFontColor(CONFIG.FORMAT_HEADER_TEXT).setFontWeight("bold");
  }

  const query = `
    query GetVorgaenge($vorgangsArten: [FilterValue!]!, $vtrNr: String!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { in: { field: fldArt, values: $vorgangsArten } },
              { eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }] }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr fldArt fldAuftrNr fldBelegNr fldDat fldGspKz fldLiLandBez fldLiNa2 fldLiNa3 fldLiOrt fldLiPLZ fldLiStr fldLiefDat fldReLandBez fldReNa2 fldReNa3 fldVtrNr fldStorniertKz fldSel14
              rowsPositions {
                fldAbrPosKz fldArtNr fldAusLagNr fldEPrNt fldMge rowArtikel { fldKuBez1 fldKuBez3 }
              }
            }
          }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  let alleVorgaenge = [];
  let hasNextPage = true;
  let cursor = null;
  const startTime = Date.now();
  const MAX_TIME_MS = 5.5 * 60 * 1000;

  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit für Vorgaenge erreicht.");
      break;
    }
    const options = {
      method: "post", contentType: "application/json", headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({ query: query, variables: { vorgangsArten: formattedVorgangsArten, vtrNr: repNo, cursor: cursor } }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());
      if (json.errors) return;
      const conRead = json.data.tblVorgang?.conRead || {};
      (conRead.edges || []).forEach(edge => { if (edge.node) alleVorgaenge.push(edge.node); });
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) { return; }
  }

  const zeilen = [];

  alleVorgaenge.forEach(vorgang => {
    // REGEL 1 & 2: Testbelege und Stornos ignorieren (clientseitig)
    if (vorgang.fldSel14 === true) return;
    if (CONFIG.BUSINESS_LOGIC.isCancelled(vorgang)) return;
    
    const artCode = String(vorgang.fldArt || "").trim();
    const artKlartext = vorgangsartenMap.get(artCode) || "";

    const kopfDaten = [
      String(vorgang.fldBelegNr || ""), String(vorgang.fldAuftrNr || ""), artCode, artKlartext, String(vorgang.fldVtrNr || ""), String(vorgang.fldAdrNr || ""),
      vorgang.fldDat ? new Date(vorgang.fldDat) : "", vorgang.fldLiefDat ? new Date(vorgang.fldLiefDat) : "", vorgang.fldGspKz || false,
      String(vorgang.fldReNa2 || ""), String(vorgang.fldReNa3 || ""), String(vorgang.fldReLandBez || ""), String(vorgang.fldLiNa2 || ""),
      String(vorgang.fldLiNa3 || ""), String(vorgang.fldLiStr || ""), String(vorgang.fldLiPLZ || ""), String(vorgang.fldLiOrt || ""), String(vorgang.fldLiLandBez || "")
    ];

    const positionen = vorgang.rowsPositions || [];
    if (positionen.length > 0) {
      positionen.forEach(pos => {
        // REGEL 5: Nur echte Abrechnungspositionen berücksichtigen
        if (pos.fldAbrPosKz !== true) return;
        
        // REGEL 3: Vorzeichenlogik via CONFIG anwenden (Menge korrigieren, EP Netto absolut)
        const menge = CONFIG.BUSINESS_LOGIC.applySignLogic(artCode, pos.fldMge || 0);
        const einzelpreisNetto = Math.abs(pos.fldEPrNt || 0);

        zeilen.push([ ...kopfDaten, String(pos.fldArtNr || ""), String(pos.rowArtikel?.fldKuBez1 || ""), String(pos.rowArtikel?.fldKuBez3 || ""), menge, einzelpreisNetto, menge * einzelpreisNetto, String(pos.fldAusLagNr || ""), String(pos.fldAbrPosKz || "") ]);
      });
    } else {
      zeilen.push([ ...kopfDaten, "", "", "", 0, 0, 0, "", "" ]);
    }
  });

  if (zeilen.length > 0) {
    zielSheet.getRange("A:F").setNumberFormat("@");
    zielSheet.getRange("J:R").setNumberFormat("@");
    zielSheet.getRange("S:S").setNumberFormat("@");
    zielSheet.getRange("T:U").setNumberFormat("@");
    zielSheet.getRange("Y:Z").setNumberFormat("@");

    zielSheet.getRange(2, 1, zeilen.length, HEADERS_DE.length).setValues(zeilen);
    zielSheet.getRange(2, 7, zeilen.length, 2).setNumberFormat("dd.mm.yyyy");
    zielSheet.getRange(2, 22, zeilen.length, 1).setNumberFormat("#,##0");
    zielSheet.getRange(2, 23, zeilen.length, 2).setNumberFormat("#,##0.00 €");
  }
}