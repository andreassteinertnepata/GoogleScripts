/**
 * Customers.gs
 * Fetches customer data and links aggregated annual sales totals (Current Year & Previous Year)
 * directly via a dedicated GraphQL query to tblVorgangArchiv using fastFilter.
 */

function fetchCustomers() {
  const BLATT_NAME = CONFIG.SHEET_CUSTOMERS || "Customers";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);

  if (sheet) {
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(BLATT_NAME);
  }

  const headers = [
    "AdrNr", "KredLimit", "VtrNr", 
    "Name2", "Name3", "Land", "EMail1", "EMail2",
    "Customer Type", "Sales Current Year", "Sales Previous Year"
  ];
  sheet.appendRow(headers);

  const currentYear = new Date().getFullYear();
  const previousYear = currentYear - 1;
  const startDateISO = previousYear + "-01-01T00:00:00Z";
  const vtrNr = CONFIG.VERTRETER_NR;

  Logger.log("Fetching archived sales data for representative " + vtrNr + " starting from " + startDateISO + "...");
  const salesMap = fetchCustomerSalesData(startDateISO, currentYear, previousYear, vtrNr);

  const queryAdressen = `
    query GetAdressen($cursor: String, $vtrNr: String!) {
      tblAdresse {
        conRead(
          first: 100, 
          after: $cursor, 
          orderBy: [{ field: fldLtzAend, desc: true }],
          fastFilter: {
            eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldKredLimit
              fldVtrNr
              fldAbwArtDatGrp
              rowsAnschriften {
                fldNa2
                fldNa3
                fldLandBez
                fldEMail1
                fldEMail2
              }
            }
          }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  const typeMapping = { "0": "End Customer", "1": "A", "2": "B", "3": "C" };
  let allRows = [];
  let hasNextPage = true;
  let cursor = null;
  let pageCount = 0;
  const MAX_PAGES = 500;
  const verarbeiteteKunden = new Set();
  
  const startTime = Date.now();
  const MAX_TIME_MS = 5.5 * 60 * 1000;

  while (hasNextPage && pageCount < MAX_PAGES) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit für Customers Adressen erreicht.");
      break;
    }
    pageCount++;

    const payload = { query: queryAdressen, variables: { cursor: cursor, vtrNr: vtrNr } };
    const options = {
      method: "post", contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      if (response.getResponseCode() !== 200) break;
      const json = JSON.parse(response.getContentText());
      if (json.errors) break;

      const conRead = json.data?.tblAdresse?.conRead || {};
      const edges = conRead.edges || [];
      if (edges.length === 0) break;

      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node || {};
        const adrNr = node.fldAdrNr ? node.fldAdrNr.toString().trim() : "";

        if (!adrNr || verarbeiteteKunden.has(adrNr)) continue;
        if (CONFIG.AUSGESCHLOSSENE_ADRESSEN && CONFIG.AUSGESCHLOSSENE_ADRESSEN.includes(adrNr)) continue;

        let abwGruppe = node.fldAbwArtDatGrp ? node.fldAbwArtDatGrp.toString().trim() : "";
        const mappedType = typeMapping[abwGruppe] || "";
        const anschriftenList = node.rowsAnschriften || [];
        const inner = anschriftenList[0] || {};
        const sales = salesMap.get(adrNr) || { currentYear: 0, previousYear: 0 };

        allRows.push([
          adrNr, node.fldKredLimit || 0, node.fldVtrNr || "",
          inner.fldNa2 || "", inner.fldNa3 || "", inner.fldLandBez || "",
          inner.fldEMail1 || "", inner.fldEMail2 || "",
          mappedType, sales.currentYear, sales.previousYear
        ]);
        verarbeiteteKunden.add(adrNr);
      }

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) { break; }
  }

  if (allRows.length > 0) {
    allRows.sort((a, b) => b[9] - a[9]);
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
    sheet.getRange(2, 10, allRows.length, 2).setNumberFormat('#,##0.00 "€"');
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("No customers found.", "Notice", 5);
  }
}

function fetchCustomerSalesData(startDateISO, currentYear, previousYear, vtrNr) {
  const salesMap = new Map();

  // WICHTIG: Regel 1 & 2 - fldSel14 und fldStorniertKz als Node abfragen. Regel 4: Filter auf fldDat
  const queryArchiv = `
    query GetSalesHistory($cursor: String, $jahrStart: DateTime!, $vtrNr: String!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $jahrStart } }] },
              { eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }] },
              { in: { 
                  field: fldArt, 
                  values: [
                    { string: "70" }, { string: "90" }, { string: "105" }, { string: "109" }, 
                    { string: "110" }, { string: "113" }, { string: "115" }, 
                    { string: "122" }, { string: "123" }, { string: "129" }, 
                    { string: "154" }, { string: "155" }, { string: "156" }
                  ] 
                } 
              }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldBelegNr
              fldArt
              fldDat
              fldStorniertKz
              fldSel14
              rowsPositions {
                fldMge
                fldEPrNt
                fldAbrPosKz
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
  const startTime = Date.now();
  const MAX_TIME_MS = 5.5 * 60 * 1000;

  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit für Customers History erreicht.");
      break;
    }
    
    const payload = { query: queryArchiv, variables: { cursor: cursor, jahrStart: startDateISO, vtrNr: vtrNr } };
    const options = {
      method: "post", contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      if (response.getResponseCode() !== 200) break;
      const json = JSON.parse(response.getContentText());
      if (json.errors) break;

      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      const edges = conRead.edges || [];
      if (edges.length === 0) break;

      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node || {};
        
        // REGEL 1 & 2: Testbelege und Stornos ignorieren
        if (node.fldSel14 === true) continue;
        if (node.fldStorniertKz === true) continue;
        
        const adrNr = node.fldAdrNr ? node.fldAdrNr.toString().trim() : "";
        const dateStr = node.fldDat || "";
        if (!adrNr || !dateStr) continue;

        const belegJahr = new Date(dateStr).getFullYear();
        const belegNr = String(node.fldBelegNr || "");
        
        // REGEL 3: Korrekturbelege erkennen
        const artCode = String(node.fldArt || "").trim();
        const isCreditNote = (artCode === "90" || artCode === "123" || artCode === "156" || belegNr.startsWith("123"));

        (node.rowsPositions || []).forEach(pos => {
          if (pos.fldAbrPosKz !== true) return; // REGEL 5: Nur Abrechnungspositionen

          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;

          if (isCreditNote) {
            mge = -Math.abs(mge);
            eprNt = Math.abs(eprNt);
          }

          const nettoGesamt = mge * eprNt;
          if (!salesMap.has(adrNr)) salesMap.set(adrNr, { currentYear: 0, previousYear: 0 });

          const record = salesMap.get(adrNr);
          if (belegJahr === currentYear) record.currentYear += nettoGesamt;
          else if (belegJahr === previousYear) record.previousYear += nettoGesamt;
        });
      }
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) { break; }
  }
  return salesMap;
}