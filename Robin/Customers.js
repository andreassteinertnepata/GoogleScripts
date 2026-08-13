/**
 * Customers.gs
 * Fetches customer data and links aggregated annual sales totals (Current Year & Previous Year)
 * directly via a dedicated GraphQL query to tblVorgangArchiv.
 */

function fetchCustomers() {
  const BLATT_NAME = CONFIG.SHEET_CUSTOMERS || "Customers";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);

  // 1. Prepare Sheet (Clear content & filters, no styling)
  if (sheet) {
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet(BLATT_NAME);
  }

  // Header definition (English, no "AbwArtDatGrp", "Land" instead of "LandBez", no "Vertreter")
  const headers = [
    "AdrNr", "KredLimit", "VtrNr", 
    "Name2", "Name3", "Land", "EMail1", "EMail2",
    "Customer Type", "Sales Current Year", "Sales Previous Year"
  ];
  sheet.appendRow(headers);

  // Determine current and previous calendar years dynamically
  const currentYear = new Date().getFullYear();
  const previousYear = currentYear - 1;
  const startDateISO = previousYear + "-01-01T00:00:00Z";

  // --- STEP 1: Fetch and aggregate historical sales via separate query ---
  Logger.log("Fetching archived sales data starting from " + startDateISO + "...");
  const salesMap = fetchCustomerSalesData(startDateISO, currentYear, previousYear);

  // --- STEP 2: Fetch Customer Addresses ---
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
          pageInfo {
            hasNextPage
            endCursor
          }
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

  Logger.log("Fetching customers for representative " + CONFIG.VERTRETER_NR + "...");

  while (hasNextPage && pageCount < MAX_PAGES) {
    pageCount++;

    const payload = {
      query: queryAdressen,
      variables: {
        cursor: cursor,
        vtrNr: CONFIG.VERTRETER_NR || "56"
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN || "e12Bfv!@Ss#asrpPFjucm8a8" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);

      if (response.getResponseCode() !== 200) {
        Logger.log("API Error: " + response.getContentText());
        break;
      }

      const json = JSON.parse(response.getContentText());
      if (json.errors) {
        Logger.log("GraphQL Error: " + JSON.stringify(json.errors));
        break;
      }

      const conRead = json.data?.tblAdresse?.conRead || {};
      const edges = conRead.edges || [];

      if (edges.length === 0) {
        hasNextPage = false;
        break;
      }

      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node || {};
        const adrNr = node.fldAdrNr ? node.fldAdrNr.toString().trim() : "";

        if (!adrNr || verarbeiteteKunden.has(adrNr)) {
          continue;
        }

        if (CONFIG.AUSGESCHLOSSENE_ADRESSEN && CONFIG.AUSGESCHLOSSENE_ADRESSEN.includes(adrNr)) {
          continue;
        }

        let abwGruppe = node.fldAbwArtDatGrp ? node.fldAbwArtDatGrp.toString().trim() : "";
        const mappedType = typeMapping[abwGruppe] || "";

        const anschriftenList = node.rowsAnschriften || [];
        const inner = anschriftenList[0] || {};

        // Link aggregated sales figures from Map
        const sales = salesMap.get(adrNr) || { currentYear: 0, previousYear: 0 };

        const rowData = [
          adrNr,
          node.fldKredLimit || 0,
          node.fldVtrNr || "",
          inner.fldNa2 || "",
          inner.fldNa3 || "",
          inner.fldLandBez || "", // Land
          inner.fldEMail1 || "",
          inner.fldEMail2 || "",
          mappedType,
          sales.currentYear,
          sales.previousYear
        ];

        allRows.push(rowData);
        verarbeiteteKunden.add(adrNr);
      }

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Error during customer fetch loop: " + e.toString());
      break;
    }
  }

  // --- STEP 3: Sort & Write data to sheet ---
  if (allRows.length > 0) {
    // Sort array descending by "Sales Current Year" (Index 9)
    allRows.sort((a, b) => b[9] - a[9]);

    // Write rows starting at row 2
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

    // Format sales columns (Column 10 & 11) as EUR currency
    sheet.getRange(2, 10, allRows.length, 2).setNumberFormat('#,##0.00 "€"');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Successfully loaded " + allRows.length + " customers with sales history!", 
      "Finished", 
      5
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("No customers found.", "Notice", 5);
  }
}

/**
 * Separate dedicated GraphQL query to tblVorgangArchiv to fetch full historical sales 
 * for current and previous years (bypassing the 90-day limitation of the local sheet).
 */
function fetchCustomerSalesData(startDateISO, currentYear, previousYear) {
  const salesMap = new Map();

  const queryArchiv = `
    query GetSalesHistory($cursor: String, $jahrStart: DateTime!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldErstDat }, { value: { datetime: $jahrStart } }] },
              { in: { 
                  field: fldArt, 
                  values: [
                    { string: "70" }, { string: "105" }, { string: "109" }, 
                    { string: "110" }, { string: "113" }, { string: "115" }, 
                    { string: "122" }, { string: "123" }, { string: "129" }, 
                    { string: "154" }, { string: "155" }
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
              fldDat
              rowsPositions {
                fldMge
                fldEPrNt
                fldAbrPosKz
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

  while (hasNextPage) {
    const payload = {
      query: queryArchiv,
      variables: {
        cursor: cursor,
        jahrStart: startDateISO
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN || "e12Bfv!@Ss#asrpPFjucm8a8" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
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
        const adrNr = node.fldAdrNr ? node.fldAdrNr.toString().trim() : "";
        if (!adrNr) continue;

        const dateStr = node.fldDat || "";
        if (!dateStr) continue;

        const belegJahr = new Date(dateStr).getFullYear();
        const belegNr = String(node.fldBelegNr || "");
        const positions = node.rowsPositions || [];

        for (let j = 0; j < positions.length; j++) {
          const pos = positions[j] || {};
          if (pos.fldAbrPosKz !== true) continue; // Only billed positions

          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;

          // Commercial sign logic for Credit Notes (Rechnungskorrektur 123)
          if (belegNr.startsWith("123")) {
            mge = -Math.abs(mge);
          }

          const nettoGesamt = mge * eprNt;

          if (!salesMap.has(adrNr)) {
            salesMap.set(adrNr, { currentYear: 0, previousYear: 0 });
          }

          const record = salesMap.get(adrNr);
          if (belegJahr === currentYear) {
            record.currentYear += nettoGesamt;
          } else if (belegJahr === previousYear) {
            record.previousYear += nettoGesamt;
          }
        }
      }

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Error fetching historical sales: " + e.toString());
      break;
    }
  }

  return salesMap;
}