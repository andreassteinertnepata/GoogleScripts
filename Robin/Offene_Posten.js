/**
 * OffenePosten.gs
 * Fetches outstanding invoices (Offene Posten) filtered by representative (fldVtrNr)
 * using a 2-step fastFilter approach for maximum performance.
 */
function fetchOffenePosten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vtrNr = String(CONFIG.VERTRETER_NR || "56").trim();
  const targetSheetName = CONFIG.SHEET_OUTSTANDING || "Outstanding Invoices";

  // Check language based on representative number
  const englishReps = ["56", "60"]; // Robin Carter Browne, Rado Kabakov
  const isEnglish = englishReps.includes(vtrNr);

  // Defined Column Headers
  const HEADERS = isEnglish ? [
    "Rep_No", "Credit_Limit", "Cust_No", "Status",
    "Name_2", "Name_3", "Country",
    "Doc_No", "Order_No", "OP_Text",
    "Issue_Date", "Net_Days", "Paid_Amount",
    "OP_Balance_Amount", "Dunning_Level", "Dunning_Date",
    "Due_Date", "Days_Overdue"
  ] : [
    "Vertreter-Nr.", "Kreditlimit", "Adress-Nr. (Kunde)", "Status",
    "Name 2", "Name 3", "Land",
    "fldBelegNr", "fldAuftrNr", "OP-Text",
    "Erstelldatum", "Netto Tage", "Bezahlt-Betrag",
    "OP-Saldo Betrag", "Mahnstufe", "Mahndatum",
    "Fällig?", "Fällig seit (Tagen)"
  ];

  // --- STEP 1: Fetch Customer Numbers (AdrNr) for target Representative ---
  const customerMap = getRepresentativeCustomers(vtrNr);

  if (customerMap.size === 0) {
   
    return;
  }

  // Convert customer numbers for GraphQL fastFilter (`in`)
  const adrNrFilterValues = Array.from(customerMap.keys()).map(adrNr => ({ string: adrNr }));

  // --- STEP 2: Prepare Target Sheet ---
  let sheet = ss.getSheetByName(targetSheetName);
  let isNewSheet = false;

  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
    isNewSheet = true;
  }

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
         .setBackground(CONFIG.FORMAT_HEADER_BG || "#2c3e50")
         .setFontColor(CONFIG.FORMAT_HEADER_TEXT || "#ffffff")
         .setFontWeight("bold");
  }

  // --- STEP 3: Query Offene Posten with fastFilter (`in: fldAdrNr`) ---
  const queryOffenePosten = `
    query GetOffenePostenFiltered($adrNrs: [FilterValue!]!, $cursor: String) {
      tblOffenerPosten {
        conRead(
          first: 100, 
          after: $cursor,
          fastFilter: {
            in: { field: fldAdrNr, values: $adrNrs }
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldBelegNr
              fldAuftrNr
              fldText
              fldErstDat
              fldNettoTg
              fldBezBet
              fldOPSaldoBet
              fldMahnSt
              fldMahnDat
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
  const allRows = [];
  const heute = new Date();
  heute.setHours(0, 0, 0, 0);

  const startTime = Date.now();
  const MAX_TIME_MS = 5 * 60 * 1000; // 5 minute safety threshold



  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit erreicht. Speichere geladene Offene Posten...");
    
      break;
    }

    const payload = {
      query: queryOffenePosten,
      variables: {
        adrNrs: adrNrFilterValues,
        cursor: cursor
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Error in fetchOffenePosten: " + JSON.stringify(json.errors || response.getContentText()));
        break;
      }

      const conRead = json.data.tblOffenerPosten?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const adrNr = String(node.fldAdrNr || "").trim();
        
        // Fetch cached address details
        const custInfo = customerMap.get(adrNr);
        if (!custInfo) return;

        const erstDatStr = node.fldErstDat || "";
        const nettoTg = parseInt(node.fldNettoTg, 10) || 0;
        let faelligkeitStr = "";
        let faelligSeitTage = "";

        if (erstDatStr) {
          const erstelldatum = new Date(erstDatStr);
          if (!isNaN(erstelldatum.getTime())) {
            const faelligkeitsDatum = new Date(erstelldatum);
            faelligkeitsDatum.setDate(faelligkeitsDatum.getDate() + nettoTg);
            
            faelligkeitStr = Utilities.formatDate(faelligkeitsDatum, Session.getScriptTimeZone(), "yyyy-MM-dd");
            
            faelligkeitsDatum.setHours(0, 0, 0, 0);
            if (faelligkeitsDatum < heute) {
              const diffTime = heute.getTime() - faelligkeitsDatum.getTime();
              faelligSeitTage = Math.floor(diffTime / (1000 * 60 * 60 * 24));
            }
          }
        }

        allRows.push([
          custInfo.vtrNr,              // Spalte A
          custInfo.kredLimit,          // Spalte B
          adrNr,                       // Spalte C
          custInfo.status,             // Spalte D
          custInfo.name2,              // Spalte E
          custInfo.name3,              // Spalte F
          custInfo.land,               // Spalte G
          String(node.fldBelegNr || ""), // Spalte H
          String(node.fldAuftrNr || ""),// Spalte I
          String(node.fldText || ""),   // Spalte J
          erstDatStr ? new Date(erstDatStr) : "", // Spalte K
          nettoTg,                     // Spalte L
          node.fldBezBet || 0,         // Spalte M
          node.fldOPSaldoBet || 0,     // Spalte N
          node.fldMahnSt || 0,         // Spalte O
          node.fldMahnDat ? new Date(node.fldMahnDat) : "", // Spalte P
          faelligkeitStr,              // Spalte Q
          faelligSeitTage              // Spalte R
        ]);
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler bei fetchOffenePosten: " + e.message);
      break;
    }
  }

  // --- STEP 4: Write & Format Sheet ---
  if (allRows.length > 0) {
    // Force Text formatting for ID columns
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("C:D").setNumberFormat("@");
    sheet.getRange("H:I").setNumberFormat("@");

    // Write Values
    sheet.getRange(2, 1, allRows.length, HEADERS.length).setValues(allRows);

    // Formats
    sheet.getRange(2, 11, allRows.length, 1).setNumberFormat("yyyy-mm-dd"); // Erstelldatum
    sheet.getRange(2, 13, allRows.length, 2).setNumberFormat('#,##0.00 "€"'); // Bezahlt & OP-Saldo
    sheet.getRange(2, 16, allRows.length, 1).setNumberFormat("yyyy-mm-dd"); // Mahndatum
    sheet.getRange(2, 17, allRows.length, 1).setNumberFormat("yyyy-mm-dd"); // Fällig?
    sheet.getRange(2, 18, allRows.length, 1).setNumberFormat("0");          // Fällig seit Tage

    // Apply Filter & Border
    const fullRange = sheet.getRange(1, 1, allRows.length + 1, HEADERS.length);
    fullRange.setBorder(true, true, true, true, true, true, CONFIG.FORMAT_BORDER_COLOR || "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    if (!sheet.getFilter()) {
      fullRange.createFilter();
    }

    sheet.autoResizeColumns(1, HEADERS.length);


  } 
}

/**
 * Helper function to fetch only customers belonging to target representative
 * @param {string} targetVtrNr 
 * @returns {Map<string, Object>} Map of AdrNr -> Customer details
 */
function getRepresentativeCustomers(targetVtrNr) {
  const customerMap = new Map();
  const queryAddresses = `
    query GetRepresentativeAddresses($vtrNr: String!, $cursor: String) {
      tblAdresse {
        conRead(
          first: 100, 
          after: $cursor,
          fastFilter: {
            eq: [{ field: fldVtrNr }, { value: { string: $vtrNr } }]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldKredLimit
              fldVtrNr
              fldStatus
              rowsAnschriften {
                fldNa2
                fldNa3
                fldLandBez
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
  const excludedAdresses = CONFIG.AUSGESCHLOSSENE_ADRESSEN || [];

  while (hasNextPage) {
    const payload = {
      query: queryAddresses,
      variables: { vtrNr: targetVtrNr, cursor: cursor }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) break;

      const conRead = json.data.tblAdresse?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const adrNr = String(node.fldAdrNr || "").trim();
        const status = String(node.fldStatus || "").trim();

        // Skip invalid, suppliers, or excluded customers
        if (!adrNr || status === "Lieferant" || excludedAdresses.includes(adrNr)) {
          return;
        }

        const anschriften = node.rowsAnschriften || [];
        const primary = anschriften.length > 0 ? anschriften[0] : {};

        customerMap.set(adrNr, {
          vtrNr: String(node.fldVtrNr || "").trim(),
          kredLimit: node.fldKredLimit || 0,
          status: status,
          name2: String(primary.fldNa2 || ""),
          name3: String(primary.fldNa3 || ""),
          land: String(primary.fldLandBez || "")
        });
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler bei getRepresentativeCustomers: " + e.message);
      break;
    }
  }

  return customerMap;
}