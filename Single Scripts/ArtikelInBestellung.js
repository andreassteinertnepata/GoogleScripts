// --- CENTRAL CONFIGURATION ---

// Es wird überprüft, in welcher Bestellung der offenen Vorgänge ein Artikel ist

const CONFIG = {
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  TARGET_SHEET_NAME: "HP7",
  TARGET_ARTICLE_NR: "109.007.151",
  // Open transaction types according to Vorgangsarten configuration
  OPEN_VORGANGS_ARTEN: [
    "30", "102", "106", "107", "108", "112", "114", "118", "120", 
    "127", "163", "164", "167", "185", "188", "189", "198", "200"
  ]
};

// GraphQL-Query: fldGspKz auf Auftragsebene (rowVorgang) ergänzt
const GRAPHQL_QUERY_POSITION_FILTER = `
query GetPositionsByArticle($targetArtNr: FilterValue!, $cursor: String) {
  tblVorgangPosition {
    conRead(
      first: 100,
      after: $cursor,
      fastFilter: {
        eq: [{ field: fldArtNr }, { value: $targetArtNr }]
      }
    ) {
      edges {
        node {
          fldArtNr
          fldMge
          fldEPrNt
          fldEEkRoh
          fldAusLagNr
          fldAbrPosKz
          rowArtikel {
            fldKuBez1
            fldKuBez3
          }
          rowVorgang {
            fldBelegNr
            fldAuftrNr
            fldArt
            lblArt
            fldVtrNr
            fldAdrNr
            fldDat
            fldLiefDat
            fldGspKz
            fldReNa2
            fldReNa3
            fldReLandBez
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

/**
 * Imports all open order positions containing Article 109.007.151 into target sheet
 */
function fetchTDSBaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  
  // 1. Target Sheet Setup
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TARGET_SHEET_NAME);
  } else {
    const activeFilter = sheet.getFilter();
    if (activeFilter) activeFilter.remove();
    sheet.clear();
  }

  ss.toast(`Fetching open orders for article ${CONFIG.TARGET_ARTICLE_NR}...`, "Datahub API", -1);

  let allRows = [];
  let hasNextPage = true;
  let cursor = null;

  const openTypesSet = new Set(CONFIG.OPEN_VORGANGS_ARTEN);

  // 2. API Data Fetching
  while (hasNextPage) {
    const payload = {
      query: GRAPHQL_QUERY_POSITION_FILTER,
      variables: {
        targetArtNr: { string: CONFIG.TARGET_ARTICLE_NR },
        cursor: cursor
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Error: " + JSON.stringify(json.errors || response.getContentText()));
        ss.toast("API Query Error. Check Logs.", "Error");
        return;
      }

      const conRead = json.data.tblVorgangPosition?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const pos = edge.node;
        if (!pos) return;

        const vorgang = pos.rowVorgang || {};
        const docType = String(vorgang.fldArt || "").trim();

        // Filtern: Nur offene Vorgangsarten verarbeiten
        if (!openTypesSet.has(docType)) return;

        const qty = pos.fldMge || 0;
        const unitPriceNet = pos.fldEPrNt || 0;
        const unitCostEk = pos.fldEEkRoh || 0;
        const totalPriceNet = qty * unitPriceNet;

        // Sperrstatus auf Auftragsebene
        const orderBlocked = vorgang.fldGspKz === true ? "Yes" : "No";

        allRows.push([
          String(vorgang.fldBelegNr || ""),
          String(vorgang.fldAuftrNr || ""),
          docType,
          String(vorgang.lblArt || docType),
          String(vorgang.fldVtrNr || ""),
          String(vorgang.fldAdrNr || ""),
          vorgang.fldDat ? new Date(vorgang.fldDat) : "",
          vorgang.fldLiefDat ? new Date(vorgang.fldLiefDat) : "",
          orderBlocked, // <--- Sperrkennzeichen auf Auftragsebene
          String(vorgang.fldReNa2 || ""),
          String(vorgang.fldReNa3 || ""),
          String(vorgang.fldReLandBez || ""),
          String(pos.fldArtNr || ""),
          String(pos.rowArtikel?.fldKuBez1 || ""),
          String(pos.rowArtikel?.fldKuBez3 || ""),
          qty,
          unitPriceNet,
          totalPriceNet,
          unitCostEk,
          String(pos.fldAusLagNr || ""),
          pos.fldAbrPosKz === true ? "Yes" : "No"
        ]);
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Execution Error: " + e.toString());
      ss.toast("Error: " + e.message, "Error");
      return;
    }
  }

  // 3. Headers & Data Writing
  const headers = [
    "DocNo", "OrderNo", "DocType", "DocType_Label", "RepNo", "CustNo", "DocDate", "DeliveryDate",
    "Order_Blocked", "Customer_Name2", "Customer_Name3", "Country",
    "ArticleNo", "Article_Desc1", "Article_Desc3", "Qty", "UnitPrice_Net", "TotalPrice_Net", "UnitCost_EK", "Warehouse", "Billed"
  ];

  if (allRows.length === 0) {
    sheet.getRange(1, 1).setValue(`No open orders found for article ${CONFIG.TARGET_ARTICLE_NR}.`);
    ss.toast("Finished - No open matching orders found.", "Complete", 5);
    return;
  }

  const outputData = [headers, ...allRows];
  const totalRows = outputData.length;
  const totalCols = headers.length;

  // 4. Format Cell Types
  sheet.getRange("A:F").setNumberFormat("@"); // Identifiers as Text
  sheet.getRange("I:L").setNumberFormat("@"); // Order Blocked & Customer Details
  sheet.getRange("M:O").setNumberFormat("@"); // Article details
  sheet.getRange("T:U").setNumberFormat("@"); // Warehouse / Billed Status

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(outputData);

  // 5. Apply Number & Date Formatting
  if (totalRows > 1) {
    sheet.getRange(2, 7, totalRows - 1, 2).setNumberFormat("yyyy-mm-dd");     // Dates (G & H)
    sheet.getRange(2, 16, totalRows - 1, 1).setNumberFormat("#,##0");          // Qty (P)
    sheet.getRange(2, 17, totalRows - 1, 3).setNumberFormat("#,##0.00 €");     // Net prices & costs (Q, R, S)
  }

  // 6. Styling & Layout
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  range.createFilter();
  sheet.autoResizeColumns(1, totalCols);

  ss.toast(`Successfully imported ${allRows.length} open positions for article ${CONFIG.TARGET_ARTICLE_NR}!`, "Success", 5);
}