// ==========================================
// CENTRAL CONFIGURATION
// ==========================================
const CONFIG = {
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  MASTER_SHEET_ID: "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ", // Oder aktive Spreadsheet-ID
  TAB_VORGANGSARTEN: "Vorgangsarten",
  TARGET_TAB: "Vorgaenge",
  VERTRETER_NR: "" // Optional: Z.B. "56" eintragen für Robin, oder leer lassen "" für alle Vertreter
};

function fetchVorgaenge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. DYNAMISCHES EINLESEN DER VORGANGSARTEN (Filter: Nur wenn Spalte C "Vorgänge" enthält) ---
  const vorgangsartenMap = new Map();
  const vorgaengeArtenList = [];

  let katalogSheet = ss.getSheetByName(CONFIG.TAB_VORGANGSARTEN);
  if (!katalogSheet) {
    try {
      const externalSs = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
      katalogSheet = externalSs.getSheetByName(CONFIG.TAB_VORGANGSARTEN);
    } catch (e) {
      Logger.log("Hinweis: Master-Sheet nicht erreichbar. Versuche lokales Blatt zu lesen.");
    }
  }

  if (katalogSheet) {
    const lastRow = katalogSheet.getLastRow();
    if (lastRow >= 2) {
      // Spalte A (Code), Spalte B (Klartext), Spalte C (Typ: Vorgänge/Archiv)
      const katalogDaten = katalogSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      katalogDaten.forEach(row => {
        const code = String(row[0] || "").trim();
        const klartext = String(row[1] || "").trim();
        const typColC = String(row[2] || "").trim().toLowerCase();

        // REGEL: Nur wenn Spalte C befüllt ist UND "vorgang" enthält (oder "x" / "ja")
        if (code !== "" && typColC !== "" && (typColC.includes("vorgang") || typColC === "x" || typColC === "ja")) {
          vorgangsartenMap.set(code, klartext);
          vorgaengeArtenList.push(code);
        }
      });
    }
  }

  if (vorgaengeArtenList.length === 0) {
    ss.toast("Keine gültigen Vorgangsarten in Spalte C ('Vorgänge') gefunden.", "Abbruch");
    return;
  }

  const formattedVorgangsArten = vorgaengeArtenList.map(art => ({ string: art }));

  // --- 2. ZIELBLATT VORBEREITEN ---
  let zielSheet = ss.getSheetByName(CONFIG.TARGET_TAB);
  if (!zielSheet) {
    zielSheet = ss.insertSheet(CONFIG.TARGET_TAB);
  } else {
    const currentFilter = zielSheet.getFilter();
    if (currentFilter) currentFilter.remove();
    zielSheet.clear();
  }

  ss.toast("Lade offene Vorgänge & Positionen von Datahub...", "Bitte warten", -1);

  // --- 3. GRAPHQL QUERY ---
  const query = `
    query GetVorgaenge($vorgangsArten: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            in: {
              field: fldArt,
              values: $vorgangsArten
            }
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldArt
              fldAuftrNr
              fldBelegNr
              fldDat
              fldGspKz
              fldLiLandBez
              fldLiNa2
              fldLiNa3
              fldLiOrt
              fldLiPLZ
              fldLiStr
              fldLiefDat
              fldReLandBez
              fldReNa2
              fldReNa3
              fldVtrNr
              rowsPositions {
                fldAbrPosKz
                fldArtNr
                fldAusLagNr
                fldEEkRoh
                fldEPrNt
                fldMge
                fldOMge
                rowArtikel {
                  fldKuBez1
                  fldKuBez3
                }
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

  let allRows = [];
  let hasNextPage = true;
  let cursor = null;

  // --- 4. PAGINATION SCHLEIFE ---
  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { 
          vorgangsArten: formattedVorgangsArten,
          cursor: cursor 
        }
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Fehler: " + JSON.stringify(json.errors || response.getContentText()));
        ss.toast("Fehler bei der API-Abfrage.", "Fehler");
        return;
      }

      const conRead = json.data.tblVorgang?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node;
        if (!node) return;

        // Falls Vertreter-Filter im CONFIG definiert ist (z.B. für Robin "56")
        if (CONFIG.VERTRETER_NR && String(node.fldVtrNr || "").trim() !== CONFIG.VERTRETER_NR) {
          return;
        }

        const artCode = String(node.fldArt || "").trim();
        const artKlartext = vorgangsartenMap.get(artCode) || "";

        const baseData = [
          String(node.fldBelegNr || ""),
          String(node.fldAuftrNr || ""),
          artCode,
          artKlartext,
          String(node.fldVtrNr || ""),
          String(node.fldAdrNr || ""),
          node.fldDat ? new Date(node.fldDat) : "",
          node.fldLiefDat ? new Date(node.fldLiefDat) : "",
          node.fldGspKz || false,
          String(node.fldReNa2 || ""),
          String(node.fldReNa3 || ""),
          String(node.fldReLandBez || ""),
          String(node.fldLiNa2 || ""),
          String(node.fldLiNa3 || ""),
          String(node.fldLiStr || ""),
          String(node.fldLiPLZ || ""),
          String(node.fldLiOrt || ""),
          String(node.fldLiLandBez || "")
        ];

        const positionen = node.rowsPositions || [];

        if (positionen.length > 0) {
          positionen.forEach(pos => {
            const mge = pos.fldMge || 0;
            const eprNt = pos.fldEPrNt || 0;
            const eekRoh = pos.fldEEkRoh || 0;

            allRows.push([
              ...baseData,
              String(pos.fldArtNr || ""),
              String(pos.rowArtikel?.fldKuBez1 || ""),
              String(pos.rowArtikel?.fldKuBez3 || ""),
              mge,
              eprNt,
              mge * eprNt,       // Gesamt-Netto
              eekRoh,
              mge * eekRoh,     // Gesamt-EK
              pos.fldOMge || 0, // Offene Menge
              String(pos.fldAusLagNr || ""),
              String(pos.fldAbrPosKz || "")
            ]);
          });
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Skriptfehler: " + e.message);
      ss.toast("Netzwerk- oder Skriptfehler: " + e.message, "Fehler");
      return;
    }
  }

  // --- 5. HEADERS & DATEN SCHREIBEN ---
  const headers = [
    "DocNo", "OrderNo", "TypeNo", "TypeName", "RepNo", "CustomerNo", "DocDate", "DeliveryDate", "IsBlocked",
    "BillName2", "BillName3", "BillCountry",
    "ShipName2", "ShipName3", "ShipStreet", "ShipZip", "ShipCity", "ShipCountry",
    "ItemNo", "ItemDesc1", "ItemDesc3", "Qty", "UnitPrice", "TotalPrice", "UnitCost", "TotalCost", "OpenQty", "WhsNo", "IsBilled"
  ];

  if (allRows.length === 0) {
    zielSheet.getRange(1, 1).setValue("Keine offenen Vorgänge für die gewählten Vorgangsarten gefunden.");
    ss.toast("Keine Daten gefunden.", "Fertig");
    return;
  }

  const rowsData = [headers, ...allRows];
  const totalRows = rowsData.length;
  const totalCols = headers.length;

  const range = zielSheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(rowsData);

  // Formatting (Nummern, Daten & Beträge)
  zielSheet.getRange("A:F").setNumberFormat("@");
  zielSheet.getRange("J:R").setNumberFormat("@");
  zielSheet.getRange("S:U").setNumberFormat("@");

  if (totalRows > 1) {
    zielSheet.getRange(2, 7, totalRows - 1, 2).setNumberFormat("yyyy-mm-dd");
    zielSheet.getRange(2, 22, totalRows - 1, 1).setNumberFormat("#,##0.00");
    zielSheet.getRange(2, 23, totalRows - 1, 4).setNumberFormat("#,##0.00 €");
    zielSheet.getRange(2, 27, totalRows - 1, 1).setNumberFormat("#,##0.00");
  }

  const headerRange = zielSheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  range.setBorder(true, true, true, true, true, true);
  if (!zielSheet.getFilter()) range.createFilter();
  zielSheet.autoResizeColumns(1, totalCols);

  ss.toast(`Erfolgreich ${totalRows - 1} Positionszeilen geladen!`, "Erfolg", 5);
}