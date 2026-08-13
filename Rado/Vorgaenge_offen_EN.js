/**
 * fetchVorgaenge.gs
 * Fetches open orders / processes (Vorgänge) filtered by representative (fldVtrNr)
 * and active order types (Spalte C condition in master catalog) using GraphQL fastFilter.
 */
function fetchVorgaenge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Vertreter & Sprache bestimmen
  const vtrNr = String(CONFIG.VERTRETER_NR || "56").trim();
  const englishReps = ["56", "60"]; // Robin Carter Browne, Rado Kabakov
  const isEnglish = englishReps.includes(vtrNr);

  const zielBlattName = CONFIG.SHEET_OPEN_ORDERS || (isEnglish ? "Open Orders" : "Vorgaenge");

  // Spalten-Header definieren (Englisch oder Deutsch)
  const HEADERS = isEnglish ? [
    "Doc_No", "Order_No", "Doc_Type", "Doc_Type_Name", "Rep_No", "Cust_No", "Date", "Delivery_Date", "Block_Flag",
    "Inv_Name2", "Inv_Name3", "Inv_Country", 
    "Del_Name2", "Del_Name3", "Del_Street", "Del_ZIP", "Del_City", "Del_Country",
    "Item_No", "Item_Name1", "Item_Name3", "Quantity", "Unit_Price_Net", "Total_Price_Net", "Warehouse_No", "Billed_Flag"
  ] : [
    "fldBelegNr", "fldAuftrNr", "fldArt", "Vorgangsart_Klartext", "fldVtrNr", "fldAdrNr", "fldDat", "fldLiefDat", "fldGspKz",
    "fldReNa2", "fldReNa3", "fldReLandBez", 
    "fldLiNa2", "fldLiNa3", "fldLiStr", "fldLiPLZ", "fldLiOrt", "fldLiLandBez",
    "fldArtNr", "fldKuBez1", "fldKuBez3", "fldMge", "fldEPrNt", "Position Gesamtpreis Netto", "fldAusLagNr", "fldAbrPosKz"
  ];

  // 2. Vorgangsarten aus externer Master-Datei einlesen (Spalte C Bedingung)
  const vorgangsartenMap = new Map();
  try {
    const masterSs = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    const katalogSheet = masterSs.getSheetByName(CONFIG.TAB_VORGANGSARTEN);

    if (katalogSheet) {
      const lastRow = katalogSheet.getLastRow();
      if (lastRow >= 2) {
        // Spalte A = Code, Spalte B = Klartext, Spalte C = Filter/Aktiv-Kennzeichnung
        const katalogDaten = katalogSheet.getRange(2, 1, lastRow - 1, 3).getValues();

        katalogDaten.forEach(row => {
          const code = String(row[0] || "").trim();
          const klartext = String(row[1] || "").trim();
          const colC = String(row[2] || "").trim();

          // REGEL: Nur aufnehmen, wenn Code vorhanden UND Spalte C NICHT LEER ist!
          if (code !== "" && colC !== "") {
            vorgangsartenMap.set(code, klartext);
          }
        });
      }
    } else {
      Logger.log(`WARNUNG: Blatt '${CONFIG.TAB_VORGANGSARTEN}' in Master-Datei nicht gefunden.`);
    }
  } catch (e) {
    Logger.log("Fehler beim Zugriff auf die Master-Vorgangsarten: " + e.message);
  }

  const vorgangsArtenList = Array.from(vorgangsartenMap.keys());
  if (vorgangsArtenList.length === 0) {
    ss.toast("Keine aktiven Vorgangsarten mit gefüllter Spalte C gefunden.", "Hinweis");
    return;
  }

  const formattedVorgangsArten = vorgangsArtenList.map(art => ({ string: art }));

  // 3. Zielblatt vorbereiten & Altdaten löschen (Inhalte ab Zeile 2 leeren)
  let zielSheet = ss.getSheetByName(zielBlattName);
  let isNewSheet = false;

  if (!zielSheet) {
    zielSheet = ss.insertSheet(zielBlattName);
    isNewSheet = true;
  }

  const currentFilter = zielSheet.getFilter();
  if (currentFilter) currentFilter.remove();

  const maxRows = zielSheet.getMaxRows();
  const maxCols = zielSheet.getMaxColumns();

  if (maxRows > 1) {
    zielSheet.getRange(2, 1, maxRows - 1, Math.max(maxCols, HEADERS.length)).clearContent();
  }

  if (isNewSheet) {
    zielSheet.getRange(1, 1, 1, HEADERS.length)
             .setValues([HEADERS])
             .setBackground(CONFIG.FORMAT_HEADER_BG || "#3bb7c4")
             .setFontColor(CONFIG.FORMAT_HEADER_TEXT || "#000000")
             .setFontWeight("bold");
  }

  // 4. GraphQL Query mit serverseitigem fastFilter (Vorgangsarten UND Vertreternummer)
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
                fldEPrNt
                fldMge
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

  let hasNextPage = true;
  let cursor = null;
  const zeilen = [];

  const startTime = Date.now();
  const MAX_TIME_MS = 5 * 60 * 1000; // 5 Min Schutzbremse

  

  // 5. Pagination Loop
  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit erreicht. Bisher geladene Vorgänge werden gespeichert.");
      ss.toast("Zeitlimit erreicht! Speichere geladene Daten...", "Achtung", 5);
      break;
    }

    const payload = {
      query: query,
      variables: {
        vorgangsArten: formattedVorgangsArten,
        vtrNr: vtrNr,
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
        Logger.log("GraphQL Error in fetchVorgaenge: " + JSON.stringify(json.errors || response.getContentText()));
        ss.toast("Fehler bei der API-Abfrage.", "Fehler");
        return;
      }

      const conRead = json.data.tblVorgang?.conRead || {};
      const edges = conRead.edges || [];

      // Entpacken der Knoten direkt in die zeilen-Matrix
      edges.forEach(edge => {
        const vorgang = edge.node;
        if (!vorgang) return;

        const artCode = String(vorgang.fldArt || "").trim();
        const artKlartext = vorgangsartenMap.get(artCode) || artCode;
        const belegNr = String(vorgang.fldBelegNr || "").trim();

        const kopfDaten = [
          belegNr,
          String(vorgang.fldAuftrNr || ""),
          artCode,
          artKlartext,
          String(vorgang.fldVtrNr || ""),
          String(vorgang.fldAdrNr || ""),
          vorgang.fldDat ? new Date(vorgang.fldDat) : "",
          vorgang.fldLiefDat ? new Date(vorgang.fldLiefDat) : "",
          vorgang.fldGspKz || false,
          String(vorgang.fldReNa2 || ""),
          String(vorgang.fldReNa3 || ""),
          String(vorgang.fldReLandBez || ""),
          String(vorgang.fldLiNa2 || ""),
          String(vorgang.fldLiNa3 || ""),
          String(vorgang.fldLiStr || ""),
          String(vorgang.fldLiPLZ || ""),
          String(vorgang.fldLiOrt || ""),
          String(vorgang.fldLiLandBez || "")
        ];

        const positionen = vorgang.rowsPositions || [];

        if (positionen.length > 0) {
          positionen.forEach(pos => {
            let menge = pos.fldMge || 0;
            let einzelpreisNetto = pos.fldEPrNt || 0;

            // KAUFMÄNNISCHE VORZEICHENLOGIK (z. B. Korrekturen/Gutschriften)
            if (artCode === "123" || belegNr.startsWith("123")) {
              menge = -Math.abs(menge);
            }

            const gesamtpreisNetto = menge * einzelpreisNetto;

            zeilen.push([
              ...kopfDaten,
              String(pos.fldArtNr || ""),
              String(pos.rowArtikel?.fldKuBez1 || ""),
              String(pos.rowArtikel?.fldKuBez3 || ""),
              menge,
              einzelpreisNetto,
              gesamtpreisNetto,
              String(pos.fldAusLagNr || ""),
              pos.fldAbrPosKz ? "Ja" : "Nein"
            ]);
          });
        } else {
          zeilen.push([
            ...kopfDaten,
            "", "", "", 0, 0, 0, "", "Nein"
          ]);
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler im Loop von fetchVorgaenge: " + e.message);
      ss.toast("Skriptfehler: " + e.message, "Fehler");
      break;
    }
  }

  // 6. Schreiben & Formatieren
  if (zeilen.length > 0) {
    // Identifikatoren als Text formatieren (Spalten A:F, J:R, S:U, Y:Z)
    zielSheet.getRange("A:F").setNumberFormat("@");
    zielSheet.getRange("J:R").setNumberFormat("@");
    zielSheet.getRange("S:U").setNumberFormat("@");
    zielSheet.getRange("Y:Z").setNumberFormat("@");

    // Werte einfügen ab Zeile 2
    zielSheet.getRange(2, 1, zeilen.length, HEADERS.length).setValues(zeilen);

    // Datumsspalten (G & H)
    zielSheet.getRange(2, 7, zeilen.length, 2).setNumberFormat("yyyy-mm-dd");
    // Mengen (Spalte V / 22)
    zielSheet.getRange(2, 22, zeilen.length, 1).setNumberFormat("#,##0");
    // Preise Netto & Gesamt Netto (Spalten W & X / 23 & 24)
    zielSheet.getRange(2, 23, zeilen.length, 2).setNumberFormat('#,##0.00 "€"');

    // Filter und Rahmen über die Gesamttabelle legen
    const fullRange = zielSheet.getRange(1, 1, zeilen.length + 1, HEADERS.length);
    fullRange.setBorder(true, true, true, true, true, true, CONFIG.FORMAT_BORDER_COLOR || "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    if (!zielSheet.getFilter()) {
      fullRange.createFilter();
    }

    zielSheet.autoResizeColumns(1, HEADERS.length);

   
  } else {
    ss.toast(`Keine offenen Vorgänge für Vertreter ${vtrNr} gefunden.`, "Hinweis", 5);
  }
}