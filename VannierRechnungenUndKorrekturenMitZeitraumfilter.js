function fetchVannierData() {
  // --- KONFIGURATION (DIREKT IN DER FUNKTION) ---
  const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
  const API_TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
  const MASTER_SHEET_ID = "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ";
  const TAB_VORGANGSARTEN = "Vorgangsarten";
  
  const BLATT_NAME = "Vannier";
  const TARGET_KUNDE = "3114158";
  const START_DATUM = "2024-01-01T00:00:00Z";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);
  if (!sheet) sheet = ss.insertSheet(BLATT_NAME);

  sheet.clear();
  sheet.clearFormats();

  // --- 1. ARCHIV-VORGANGSARTEN EXTERN AUSLESEN ---
  const vorgangsartenMap = new Map();
  const archivArten = [];

  try {
    const externalSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const katalogSheet = externalSs.getSheetByName(TAB_VORGANGSARTEN);
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
    Logger.log("Fehler beim Laden der Vorgangsarten aus Master Sheet: " + e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fehler beim Zugriff auf Master Sheet: " + e.message, "Abbruch", 10);
    return;
  }

  if (archivArten.length === 0) {
    sheet.getRange(1, 1).setValue("Keine Archiv-Vorgangsarten im Master-Sheet gefunden.");
    return;
  }

  const formattedVorgangsArten = archivArten.map(art => ({ string: art }));

  // --- 2. GRAPHQL QUERY ---
  const query = `
    query GetVannierArchiv($vorgangsArten: [FilterValue!]!, $startDatum: DateTime!, $kundeNr: String!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { eq: [{ field: fldAdrNr }, { value: { string: $kundeNr } }] },
              { ge: [{ field: fldDat }, { value: { datetime: $startDatum } }] },
              { in: { field: fldArt, values: $vorgangsArten } }
            ]
          },
          orderBy: [{ field: fldDat, desc: true }]
        ) {
          edges {
            node {
              fldAdrNr
              fldArt
              fldAuftrNr
              fldBelegNr
              fldDat
              fldReNa2
              fldReNa3
              fldReLandBez
              fldLiLandBez
              fldZahlBed
              fldMemo
              fldStorniertKz
              rowsPositions {
                fldArtNr
                fldBez
                fldMge
                fldEPrNt
                fldOMge
                fldEEkRoh
                fldAbrPosKz
                rowArtikel {
                  fldKuBez1
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

  const headers = [
    "fldAdrNr", 
    "fldArt",
    "lblArt", 
    "fldAuftrNr", 
    "fldBelegNr", 
    "fldDat", 
    "fldReNa2", 
    "fldReNa3", 
    "fldReLandBez", 
    "fldLiLandBez", 
    "fldZahlBed", 
    "fldArtNr", 
    "fldBez", 
    "fldKuBez1", 
    "fldMge", 
    "fldEPrNt", 
    "fldGesamtNetto",
    "fldEEkRoh",
    "fldEkRohGesamt",
    "fldRoherloes",
    "fldOMge",
    "fldMemo"
  ];

  let allRows = [headers];
  let hasNextPage = true;
  let cursor = null;

  SpreadsheetApp.getActiveSpreadsheet().toast("Lade Vannier-Daten...", "Bitte warten");

  // --- 3. PAGINATION SCHLEIFE ---
  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: {
          vorgangsArten: formattedVorgangsArten,
          startDatum: START_DATUM,
          kundeNr: TARGET_KUNDE,
          cursor: cursor
        }
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode !== 200) {
        throw new Error(`HTTP ${responseCode}: ${responseText}`);
      }

      const json = JSON.parse(responseText);
      if (json.errors) {
        throw new Error(`GraphQL Error: ${JSON.stringify(json.errors)}`);
      }

      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const artCode = String(node.fldArt || "").trim();
        const belegNr = String(node.fldBelegNr || "").trim();

        (node.rowsPositions || []).forEach(pos => {
          if (pos.fldAbrPosKz !== true) return;

          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;
          
          // WICHTIG: Einzel-EK für Roherlös direkt von der Belegposition nehmen
          let eekRoh = pos.fldEEkRoh || 0;

          // Positions-Bezeichnung auf max. 120 Zeichen beschränken
          let descText = Array.isArray(pos.fldBez) ? pos.fldBez.join(" ") : (pos.fldBez || "");
          if (descText.length > 120) {
            descText = descText.substring(0, 120);
          }

          const descDE = pos.rowArtikel?.fldKuBez1 || "";

          // RECHNUNGSKORREKTUR / GUTSCHRIFT (Vorgangsart 123 oder BelegNr beginnt mit "123")
          const istGutschrift = (artCode === "123" || belegNr.startsWith("123"));

          // Bei Gutschriften die Menge konsequent negativ erzwingen
          if (istGutschrift) {
            mge = -Math.abs(mge);
          }

          // Kaufmännische Gesamtwert- und Roherlösberechnung
          const gesamtNetto = mge * eprNt;
          const ekRohGesamt = mge * eekRoh;
          const roherloes = gesamtNetto - ekRohGesamt;

          allRows.push([
            String(node.fldAdrNr || ""),
            artCode,
            vorgangsartenMap.get(artCode) || "",
            String(node.fldAuftrNr || ""),
            belegNr,
            node.fldDat ? new Date(node.fldDat) : "",
            String(node.fldReNa2 || ""),
            String(node.fldReNa3 || ""),
            String(node.fldReLandBez || ""),
            String(node.fldLiLandBez || ""),
            String(node.fldZahlBed || ""),
            String(pos.fldArtNr || ""),
            descText,
            descDE,
            mge,
            eprNt,
            gesamtNetto,
            eekRoh,
            ekRohGesamt,
            roherloes,
            pos.fldOMge || 0,
            String(node.fldMemo || "")
          ]);
        });
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler bei fetchVannierData: " + e.message);
      break;
    }
  }

  // --- 4. ERGEBNISSE SCHREIBEN UND FORMATIEREN ---
  if (allRows.length === 1) {
    sheet.getRange(1, 1).setValue("Keine Archiv-Vorgänge für Kunde 3114158 seit 01.01.2024 gefunden.");
    return;
  }

  const totalRows = allRows.length;
  const totalCols = headers.length;

  sheet.getRange("A:E").setNumberFormat("@"); 
  sheet.getRange("L:L").setNumberFormat("@");

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(allRows);

  if (totalRows > 1) {
    sheet.getRange(2, 6, totalRows - 1, 1).setNumberFormat("yyyy-mm-dd");     // fldDat
    sheet.getRange(2, 15, totalRows - 1, 1).setNumberFormat("#,##0.00");      // fldMge
    sheet.getRange(2, 16, totalRows - 1, 2).setNumberFormat("#,##0.00 €");    // fldEPrNt & fldGesamtNetto
    sheet.getRange(2, 18, totalRows - 1, 3).setNumberFormat("#,##0.00 €");    // fldEEkRoh, fldEkRohGesamt & fldRoherloes
    sheet.getRange(2, 21, totalRows - 1, 1).setNumberFormat("#,##0.00");      // fldOMge
  }

  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  range.setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, totalCols);

  SpreadsheetApp.getActiveSpreadsheet().toast("Vannier-Archivdaten erfolgreich geladen!", "Fertig", 5);
}