function fetchVannierData() {
  // --- CONFIGURATION ---
  const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
  const API_TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
  const MASTER_SHEET_ID = "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ";
  const TAB_VORGANGSARTEN = "Vorgangsarten";
  
  const SHEET_NAME = "Vannier";
  const TARGET_KUNDE = "3114158";
  const START_DATUM = "2024-01-01T00:00:00Z";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
    sheet.clearFormats();
  }

  // --- 1. LOAD ARCHIVE DOCUMENT TYPES FROM MASTER SHEET ---
  const vorgangsartenMap = new Map();
  const archivArten = [];

  try {
    const externalSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const katalogSheet = externalSs.getSheetByName(TAB_VORGANGSARTEN);
    if (katalogSheet) {
      const katalogData = katalogSheet.getDataRange().getValues();
      // Skip header row (row 0)
      for (let i = 1; i < katalogData.length; i++) {
        const code = String(katalogData[i][0] || "").trim();      // Col A: Code
        const label = String(katalogData[i][1] || "").trim();     // Col B: Bezeichnung
        const category = String(katalogData[i][2] || "").trim().toLowerCase(); // Col C: Vorgangsarten

        // Strictly ignore empty Column C
        if (code !== "" && category !== "") {
          if (category === "archiv") {
            vorgangsartenMap.set(code, label);
            archivArten.push(code); 
          }
        }
      }
    }
  } catch (e) {
    Logger.log("Error loading document types from Master Sheet: " + e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error accessing Master Sheet: " + e.message, "Error", 10);
    return;
  }

  if (archivArten.length === 0) {
    sheet.getRange(1, 1).setValue("No archive document types found in Master Sheet.");
    return;
  }

  const formattedVorgangsArten = archivArten.map(art => ({ string: art }));

  // --- 2. GRAPHQL QUERY (With fldErstDat as Fallback for Date) ---
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
              fldErstDat
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

  // English Column Headers
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

  SpreadsheetApp.getActiveSpreadsheet().toast("Loading Vannier data...", "Please wait");

  // --- 3. PAGINATION LOOP ---
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

        // --- SAFE DATE PARSING ---
        // Uses fldDat if available, otherwise falls back to fldErstDat
        const rawDateStr = node.fldDat || node.fldErstDat;
        let documentDate = "";
        if (rawDateStr) {
          const parsedDate = new Date(rawDateStr);
          if (!isNaN(parsedDate.getTime())) {
            documentDate = parsedDate;
          }
        }

        (node.rowsPositions || []).forEach(pos => {
          // Only billed positions
          if (pos.fldAbrPosKz !== true) return;

          let mge = pos.fldMge || 0;
          let eprNt = pos.fldEPrNt || 0;
          let eekRoh = pos.fldEEkRoh || 0;

          // Description formatting
          let descText = Array.isArray(pos.fldBez) ? pos.fldBez.join(" ") : (pos.fldBez || "");
          if (descText.length > 120) {
            descText = descText.substring(0, 120);
          }

          const descDE = pos.rowArtikel?.fldKuBez1 || "";

          // --- COMMERCIAL SIGN LOGIC FOR CORRECTIONS (DocType 123) ---
          const isCorrection = (artCode === "123" || belegNr.startsWith("123"));

          // Force quantity to be negative for credit notes/corrections
          if (isCorrection) {
            mge = -Math.abs(mge);
          }

          // Commercial calculations
          const gesamtNetto = mge * eprNt;
          const ekRohGesamt = mge * eekRoh;
          const roherloes = gesamtNetto - ekRohGesamt;

          allRows.push([
            String(node.fldAdrNr || ""),
            artCode,
            vorgangsartenMap.get(artCode) || "",
            String(node.fldAuftrNr || ""),
            belegNr,
            documentDate, // Safe JavaScript Date Object
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
      Logger.log("Error in fetchVannierData: " + e.message);
      break;
    }
  }

  // --- 4. WRITE DATA & FORMATTING ---
  if (allRows.length === 1) {
    sheet.getRange(1, 1).setValue("No archive documents found for customer " + TARGET_KUNDE + " since " + START_DATUM);
    return;
  }

  const totalRows = allRows.length;
  const totalCols = headers.length;

  // Set ID columns explicitly as Text
  sheet.getRange("A:E").setNumberFormat("@"); 
  sheet.getRange("L:L").setNumberFormat("@");

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(allRows);

  // Apply Cell Number Formats
  if (totalRows > 1) {
    sheet.getRange(2, 6, totalRows - 1, 1).setNumberFormat("yyyy-mm-dd");      // fldDat (Spalte F)
    sheet.getRange(2, 15, totalRows - 1, 1).setNumberFormat("#,##0.00");       // fldMge (Spalte O)
    sheet.getRange(2, 16, totalRows - 1, 2).setNumberFormat("#,##0.00 €");     // fldEPrNt & fldGesamtNetto (Spalten P & Q)
    sheet.getRange(2, 18, totalRows - 1, 3).setNumberFormat("#,##0.00 €");     // fldEEkRoh, fldEkRohGesamt & fldRoherloes (Spalten R, S, T)
    sheet.getRange(2, 21, totalRows - 1, 1).setNumberFormat("#,##0.00");       // fldOMge (Spalte U)
  }

  // Header Styling
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  range.setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, totalCols);

  SpreadsheetApp.getActiveSpreadsheet().toast("Vannier archive data loaded successfully!", "Done", 5);
}