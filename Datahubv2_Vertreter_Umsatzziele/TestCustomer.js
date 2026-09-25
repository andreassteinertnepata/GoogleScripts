/**
 * TestCustomer.gs - Diagnose-Funktion speziell für Kunde 3133582
 */

function parseBelegDatum(raw) {
  if (!raw) return null;
  if (raw instanceof Date) return raw;

  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d;

  if (typeof raw === "string" && raw.includes(".")) {
    const parts = raw.split(".");
    if (parts.length === 3) {
      const tag = parseInt(parts[0], 10);
      const monat = parseInt(parts[1], 10) - 1;
      const jahr = parseInt(parts[2], 10);
      return new Date(jahr, monat, tag);
    }
  }

  return null;
}

function testCustomer3133582() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheetName = "Test_Kunde_3133582";
  let sheet = ss.getSheetByName(testSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(testSheetName);
  } else {
    sheet.clear();
  }

  const targetKundeNr = "3133582";
  const now = new Date();
  const startOfYear = new Date(now.getFullYear(), 0, 1, 0, 0, 0);
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const dateFrom12M = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  const dateFromIso = dateFrom12M.toISOString();

  const headers = [
    "Tabelle", "Beleg-Nr.", "Belegart", "Belegdatum", "Vertreter-ID", "Kunden-Nr.", "Firmenname (ReNa2/3)",
    "Artikel-Nr.", "Menge", "Einzelpreis (€)", "Zeilensumme Netto (€)", 
    "Abrechnungspos? (fldAbrPosKz)", "Katalog", "Korrekturbeleg?", "Storniert?", "Sel14?"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground("#1c4587").setFontColor("#ffffff").setFontWeight("bold");

  const vorgangsarten = getActiveVorgangsarten();
  const allowedTypes = vorgangsarten.archiv;
  const formattedTypes = allowedTypes.map(function(art) { return { string: art }; });

  const query = `
    query GetTestCustomer($cursor: String, $dateFrom: DateTime!, $vorgangsArten: [FilterValue!]!, $kundeNr: String!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { eq: [{ field: fldAdrNr }, { value: { string: $kundeNr } }] },
              { ge: [{ field: fldDat }, { value: { datetime: $dateFrom } }] },
              { in: { field: fldArt, values: $vorgangsArten } }
            ]
          }
        ) {
          pageInfo { hasNextPage endCursor }
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              fldVtrNr
              fldAdrNr
              fldReNa1
              fldReNa2
              fldReNa3
              fldStorniertKz
              fldSel14
              rowsPositions {
                fldMge
                fldEPrNt
                fldAbrPosKz
                fldArtNr
                rowArtikel {
                  fldKatalog
                  fldArtNr
                }
              }
            }
          }
        }
      }
    }
  `;

  let hasNextPage = true;
  let cursor = null;
  const rows = [];
  let sum12M = 0;
  let sumYTD = 0;
  let sumMTD = 0;

  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { cursor: cursor, dateFrom: dateFromIso, vorgangsArten: formattedTypes, kundeNr: targetKundeNr }
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.errors) {
      Logger.log("GraphQL Error: " + JSON.stringify(json.errors));
      break;
    }

    const conRead = json.data && json.data.tblVorgangArchiv ? json.data.tblVorgangArchiv.conRead : null;
    if (!conRead || !conRead.edges) break;

    conRead.edges.forEach(function(edge) {
      const node = edge.node;
      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      
      const belegDatum = parseBelegDatum(node.fldDat) || parseBelegDatum(node.fldErstDat);
      const belegDatStr = belegDatum ? belegDatum.toLocaleDateString("de-DE") : "";
      
      const vtrNr = String(node.fldVtrNr || "");
      const adrNr = String(node.fldAdrNr || "");
      const storniert = node.fldStorniertKz === true;
      const sel14 = node.fldSel14 === true;

      const na1 = String(node.fldReNa1 || "").trim();
      const na2 = String(node.fldReNa2 || "").trim();
      const na3 = String(node.fldReNa3 || "").trim();
      let companyName = [na2, na3].filter(Boolean).join(" ").trim() || na1;

      const isMTD = Boolean(belegDatum && belegDatum.getTime() >= startOfMonth.getTime());
      const isYTD = Boolean(belegDatum && belegDatum.getTime() >= startOfYear.getTime());
      const isKorrektur = CONFIG.CORRECTION_TYPES.includes(artCode) || belegNr.startsWith("123");

      (node.rowsPositions || []).forEach(function(pos) {
        if (pos.fldAbrPosKz === false) return;

        const artNr = String(pos.fldArtNr || (pos.rowArtikel ? pos.rowArtikel.fldArtNr : "") || "").trim();
        const katalog = pos.rowArtikel ? String(pos.rowArtikel.fldKatalog || "") : "";
        const abrPos = pos.fldAbrPosKz;

        const mge = pos.fldMge || 0;
        const epr = pos.fldEPrNt || 0;

        // RAW ZEILENSUMME
        let lineTotal = mge * epr;

        // GUTSCHRIFTEN-INVERTIERUNG: Wenn der Beleg eine Korrektur/Gutschrift ist, Betrag negieren!
        if (isKorrektur) {
          lineTotal = -lineTotal;
        }

        const isValidPos = !storniert && !sel14 && (abrPos !== false);

        if (isValidPos) {
          sum12M += lineTotal;
          if (isYTD) sumYTD += lineTotal;
          if (isMTD) sumMTD += lineTotal;
        }

        rows.push([
          "tblVorgangArchiv",
          belegNr,
          artCode,
          belegDatStr,
          vtrNr,
          adrNr,
          companyName,
          artNr,
          mge,
          epr,
          lineTotal,
          abrPos === false ? "FALSE (Ignoriert)" : "TRUE/NULL (Gültig)",
          katalog || "Kein Katalog",
          isKorrektur ? "JA (Korrektur negiert)" : "NEIN (Rechnung)",
          storniert ? "JA (Ausschluss)" : "NEIN",
          sel14 ? "JA (Ausschluss)" : "NEIN"
        ]);
      });
    });

    hasNextPage = conRead.pageInfo ? conRead.pageInfo.hasNextPage : false;
    cursor = conRead.pageInfo ? conRead.pageInfo.endCursor : null;
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 6).setNumberFormat("@");
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 9, rows.length, 1).setNumberFormat("#,##0.00");
    sheet.getRange(2, 10, rows.length, 2).setNumberFormat("#,##0.00 €");
    
    sheet.appendRow(["SUMMEN (Inklusive stornierender Gutschriften)", "", "", "", "", targetKundeNr, "MTD (Sep):", sumMTD, "YTD (2026):", sumYTD, "12M Gesamt:", sum12M, "", "", "", ""]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.getRange(lastRow, 8).setNumberFormat("#,##0.00 €");
    sheet.getRange(lastRow, 10).setNumberFormat("#,##0.00 €");
    sheet.getRange(lastRow, 12).setNumberFormat("#,##0.00 €");

    sheet.autoResizeColumns(1, headers.length);
  } else {
    sheet.getRange(2, 1).setValue("Keine Belege für Kunde " + targetKundeNr + " gefunden.");
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Test für Kunde ${targetKundeNr} beendet. MTD: ${sumMTD.toFixed(2)}€ | YTD: ${sumYTD.toFixed(2)}€ | 12M: ${sum12M.toFixed(2)}€`, "Kunden-Test", 10);
}