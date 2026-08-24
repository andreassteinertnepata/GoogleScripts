function verkaeufeAktionPruefenUndSenden() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBedingungen = ss.getSheetByName("Bedingungen");

  if (!sheetBedingungen) {
    Logger.log("Fehler: Tabellenblatt 'Bedingungen' wurde nicht gefunden!");
    return;
  }

  // ==========================================
  // 1. KOPFDATEN LESEN (B2: EMPFÄNGER, B3: BEGINN, B4: ENDE)
  // ==========================================
  const empfaenger = String(sheetBedingungen.getRange("B2").getValue() || "")
    .replace(/;/g, ",")
    .trim();
    
  const rawBeginn = sheetBedingungen.getRange("B3").getValue();
  const rawEnde   = sheetBedingungen.getRange("B4").getValue();

  if (!rawBeginn || !rawEnde) {
    Logger.log("Fehler: Beginn- (B3) oder Ende-Datum (B4) fehlt!");
    return;
  }

  const startDatumISO = new Date(rawBeginn).toISOString();
  const endeDatum = new Date(rawEnde);
  endeDatum.setHours(23, 59, 59, 999);
  const endeDatumISO = endeDatum.toISOString();

  // Mapping für fldAbwArtDatGrp
  const typeMapping = { "0": "End Customer", "1": "A", "2": "B", "3": "C" };

  // ==========================================
  // 2. ARTIKELLISTE AUSLESEN (ZEILE 8 BIS ENDE)
  // ==========================================
  const lastRow = sheetBedingungen.getLastRow();
  if (lastRow < 8) {
    Logger.log("Keine Artikel in Zeile 8ff. gefunden.");
    return;
  }

  const artikelRange = sheetBedingungen.getRange(8, 1, lastRow - 7, 4);
  const artikelDaten = artikelRange.getValues();

  const zielArtikelMap = new Map();
  const artikelNummernListe = [];

  for (let i = 0; i < artikelDaten.length; i++) {
    const artNr = String(artikelDaten[i][0] || "").trim();
    const artName = String(artikelDaten[i][1] || "").trim();
    const schwellwert = Number(artikelDaten[i][2] || 0);

    if (artNr !== "") {
      zielArtikelMap.set(artNr, {
        zeileIndex: i + 8,
        name: artName || artNr,
        schwellwert: schwellwert
      });
      artikelNummernListe.push(artNr);
    }
  }

  if (artikelNummernListe.length === 0) {
    Logger.log("Keine gültigen Artikelnummern in Spalte A gefunden.");
    return;
  }

  // ==========================================
  // 3. VORGANGSARTEN LADEN
  // ==========================================
  const archivArten = [];
  try {
    const externalSs = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    const katalogSheet = externalSs.getSheetByName(CONFIG.TAB_VORGANGSARTEN);
    if (katalogSheet) {
      const katalogData = katalogSheet.getDataRange().getValues();
      for (let i = 1; i < katalogData.length; i++) {
        const code = String(katalogData[i][0]).trim();
        const archivWert = String(katalogData[i][2]).trim().toLowerCase();
        if (code !== "" && (archivWert === "archiv" || archivWert === "x" || archivWert === "ja" || katalogData[i][2] === true)) {
          archivArten.push(code);
        }
      }
    }
  } catch (e) {
    Logger.log("Hinweis: Master-Sheet Vorgangsarten konnte nicht geladen werden.");
  }

  const finaleArten = archivArten.length > 0 ? archivArten : ["70", "105", "109", "110", "113", "115", "122", "123", "129", "154", "155"];
  const formattedVorgangsArten = finaleArten.map(art => ({ string: art }));

  // ==========================================
  // 4. GRAPHQL DATAHUB ABFRAGE (INKL. fldAbwArtDatGrp)
  // ==========================================
  const query = `
    query GetVerkaeufeAktion($vorgangsArten: [FilterValue!]!, $cursor: String, $startDatum: DateTime!, $endeDatum: DateTime!) {
      tblVorgangArchiv {
        conRead(
          first: 100, 
          after: $cursor, 
          fastFilter: { 
            and: [ 
              { ge: [{ field: fldErstDat }, { value: { datetime: $startDatum } }] },
              { le: [{ field: fldErstDat }, { value: { datetime: $endeDatum } }] },
              { in: { field: fldArt, values: $vorgangsArten } }
            ] 
          }
        ) {
          edges { 
            node { 
              fldAdrNr 
              rowAdresse {
                fldAbwArtDatGrp
              }
              fldArt 
              fldAuftrNr 
              fldBelegNr 
              fldDat 
              fldReNa2 
              fldReNa3 
              fldReLandBez 
              fldLiLandBez 
              fldZahlBed 
              rowsPositions { 
                fldArtNr 
                fldMge 
                fldEPrNt 
                fldAbrPosKz 
                rowArtikel { 
                  fldKuBez1 
                } 
              } 
            } 
          }
          pageInfo { hasNextPage endCursor }
        }
      }
    }
  `;

  const ergebnisseMap = new Map();
  artikelNummernListe.forEach(artNr => {
    ergebnisseMap.set(artNr, { summe: 0, zeilen: [] });
  });

  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = {
      query: query,
      variables: {
        vorgangsArten: formattedVorgangsArten,
        cursor: cursor,
        startDatum: startDatumISO,
        endeDatum: endeDatumISO
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
        Logger.log("GraphQL Error: " + JSON.stringify(json.errors || response.getContentText()));
        break;
      }

      const conRead = json.data?.tblVorgangArchiv?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node || {};
        const artCode = String(node.fldArt || "").trim();
        const belegNr = String(node.fldBelegNr || "").trim();
        
        // Auslesen & Mappen von fldAbwArtDatGrp aus rowAdresse
        const rawAbwGruppe = node.rowAdresse?.fldAbwArtDatGrp ? String(node.rowAdresse.fldAbwArtDatGrp).trim() : "";
        const customerType = typeMapping[rawAbwGruppe] || rawAbwGruppe;

        (node.rowsPositions || []).forEach(pos => {
          if (pos.fldAbrPosKz !== true) return;
          
          const posArtNr = String(pos.fldArtNr || "").trim();
          
          if (ergebnisseMap.has(posArtNr)) {
            let mge = pos.fldMge || 0;
            let eprNt = pos.fldEPrNt || 0;

            if (artCode === "123" || belegNr.startsWith("123")) {
              eprNt = Math.abs(eprNt);
              mge = -Math.abs(mge);
            }

            const artData = ergebnisseMap.get(posArtNr);
            artData.summe += mge;

            // Ohne Artikel-Bez 3
            artData.zeilen.push([
              String(node.fldAdrNr || ""), 
              customerType,
              artCode, 
              String(node.fldAuftrNr || ""), 
              belegNr, 
              node.fldDat ? new Date(node.fldDat) : "",
              String(node.fldReNa2 || ""), 
              String(node.fldReNa3 || ""), 
              String(node.fldReLandBez || ""), 
              String(node.fldLiLandBez || ""), 
              String(node.fldZahlBed || ""),
              posArtNr, 
              String(pos.rowArtikel?.fldKuBez1 || ""), 
              mge, 
              eprNt, 
              mge * eprNt
            ]);
          }
        });
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler bei Datahub-Abfrage: " + e.toString());
      break;
    }
  }

  // ==========================================
  // 5. BLÄTTER SCHREIBEN & SPALTE D AKTUALISIEREN
  // ==========================================
  // Bereinigte Header (ohne Artikel-Bez 3)
  const HEADERS = [
    "Adress-Nr.", "Customer Type", "Vorgangsart", "Auftrags-Nr.", "Beleg-Nr.", "Datum", 
    "Re-Name 2", "Re-Name 3", "Rechnungsland", "Lieferland", "Zahlungsart", 
    "Artikel-Nr.", "Artikel-Bez 1", "Menge", "EP Netto", "Gesamt Netto"
  ];

  const ueberschritteneArtikel = [];

  zielArtikelMap.forEach((info, artNr) => {
    const ergebnis = ergebnisseMap.get(artNr);
    const aktuelleVerkaeufe = ergebnis.summe;

    // 5a. Verkäufe in Spalte D des Blattes "Bedingungen" eintragen
    sheetBedingungen.getRange(info.zeileIndex, 4).setValue(aktuelleVerkaeufe);
    sheetBedingungen.getRange(info.zeileIndex, 4).setNumberFormat("#,##0");

    // 5b. Schwellenwert prüfen
    if (aktuelleVerkaeufe > info.schwellwert) {
      ueberschritteneArtikel.push({
        artNr: artNr,
        name: info.name,
        verkaeufe: aktuelleVerkaeufe,
        schwellwert: info.schwellwert
      });
    }

    // 5c. Artikelblatt befüllen
    const zielBlattName = info.name;
    if (zielBlattName !== "") {
      let sheetZiel = ss.getSheetByName(zielBlattName);
      if (!sheetZiel) {
        sheetZiel = ss.insertSheet(zielBlattName);
      }

      const currentFilter = sheetZiel.getFilter();
      if (currentFilter) currentFilter.remove();
      sheetZiel.clear();

      sheetZiel.getRange(1, 1, 1, HEADERS.length)
               .setValues([HEADERS])
               .setBackground(CONFIG.FORMAT_HEADER_BG || "#3bb7c4")
               .setFontColor(CONFIG.FORMAT_HEADER_TEXT || "#000000")
               .setFontWeight("bold");

      if (ergebnis.zeilen.length > 0) {
        sheetZiel.getRange("A:E").setNumberFormat("@"); 
        sheetZiel.getRange("L:L").setNumberFormat("@");
        
        sheetZiel.getRange(2, 1, ergebnis.zeilen.length, HEADERS.length).setValues(ergebnis.zeilen);
        
        // Anpassung der Spaltenindizes nach Wegfall von Artikel-Bez 3
        sheetZiel.getRange(2, 6, ergebnis.zeilen.length, 1).setNumberFormat("dd.mm.yyyy");   // Datum (Spalte F)
        sheetZiel.getRange(2, 14, ergebnis.zeilen.length, 1).setNumberFormat("#,##0");       // Menge (Spalte N)
        sheetZiel.getRange(2, 15, ergebnis.zeilen.length, 2).setNumberFormat("#,##0.00 €"); // EP & Gesamt (Spalten O & P)
      }
    }
  });

  // ==========================================
  // 6. E-MAIL BENACHRICHTIGUNG VERSENDEN
  // ==========================================
  if (empfaenger && ueberschritteneArtikel.length > 0) {
    const startStr = Utilities.formatDate(new Date(rawBeginn), Session.getScriptTimeZone(), "dd.MM.yyyy");
    const endeStr = Utilities.formatDate(new Date(rawEnde), Session.getScriptTimeZone(), "dd.MM.yyyy");
    const tabellenLink = ss.getUrl() + "#gid=" + sheetBedingungen.getSheetId();

    const betreff = "Secabo Summer Sales 2026: Schwellenwert bei " + ueberschritteneArtikel.length + " Artikel(n) überschritten!";

    let mailBody = "Hallo,\n\nim Aktionszeitraum (" + startStr + " bis " + endeStr + ") wurde bei folgenden Artikeln der Schwellenwert überschritten:\n\n";

    ueberschritteneArtikel.forEach(item => {
      mailBody += "• " + item.name + " (" + item.artNr + ")\n" +
                  "  Aktuelle Verkäufe: " + item.verkaeufe + " Stk. | Schwellenwert: " + item.schwellwert + " Stk.\n\n";
    });

    mailBody += "----------------------------------------\n" +
                "Die Verkaufszeilen wurden auf den jeweiligen Artikelblättern aktualisiert.\n\n" +
                "Hier geht es direkt zur Tabelle:\n" + tabellenLink;

    MailApp.sendEmail(empfaenger, betreff, mailBody);
    Logger.log("Sammel-E-Mail erfolgreich gesendet an: " + empfaenger);
  } else {
    Logger.log("Keine Schwellenwerte überschritten oder kein Mailempfänger definiert.");
  }
}