/**
 * Offene_Posten_Buchhaltung.gs
 * Lädt alle offenen Posten (ohne Vertreter-Filter) für die Buchhaltung.
 * Erstellt eine Beleg-, Kunden- und Vertreterübersicht mit Gesamtsaldo, überfälliger Summe,
 * noch nicht fälliger Summe, Guthaben (positiv ausgewiesen) sowie "Überfällig minus Guthaben".
 * 
 * EIGENSCHAFTEN:
 * - Dynamische Summenzeilen per SUBTOTAL(109; ...) mit Semikolon (;) für deutsche Ländereinstellung
 * - Vollständige Erfassung unzugeordneter Belege ("Ohne Vertreter")
 * - Zeilenfarben (Rot = Mahnstufe > 0, Orange = Guthaben < 0, Grün = Noch nicht fällig)
 * - Menü "Update" in Google Sheets
 */

// =========================================================================
// 1. ZENTRALE KONFIGURATION & VERTRETER-MAPPING
// =========================================================================
const CONFIG = {
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  AUSGESCHLOSSENE_ADRESSEN: [] 
};

// Zuordnung Vertreter-Nummer -> Name & Sprache
const REPS = {
  "28": { name: "Alexis Fonte", lang: "DE" },
  "32": { name: "Oliver Bayer", lang: "DE" },
  "36": { name: "Imad Nassef", lang: "DE" },
  "43": { name: "Andreas Steinert", lang: "DE" },
  "46": { name: "Stefanie Binder", lang: "DE" },
  "50": { name: "Alicia Hierl", lang: "DE" },
  "55": { name: "Jessica Erb", lang: "DE" },
  "56": { name: "Robin Carter Browne", lang: "EN" },
  "59": { name: "Aneta Nedyalkova", lang: "DE" },
  "60": { name: "Rado Kabakov", lang: "EN" }
};

/**
 * Hilfsfunktion zur Ermittlung des Vertreter-Namens mit sauberem Fallback
 */
function getVertreterName(vtrNr) {
  if (vtrNr === null || vtrNr === undefined) return "Ohne Vertreter";
  
  const nr = String(vtrNr).trim();
  
  if (!nr || nr === "0" || nr === "00" || nr === "0.0") {
    return "Ohne Vertreter";
  }
  
  if (REPS[nr]) {
    return REPS[nr].name;
  }
  
  return "Vertreter " + nr;
}

// =========================================================================
// 2. HAUPTFUNKTION: OFFENE POSTEN ABRUFEN
// =========================================================================
function fetchOffenePostenBuchhaltung() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Namen der 3 Tabellenblätter
  const SHEET_OP = "Offene Posten (Belege)";
  const SHEET_KUNDE = "OP-Summen pro Kunde";
  const SHEET_VTR = "OP-Summen pro Vertreter";

  // --- KOPFZEILEN ---
  const headersOP = [
    "Vertreter-Nr.", "Vertreter", "Kreditlimit", "Adress-Nr. (Kunde)", "Kundengruppe", "Status",
    "Name 2", "Name 3", "Land",
    "Beleg-Nr.", "Auftrags-Nr.", "OP-Text",
    "Erstelldatum", "Netto Tage", "Bezahlt-Betrag",
    "OP-Saldo Betrag", "Mahnstufe", "Mahndatum",
    "Fälligkeitsdatum", "Überfällig seit (Tagen)"
  ];

  const headersKunde = [
    "Adress-Nr.", "Name (Anschrift 2)", "Kundengruppe", "Vertreter-Nr.", "Vertreter", 
    "Gesamtsaldo", "Überfällig", "Nicht fällig", "Guthaben", "Überfällig abzgl. Guthaben"
  ];
  
  const headersVtr = [
    "Vertreter-Nr.", "Vertreter", 
    "Gesamtsaldo", "Überfällig", "Nicht fällig", "Guthaben", "Überfällig abzgl. Guthaben"
  ];

  const typeMapping = { "0": "Endkunde", "1": "A", "2": "B", "3": "C" };
  const ausgeschlosseneAdressen = CONFIG.AUSGESCHLOSSENE_ADRESSEN || [];

  if (typeof showStatusBox === "function") {
    showStatusBox("Update", "Lade alle offenen Posten und aggregiere Daten...");
  } else {
    ss.toast("Lade alle offenen Posten...", "Update", 5);
  }

  // --- GRAPHQL QUERY ---
  const queryAllOP = `
    query GetAllOPBuchhaltung($cursor: String) {
      tblOffenerPosten {
        conRead(first: 100, after: $cursor) {
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
              rowAdresse {
                fldVtrNr
                fldKredLimit
                fldStatus
                fldAbwArtDatGrp
                rowsAnschriften {
                  fldNa2
                  fldNa3
                  fldLandBez
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
  
  const allRowsOP = [];
  const mapKunde = new Map();
  const mapVtr = new Map();

  const heute = new Date();
  heute.setHours(0, 0, 0, 0);

  const startTime = Date.now();
  const MAX_TIME_MS = 5.5 * 60 * 1000; 

  // --- DATEN ABRUFEN ---
  while (hasNextPage) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("Zeitlimit für OP-Buchhaltung erreicht. Speichere bisher geladene Daten.");
      break;
    }

    const payload = { query: queryAllOP, variables: { cursor: cursor } };
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      if (response.getResponseCode() !== 200) break;
      
      const json = JSON.parse(response.getContentText());
      if (json.errors) break;

      const conRead = json.data?.tblOffenerPosten?.conRead || {};
      const edges = conRead.edges || [];

      if (edges.length === 0) break;

      edges.forEach(edge => {
        const node = edge.node || {};
        const adrNr = String(node.fldAdrNr || "").trim();
        
        // --- 1. Adressdaten extrahieren ---
        let vtrNr = "";
        let kredLimit = 0;
        let status = "";
        let kundengruppe = "";
        let name2 = "";
        let name3 = "";
        let land = "";

        const addrNode = node.rowAdresse;
        if (addrNode) {
          vtrNr = String(addrNode.fldVtrNr || "").trim();
          kredLimit = addrNode.fldKredLimit || 0;
          status = String(addrNode.fldStatus || "").trim();
          
          let abwGrp = String(addrNode.fldAbwArtDatGrp || "").trim();
          kundengruppe = typeMapping[abwGrp] || abwGrp;

          const anschriften = addrNode.rowsAnschriften || [];
          if (anschriften.length > 0) {
            name2 = String(anschriften[0].fldNa2 || "");
            name3 = String(anschriften[0].fldNa3 || "");
            land = String(anschriften[0].fldLandBez || "");
          }
        }

        if (status === "Lieferant" || ausgeschlosseneAdressen.includes(adrNr)) return;

        const vtrName = getVertreterName(vtrNr);

        // --- 2. Fälligkeit und Überfälligkeit berechnen ---
        const erstDatStr = node.fldErstDat || "";
        const nettoTg = parseInt(node.fldNettoTg, 10) || 0;
        let faelligkeitStr = "";
        let faelligSeitTage = "";
        let isOverdue = false;

        if (erstDatStr) {
          const erstelldatum = new Date(erstDatStr);
          if (!isNaN(erstelldatum.getTime())) {
            const faelligkeitsDatum = new Date(erstelldatum);
            faelligkeitsDatum.setDate(faelligkeitsDatum.getDate() + nettoTg);
            
            faelligkeitStr = Utilities.formatDate(faelligkeitsDatum, Session.getScriptTimeZone(), "yyyy-MM-dd");
            
            faelligkeitsDatum.setHours(0, 0, 0, 0);
            if (faelligkeitsDatum < heute) {
              isOverdue = true;
              const diffTime = heute.getTime() - faelligkeitsDatum.getTime();
              faelligSeitTage = Math.floor(diffTime / (1000 * 60 * 60 * 24));
            }
          }
        }

        const opSaldo = parseFloat(node.fldOPSaldoBet) || 0;
        const opText = String(node.fldText || "").trim();
        const mahnstufe = parseInt(node.fldMahnSt, 10) || 0;

        // Einzelzeile für Sheet 1 (Belege)
        allRowsOP.push([
          vtrNr, vtrName, kredLimit, adrNr, kundengruppe, status,
          name2, name3, land,
          String(node.fldBelegNr || ""), String(node.fldAuftrNr || ""), opText,
          erstDatStr ? new Date(erstDatStr) : "",
          nettoTg,
          parseFloat(node.fldBezBet) || 0,
          opSaldo,
          mahnstufe,
          node.fldMahnDat ? new Date(node.fldMahnDat) : "",
          faelligkeitStr,
          faelligSeitTage
        ]);

        // --- 3. KUNDEN-AGGREGATION ---
        if (adrNr) {
          if (!mapKunde.has(adrNr)) {
            mapKunde.set(adrNr, {
              adrNr: adrNr,
              name: name2,
              gruppe: kundengruppe,
              vtr: vtrNr,
              vtrName: vtrName,
              gesamtSaldo: 0,
              ueberfaellig: 0,
              nichtFaellig: 0,
              guthaben: 0
            });
          }

          const k = mapKunde.get(adrNr);
          k.gesamtSaldo += opSaldo;

          if (opSaldo < 0) {
            k.guthaben += Math.abs(opSaldo);
          } else {
            if (isOverdue) {
              k.ueberfaellig += opSaldo;
            } else {
              k.nichtFaellig += opSaldo;
            }
          }
        }

        // --- 4. VERTRETER-AGGREGATION (Inklusive "Ohne Vertreter") ---
        const vtrKey = (!vtrNr || vtrNr === "0" || vtrNr === "00" || vtrNr === "0.0") ? "0" : vtrNr;
        if (!mapVtr.has(vtrKey)) {
          mapVtr.set(vtrKey, {
            vtrNr: vtrKey,
            vtrName: getVertreterName(vtrKey),
            gesamtSaldo: 0,
            ueberfaellig: 0,
            nichtFaellig: 0,
            guthaben: 0
          });
        }

        const v = mapVtr.get(vtrKey);
        v.gesamtSaldo += opSaldo;

        if (opSaldo < 0) {
          v.guthaben += Math.abs(opSaldo);
        } else {
          if (isOverdue) {
            v.ueberfaellig += opSaldo;
          } else {
            v.nichtFaellig += opSaldo;
          }
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
    } catch (e) {
      Logger.log("Fehler bei Offenen Posten Buchhaltung: " + e.message);
      break;
    }
  }

  // --- 5. DATEN IN DIE SHEETS SCHREIBEN ---
  
  // 5.1 Hauptblatt: Offene Posten (Belege) mit Farbregeln
  writeBuchhaltungSheet(ss, SHEET_OP, headersOP, allRowsOP, (sheet) => {
     sheet.getRange("A:B").setNumberFormat("@"); 
     sheet.getRange("D:D").setNumberFormat("@"); 
     sheet.getRange("J:K").setNumberFormat("@"); 
     if (allRowsOP.length > 0) {
       sheet.getRange(2, 13, allRowsOP.length, 1).setNumberFormat("yyyy-mm-dd"); 
       sheet.getRange(2, 14, allRowsOP.length, 1).setNumberFormat("0");          
       sheet.getRange(2, 15, allRowsOP.length, 2).setNumberFormat('#,##0.00 "€"'); 
       sheet.getRange(2, 17, allRowsOP.length, 1).setNumberFormat("0");          
       sheet.getRange(2, 18, allRowsOP.length, 2).setNumberFormat("yyyy-mm-dd"); 
       sheet.getRange(2, 20, allRowsOP.length, 1).setNumberFormat("0");          

       // --- ZEILENFARBEN ANWENDEN ---
       const backgrounds = [];
       for (let i = 0; i < allRowsOP.length; i++) {
         const row = allRowsOP[i];
         const opSaldo = row[15];         // Spalte P (OP-Saldo Betrag)
         const mahnstufe = row[16];       // Spalte Q (Mahnstufe)
         const faelligTage = row[19];     // Spalte T (Überfällig seit Tagen)

         let rowColor = (i % 2 === 0) ? "#ffffff" : "#f9fbfd"; 

         if (mahnstufe > 0) {
           rowColor = "#fce8e6"; // ROT: Mahnstufe > 0
         } else if (opSaldo < 0) {
           rowColor = "#fce5cd"; // ORANGE: OP-Saldo Betrag negativ
         } else if (faelligTage === "" || faelligTage === null || faelligTage === undefined) {
           rowColor = "#d9ead3"; // GRÜN: Nicht überfällig
         }

         backgrounds.push(new Array(headersOP.length).fill(rowColor));
       }

       sheet.getRange(2, 1, allRowsOP.length, headersOP.length).setBackgrounds(backgrounds);
     }
  }, null);

  // 5.2 Kundenblatt: OP-Summen pro Kunde mit DYNAMISCHER SUMMENZEILE (mit Semikolon)
  const rowsKunde = Array.from(mapKunde.values())
    .map(k => [
      k.adrNr,
      k.name,
      k.gruppe,
      k.vtr,
      k.vtrName,
      k.gesamtSaldo,
      k.ueberfaellig,
      k.nichtFaellig,
      k.guthaben,
      k.ueberfaellig - k.guthaben
    ])
    .sort((a, b) => b[5] - a[5]);

  writeBuchhaltungSheet(ss, SHEET_KUNDE, headersKunde, rowsKunde, (sheet, dataCount) => {
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("D:E").setNumberFormat("@"); 
    if (dataCount > 0) {
      sheet.getRange(2, 6, dataCount + 1, 5).setNumberFormat('#,##0.00 "€"');
    }
  }, {
    startTextCol: 1, textValues: [["Gesamtsumme", "", "", "", ""]],
    formulaStartCol: 6, formulaCols: ["F", "G", "H", "I", "J"]
  });

  // 5.3 Vertreterblatt: OP-Summen pro Vertreter mit DYNAMISCHER SUMMENZEILE (mit Semikolon)
  const rowsVtr = Array.from(mapVtr.values())
    .map(v => [
      v.vtrNr,
      v.vtrName,
      v.gesamtSaldo,
      v.ueberfaellig,
      v.nichtFaellig,
      v.guthaben,
      v.ueberfaellig - v.guthaben
    ])
    .sort((a, b) => b[2] - a[2]);

  writeBuchhaltungSheet(ss, SHEET_VTR, headersVtr, rowsVtr, (sheet, dataCount) => {
    sheet.getRange("A:B").setNumberFormat("@");
    if (dataCount > 0) {
      sheet.getRange(2, 3, dataCount + 1, 5).setNumberFormat('#,##0.00 "€"');
    }
  }, {
    startTextCol: 1, textValues: [["Gesamtsumme", ""]],
    formulaStartCol: 3, formulaCols: ["C", "D", "E", "F", "G"]
  });

  if (typeof closeStatusBox === "function") closeStatusBox();
  ss.toast("Daten wurden erfolgreich aktualisiert!", "Fertig", 5);
}

// =========================================================================
// 3. HILFSFUNKTION FÜR DAS SCHREIBEN, FORMELN & FORMATIEREN DER SHEETS
// =========================================================================
function writeBuchhaltungSheet(ss, sheetName, headers, data, applyFormats, dynamicTotalConfig) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();
  
  // Header schreiben und stylen
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  
  const dataCount = data.length;

  if (dataCount > 0) {
    // 1. Datenzeilen schreiben
    sheet.getRange(2, 1, dataCount, headers.length).setValues(data);

    // 2. Dynamische Summenzeile anfügen
    if (dynamicTotalConfig) {
      const sumRowIdx = dataCount + 2; // Row 1 ist Header
      
      // Beschriftung ("Gesamtsumme") setzen
      if (dynamicTotalConfig.textValues) {
        sheet.getRange(sumRowIdx, dynamicTotalConfig.startTextCol, 1, dynamicTotalConfig.textValues[0].length)
             .setValues(dynamicTotalConfig.textValues);
      }
      
      // Dynamische Formel mit Semikolon (;) erzeugen und mit setFormulas eintragen
      if (dynamicTotalConfig.formulaCols && dynamicTotalConfig.formulaCols.length > 0) {
        const formulas = [dynamicTotalConfig.formulaCols.map(col => `=SUBTOTAL(109; ${col}2:${col}${dataCount + 1})`)];
        sheet.getRange(sumRowIdx, dynamicTotalConfig.formulaStartCol, 1, formulas[0].length)
             .setFormulas(formulas);
      }

      // Styling der Summenzeile
      const sumRange = sheet.getRange(sumRowIdx, 1, 1, headers.length);
      sumRange.setFontWeight("bold")
              .setBackground("#e2e8f0")
              .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  if (applyFormats) applyFormats(sheet, dataCount);
  
  sheet.setFrozenRows(1);
  
  // Filter legen wir EXKLUSIV auf den Datenbereich (ohne Summenzeile)
  if (dataCount > 0) {
    const filterRange = sheet.getRange(1, 1, dataCount + 1, headers.length);
    filterRange.createFilter();

    // Rahmen um Daten + Summenzeile
    const totalRowsCount = dynamicTotalConfig ? dataCount + 1 : dataCount;
    const fullRange = sheet.getRange(1, 1, totalRowsCount + 1, headers.length);
    fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.autoResizeColumns(1, headers.length);
}

// =========================================================================
// 4. MENU-BUTTON HINZUFÜGEN ("Update")
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update')
    .addItem('OP Aktualisieren', 'fetchOffenePostenBuchhaltung')
    .addToUi();
}