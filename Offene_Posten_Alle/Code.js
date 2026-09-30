/**
 * Offene_Posten_Buchhaltung.gs
 * Lädt alle offenen Posten (ohne Vertreter-Filter) für die Buchhaltung.
 * Erstellt eine Kunden- und Vertreterübersicht mit Gesamtsaldo, überfälliger Summe,
 * noch nicht fälliger Summe, Guthaben (positiv ausgewiesen) sowie "Überfällig minus Guthaben".
 * NEU: Inklusive automatischer Summenzeile am Ende der Kunden- und Vertreter-Blätter.
 */

// =========================================================================
// 1. ZENTRALE KONFIGURATION & VOLLSTÄNDIGES VERTRETER-MAPPING
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
 * Hilfsfunktion zur Ermittlung des Vertreter-Namens
 */
function getVertreterName(vtrNr) {
  const nr = String(vtrNr || "").trim();
  return REPS[nr] ? REPS[nr].name : (nr || "Unbekannt");
}

// =========================================================================
// 2. HAUPTFUNKTION: OFFENE POSTEN ABRUFEN
// =========================================================================
function fetchOffenePostenBuchhaltung() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Namen der 3 Tabellenblätter
  const SHEET_OP = "Offene Posten (Buchhaltung)";
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
    showStatusBox("Buchhaltung", "Lade alle offenen Posten und aggregiere Daten...");
  } else {
    ss.toast("Lade alle offenen Posten...", "Buchhaltung", 5);
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

        // Einzelzeile für Sheet 1
        allRowsOP.push([
          vtrNr, vtrName, kredLimit, adrNr, kundengruppe, status,
          name2, name3, land,
          String(node.fldBelegNr || ""), String(node.fldAuftrNr || ""), opText,
          erstDatStr ? new Date(erstDatStr) : "",
          nettoTg,
          parseFloat(node.fldBezBet) || 0,
          opSaldo,
          node.fldMahnSt || 0,
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

        // --- 4. VERTRETER-AGGREGATION ---
        if (vtrNr) {
          if (!mapVtr.has(vtrNr)) {
            mapVtr.set(vtrNr, {
              vtrNr: vtrNr,
              vtrName: vtrName,
              gesamtSaldo: 0,
              ueberfaellig: 0,
              nichtFaellig: 0,
              guthaben: 0
            });
          }

          const v = mapVtr.get(vtrNr);
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
  
  // 5.1 Hauptblatt: Alle offenen Posten (ohne Gesamtsummenzeile)
  writeBuchhaltungSheet(ss, SHEET_OP, headersOP, allRowsOP, (sheet) => {
     sheet.getRange("A:B").setNumberFormat("@"); 
     sheet.getRange("D:D").setNumberFormat("@"); 
     sheet.getRange("J:K").setNumberFormat("@"); 
     if (allRowsOP.length > 0) {
       sheet.getRange(2, 13, allRowsOP.length, 1).setNumberFormat("yyyy-mm-dd"); 
       sheet.getRange(2, 15, allRowsOP.length, 2).setNumberFormat('#,##0.00 "€"'); 
       sheet.getRange(2, 18, allRowsOP.length, 1).setNumberFormat("yyyy-mm-dd"); 
       sheet.getRange(2, 19, allRowsOP.length, 1).setNumberFormat("yyyy-mm-dd"); 
       sheet.getRange(2, 20, allRowsOP.length, 1).setNumberFormat("0");          
     }
  }, false);

  // 5.2 Kundenblatt: Einzelzeilen + Gesamtsumme als letzte Zeile
  let sumKundeGesamt = 0;
  let sumKundeUeber = 0;
  let sumKundeNicht = 0;
  let sumKundeGut = 0;

  const rowsKunde = Array.from(mapKunde.values())
    .map(k => {
      const ueberfaelligNetto = k.ueberfaellig - k.guthaben;
      sumKundeGesamt += k.gesamtSaldo;
      sumKundeUeber += k.ueberfaellig;
      sumKundeNicht += k.nichtFaellig;
      sumKundeGut += k.guthaben;

      return [
        k.adrNr,
        k.name,
        k.gruppe,
        k.vtr,
        k.vtrName,
        k.gesamtSaldo,
        k.ueberfaellig,
        k.nichtFaellig,
        k.guthaben,
        ueberfaelligNetto
      ];
    })
    .sort((a, b) => b[5] - a[5]);

  // Gesamtsummenzeile für Kunden anhängen
  const sumKundeNetto = sumKundeUeber - sumKundeGut;
  rowsKunde.push([
    "Gesamtsumme", "", "", "", "",
    sumKundeGesamt, sumKundeUeber, sumKundeNicht, sumKundeGut, sumKundeNetto
  ]);

  writeBuchhaltungSheet(ss, SHEET_KUNDE, headersKunde, rowsKunde, (sheet) => {
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("D:E").setNumberFormat("@"); 
    if (rowsKunde.length > 0) {
      sheet.getRange(2, 6, rowsKunde.length, 5).setNumberFormat('#,##0.00 "€"');
    }
  }, true);

  // 5.3 Vertreterblatt: Einzelzeilen + Gesamtsumme als letzte Zeile
  let sumVtrGesamt = 0;
  let sumVtrUeber = 0;
  let sumVtrNicht = 0;
  let sumVtrGut = 0;

  const rowsVtr = Array.from(mapVtr.values())
    .map(v => {
      const ueberfaelligNetto = v.ueberfaellig - v.guthaben;
      sumVtrGesamt += v.gesamtSaldo;
      sumVtrUeber += v.ueberfaellig;
      sumVtrNicht += v.nichtFaellig;
      sumVtrGut += v.guthaben;

      return [
        v.vtrNr,
        v.vtrName,
        v.gesamtSaldo,
        v.ueberfaellig,
        v.nichtFaellig,
        v.guthaben,
        ueberfaelligNetto
      ];
    })
    .sort((a, b) => b[2] - a[2]);

  // Gesamtsummenzeile für Vertreter anhängen
  const sumVtrNetto = sumVtrUeber - sumVtrGut;
  rowsVtr.push([
    "Gesamtsumme", "",
    sumVtrGesamt, sumVtrUeber, sumVtrNicht, sumVtrGut, sumVtrNetto
  ]);

  writeBuchhaltungSheet(ss, SHEET_VTR, headersVtr, rowsVtr, (sheet) => {
    sheet.getRange("A:B").setNumberFormat("@");
    if (rowsVtr.length > 0) {
      sheet.getRange(2, 3, rowsVtr.length, 5).setNumberFormat('#,##0.00 "€"');
    }
  }, true);

  if (typeof closeStatusBox === "function") closeStatusBox();
  ss.toast("Daten wurden erfolgreich aktualisiert!", "Fertig", 5);
}

// =========================================================================
// 3. HILFSFUNKTION FÜR DAS SCHREIBEN & FORMATIEREN DER SHEETS
// =========================================================================
function writeBuchhaltungSheet(ss, sheetName, headers, data, applyFormats, hasTotalRow = false) {
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
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  
  if (applyFormats) applyFormats(sheet);
  
  sheet.setFrozenRows(1);
  
  const totalRowsCount = data.length;
  const dataRowsCount = hasTotalRow ? Math.max(0, totalRowsCount - 1) : totalRowsCount;

  // Filter anlegen NUR auf den Datenbereich (ohne Gesamtsummenzeile),
  // damit die Gesamtsummenzeile beim Sortieren/Filtern nicht verrutscht!
  if (dataRowsCount > 0) {
    const filterRange = sheet.getRange(1, 1, dataRowsCount + 1, headers.length);
    filterRange.createFilter();
  }

  // Rahmen ziehen
  if (totalRowsCount > 0) {
    const fullRange = sheet.getRange(1, 1, totalRowsCount + 1, headers.length);
    fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Hervorhebung der Gesamtsummenzeile (falls vorhanden)
  if (hasTotalRow && totalRowsCount > 0) {
    const sumRowIndex = totalRowsCount + 1; // Row 1 is Header
    const sumRange = sheet.getRange(sumRowIndex, 1, 1, headers.length);
    sumRange.setFontWeight("bold")
            .setBackground("#e2e8f0")
            .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  sheet.autoResizeColumns(1, headers.length);
}

// =========================================================================
// 4. MENU-BUTTON HINZUFÜGEN
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update')
    .addItem('OP Aktualisieren', 'fetchOffenePostenBuchhaltung')
    .addToUi();
}