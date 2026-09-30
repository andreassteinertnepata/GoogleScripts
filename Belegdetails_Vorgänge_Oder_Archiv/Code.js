/**
 * ZENTRALE KONFIGURATION & CREDENTIALS (Nepata Datahub)
 */
const CONFIG = {
  API: {
    URL: 'https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql',
    TOKEN: 'e12Bfv!@Ss#asrpPFjucm8a8'
  },

  // Relevante Belegarten für Vorzeichenumkehr (Gutschriften, Rechnungskorrekturen, Erstattungen)
  CORRECTION_TYPES: ['90', '123', '156'],

  // Internationales Vertriebsteam (Zentrale Vertreter-Zuordnung)
  SALES_TEAM: {
    '28': { name: 'Alexis Fonte', lang: 'Deutsch' },
    '36': { name: 'Imad Nassef', lang: 'Deutsch' },
    '43': { name: 'Andreas Steinert', lang: 'Deutsch' },
    '46': { name: 'Stefanie Binder', lang: 'Deutsch' },
    '56': { name: 'Robin Carter Browne', lang: 'Englisch' },
    '59': { name: 'Aneta Nedyalkova', lang: 'Deutsch' },
    '60': { name: 'Rado Kabakov', lang: 'Englisch' }
  }
};

/**
 * NATIVE NEPATA DATAHUB GRAPHQL QUERIES
 */

// 0 = Vorgänge (tblVorgang - Offene Belege)
const GRAPHQL_VORGANG = `
  query GetVorgangByBelegNr($docNo: String!) {
    tblVorgang {
      conRead(
        fastFilter: {
          eq: [{ field: fldBelegNr }, { value: { string: $docNo } }]
        }
      ) {
        edges {
          node {
            fldAdrNr
            fldArt
            fldAuftrNr
            fldBelegNr
            fldDat
            fldVtrNr
            rowsPositions {
              fldAbrPosKz
              fldArtNr
              fldBez
              fldMge
              fldEPrNt
              fldEEkRoh
              rowArtikel {
                fldKuBez1
                fldEkRoh
              }
            }
          }
        }
      }
    }
  }
`;

// 1 = Archiv (tblVorgangArchiv - Abgeschlossene Belege)
const GRAPHQL_ARCHIV = `
  query GetArchivByBelegNr($docNo: String!) {
    tblVorgangArchiv {
      conRead(
        fastFilter: {
          eq: [{ field: fldBelegNr }, { value: { string: $docNo } }]
        }
      ) {
        edges {
          node {
            fldAdrNr
            fldArt
            fldAuftrNr
            fldBelegNr
            fldDat
            fldStorniertKz
            fldVtrNr
            rowsPositions {
              fldAbrPosKz
              fldArtNr
              fldBez
              fldMge
              fldEPrNt
              fldEEkRoh
              rowArtikel {
                fldKuBez1
                fldEkRoh
              }
            }
          }
        }
      }
    }
  }
`;


/**
 * CUSTOM MENU: Erzeugt ein Menü in Google Sheets beim Öffnen der Datei
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Nepata Datahub')
    .addItem('Daten für Belegnummer abrufen', 'fetchPositionsByDocumentNumber')
    .addSeparator()
    .addItem('Info-Blatt erstellen/aktualisieren', 'createInfoSheet')
    .addToUi();
}


/**
 * HAUPTFUNKTION: Dialog zur Eingabe von Datenquelle (0/1) und Belegnummer
 */
function fetchPositionsByDocumentNumber() {
  const ui = SpreadsheetApp.getUi();

  // 1. Wahl der Datenquelle: 0 = Vorgänge, 1 = Archiv
  const typeResponse = ui.prompt(
    'Datenquelle wählen',
    'Bitte wähle den Belegtyp:\n0 = Vorgänge (offene Belege)\n1 = Archiv (abgeschlossene Belege)',
    ui.ButtonSet.OK_CANCEL
  );

  if (typeResponse.getSelectedButton() !== ui.Button.OK) return;

  const selection = typeResponse.getResponseText().trim();
  if (selection !== '0' && selection !== '1') {
    ui.alert('Ungültige Eingabe', 'Bitte gib exakt "0" für Vorgänge oder "1" für Archiv ein.', ui.ButtonSet.OK);
    return;
  }

  // 2. Eingabe der Belegnummer
  const docResponse = ui.prompt(
    'Belegnummer eingeben',
    'Bitte Belegnummer eingeben (z. B. 702604347 oder 302604960):',
    ui.ButtonSet.OK_CANCEL
  );

  if (docResponse.getSelectedButton() !== ui.Button.OK) return;

  const docNo = docResponse.getResponseText().trim();
  if (!docNo) {
    ui.alert('Hinweis', 'Es wurde keine Belegnummer eingegeben.', ui.ButtonSet.OK);
    return;
  }

  const isArchive = (selection === '1');
  const positions = loadPositionsFromDatahub(docNo, isArchive);

  if (!positions || positions.length === 0) {
    const areaName = isArchive ? 'Archiv (1)' : 'Vorgänge (0)';
    ui.alert(
      'Keine Daten gefunden',
      `Keine Positionen für Belegnummer "${docNo}" im Bereich "${areaName}" gefunden.\n\n` +
      `Mögliche Ursachen:\n` +
      `- Bei Vorgänge (0): Der Beleg ist bereits ausgeliefert/fakturiert und liegt im Archiv (1).\n` +
      `- Bei Archiv (1): Der Beleg ist storniert worden (fldStorniertKz = true) oder existiert nicht.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 3. Ausgabe im Sheet mit Belegnummer als Blattname
  writePositionsToSheet(positions, docNo);
  ui.alert('Erfolg', `${positions.length} Position(en) für Belegnummer "${docNo}" im Blatt "${docNo}" geladen.`, ui.ButtonSet.OK);
}

/**
 * Abfrage- und Aufbereitungslogik
 */
function loadPositionsFromDatahub(docNo, isArchive) {
  const positions = [];

  if (!isArchive) {
    // VORGÄNGE (0)
    const data = executeDatahubGraphQL(GRAPHQL_VORGANG, { docNo: docNo });
    const edges = data.tblVorgang?.conRead?.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      extractPositionsFromNode(node, docNo, false, positions);
    });

  } else {
    // ARCHIV (1)
    const data = executeDatahubGraphQL(GRAPHQL_ARCHIV, { docNo: docNo });
    const edges = data.tblVorgangArchiv?.conRead?.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};

      // REGEL 1: Stornierte Archivbelege strikt ignorieren (fldStorniertKz)
      if (node.fldStorniertKz === true) {
        return;
      }

      extractPositionsFromNode(node, docNo, true, positions);
    });
  }

  return positions;
}

/**
 * Extrahiert Positionen aus dem Belegknoten
 */
function extractPositionsFromNode(node, docNo, isArchive, positions) {
  const docType = String(node.fldArt || '').trim();
  // REGEL 3: Kaufmännisches Belegdatum verwenden (fldDat)
  const docDate = node.fldDat ? new Date(node.fldDat) : '';
  const salesRepId = String(node.fldVtrNr || '').trim();

  // Vertreter-Name aus CONFIG auflösen
  const repConfig = CONFIG.SALES_TEAM[salesRepId];
  const salesRepName = repConfig ? repConfig.name : (salesRepId ? `Vertreter ${salesRepId}` : 'Nicht zugewiesen');

  const rows = node.rowsPositions || [];

  rows.forEach(pos => {
    let quantity = Number(pos.fldMge) || 0;

    // REGEL 2: Vorzeichenumkehr bei Gutschriften/Korrekturen (z. B. "90", "123", "156")
    if (CONFIG.CORRECTION_TYPES.includes(docType)) {
      quantity = -Math.abs(quantity);
    }

    const salePrice = Number(pos.fldEPrNt) || 0;
    const costPrice = Number(pos.fldEEkRoh || pos.rowArtikel?.fldEkRoh) || 0;

    const totalCost = quantity * costPrice;
    const totalSale = quantity * salePrice;
    const grossProfit = totalSale - totalCost;
    const marginPercent = totalSale !== 0 ? grossProfit / totalSale : 0;

    let itemName = pos.rowArtikel?.fldKuBez1 || '';
    if (!itemName) {
      itemName = Array.isArray(pos.fldBez) ? pos.fldBez.join(' ') : String(pos.fldBez || '');
    }

    positions.push({
      documentNumber: node.fldBelegNr || docNo,
      documentType: docType,
      documentDate: docDate,
      itemNumber: String(pos.fldArtNr || ''),
      itemName: itemName,
      quantity: quantity,
      costPrice: costPrice,
      totalCost: totalCost,
      salePrice: salePrice,
      totalSale: totalSale,
      grossProfit: grossProfit,
      marginPercent: marginPercent,
      salesRep: salesRepName,
      source: isArchive ? 'Archiv' : 'Vorgang'
    });
  });
}

/**
 * Sendet GraphQL Post Requests an die Nepata API
 */
function executeDatahubGraphQL(query, variables) {
  const payload = JSON.stringify({
    query: query,
    variables: variables
  });

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-API-Token': CONFIG.API.TOKEN,
      'x-hasura-admin-secret': CONFIG.API.TOKEN
    },
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.API.URL, options);
    const result = JSON.parse(response.getContentText());

    if (result.errors) {
      Logger.log('GraphQL Error: ' + JSON.stringify(result.errors));
      return {};
    }

    return result.data || {};
  } catch (error) {
    Logger.log('HTTP Fetch Error: ' + error.toString());
    return {};
  }
}

/**
 * Erstellt/Öffnet das Blatt mit der Belegnummer und schreibt Daten inkl. direkter Gesamtsummen
 */
function writePositionsToSheet(positions, docNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = String(docNo).trim();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // Deutsche Spaltenüberschriften
  const headers = [
    'Belegnummer',
    'Belegart',
    'Belegdatum',
    'Artikelnummer',
    'Artikelname',
    'Stückzahl',
    'EK einzeln',
    'EK gesamt',
    'VK einzeln',
    'VK gesamt',
    'Roherlös',
    'Marge (%)',
    'Betreuer',
    'Quelle'
  ];

  const rows = [headers];

  // Variablen für die direkte Summenberechnung in JavaScript
  let totalQuantity = 0;
  let totalCostSum = 0;
  let totalSaleSum = 0;
  let totalGrossProfitSum = 0;

  positions.forEach(pos => {
    rows.push([
      pos.documentNumber,
      pos.documentType,
      pos.documentDate,
      pos.itemNumber,
      pos.itemName,
      pos.quantity,
      pos.costPrice,
      pos.totalCost,
      pos.salePrice,
      pos.totalSale,
      pos.grossProfit,
      pos.marginPercent,
      pos.salesRep,
      pos.source
    ]);

    totalQuantity += pos.quantity;
    totalCostSum += pos.totalCost;
    totalSaleSum += pos.totalSale;
    totalGrossProfitSum += pos.grossProfit;
  });

  const lastDataRow = rows.length;

  // Schreiben der Datenzeilen
  const range = sheet.getRange(1, 1, lastDataRow, headers.length);
  range.setValues(rows);

  // Kopfzeilen-Formatierung
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');

  if (lastDataRow > 1) {
    // Belegdatum (Spalte C)
    sheet.getRange(2, 3, lastDataRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    // Stückzahl (Spalte F)
    sheet.getRange(2, 6, lastDataRow - 1, 1).setNumberFormat('#,##0.00');
    // EK einzeln, EK gesamt, VK einzeln, VK gesamt, Roherlös (Spalten G, H, I, J, K)
    sheet.getRange(2, 7, lastDataRow - 1, 5).setNumberFormat('#,##0.00 "€"');
    // Marge (%) (Spalte L)
    sheet.getRange(2, 12, lastDataRow - 1, 1).setNumberFormat('0.00%');

    // ==========================================
    // SUMMENZEILE AN ZEILE (lastDataRow + 1)
    // ==========================================
    const sumRowIndex = lastDataRow + 1;

    // Gesamtmarge direkt berechnen (kein Formeleintrags-Problem mehr)
    const totalMarginPercent = totalSaleSum !== 0 ? totalGrossProfitSum / totalSaleSum : 0;

    // Errechnete Festwerte direkt in die Summenzeile schreiben
    sheet.getRange(sumRowIndex, 5).setValue('Gesamtsumme');
    sheet.getRange(sumRowIndex, 6).setValue(totalQuantity);        // Stückzahl
    sheet.getRange(sumRowIndex, 8).setValue(totalCostSum);         // EK gesamt
    sheet.getRange(sumRowIndex, 10).setValue(totalSaleSum);        // VK gesamt
    sheet.getRange(sumRowIndex, 11).setValue(totalGrossProfitSum); // Roherlös (€)
    sheet.getRange(sumRowIndex, 12).setValue(totalMarginPercent);  // Marge (%) als Zahlenwert

    // Formatierung der Summenzeile
    const sumRange = sheet.getRange(sumRowIndex, 1, 1, headers.length);
    sumRange
      .setFontWeight('bold')
      .setBackground('#E8EEF5');

    // Kaufmännischer Doppelrahmen unten
    sumRange.setBorder(
      true, null, true, null, null, null,
      '#1F4E78',
      SpreadsheetApp.BorderStyle.DOUBLE
    );

    // Zahlenformate für die Summenfelder
    sheet.getRange(sumRowIndex, 6).setNumberFormat('#,##0.00');
    sheet.getRange(sumRowIndex, 8).setNumberFormat('#,##0.00 "€"');
    sheet.getRange(sumRowIndex, 10, 1, 2).setNumberFormat('#,##0.00 "€"'); // VK gesamt & Roherlös
    sheet.getRange(sumRowIndex, 12).setNumberFormat('0.00%');             // Marge (%)
  }

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Erstellt oder aktualisiert das Info-Blatt mit dynamischen Variablen aus der CONFIG
 */
function createInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Info';
  let sheet = ss.getSheetByName(sheetName);

  // Wenn das Blatt noch nicht existiert, erstelle es an erster Stelle (Index 0)
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, 0);
  } else {
    // Wenn es existiert, Inhalt und Formate für ein Update löschen
    sheet.clear();
  }

  const rows = [];

  // --- 1. Titel ---
  rows.push(['📊 Nepata Datahub - Skript Informationen', '']);
  rows.push(['', '']);

  // --- 2. Beschreibung ---
  rows.push(['Was macht dieses Skript?', 'Das Skript ruft Belegpositionen (Vorgänge oder Archiv) direkt aus dem Nepata Datahub ab, berechnet automatisch die Einkaufs- und Verkaufssummen, den Roherlös sowie die Marge. Die Daten werden formatiert in ein eigenes Tabellenblatt geschrieben, welches nach der Belegnummer benannt ist.']);
  rows.push(['Ablauf & Logik:', '1. Auswahl der Datenquelle (0 = Vorgänge/Offen, 1 = Archiv/Abgeschlossen)\n2. Abfrage via GraphQL API.\n3. Belege, die im Archiv als storniert markiert sind, werden ignoriert.\n4. Bei Gutschriften/Korrekturen wird das Vorzeichen (Stückzahl) automatisch umgekehrt.\n5. Zuordnung der Vertreter über die zentrale ID-Matrix.']);
  rows.push(['', '']);

  // --- 3. Dynamische Config-Werte (API & Belegarten) ---
  rows.push(['⚙️ Dynamische Konfigurationswerte (Live aus dem Code)', '']);
  rows.push(['API Endpoint:', CONFIG.API.URL]);
  rows.push(['Belegarten für Vorzeichenumkehr (z.B. Gutschriften):', CONFIG.CORRECTION_TYPES.join(', ')]);
  rows.push(['', '']);

  // --- 4. Dynamische Config-Werte (Vertriebsteam) ---
  rows.push(['👥 Hinterlegtes Vertriebsteam (Zentrale Zuordnung)', '']);
  
  // Schleife durch das CONFIG.SALES_TEAM Objekt
  for (const [id, data] of Object.entries(CONFIG.SALES_TEAM)) {
    rows.push([`Vertreter-ID: ${id}`, `${data.name} (Sprache: ${data.lang})`]);
  }

  // --- Daten in das Tabellenblatt schreiben ---
  const range = sheet.getRange(1, 1, rows.length, 2);
  range.setValues(rows);

  // --- Formatierung des Info-Blatts ---
  
  // Spaltenbreiten & Textumbruch
  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 600);
  range.setWrap(true);
  range.setVerticalAlignment('middle');

  // Hauptüberschrift (Zeile 1)
  sheet.getRange(1, 1, 1, 2)
    .merge()
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // Dynamische Teilüberschriften formatieren
  const subheaderBackground = '#E8EEF5';

  for(let i = 0; i < rows.length; i++) {
    const leftCellText = rows[i][0];
    const rowNum = i + 1;

    // Teilüberschriften erkennen und grau hinterlegen
    if (leftCellText.startsWith('⚙️') || leftCellText.startsWith('👥')) {
       sheet.getRange(rowNum, 1, 1, 2)
         .merge()
         .setFontWeight('bold')
         .setBackground(subheaderBackground);
    } 
    // Normale Bezeichner fett machen (Spalte A)
    else if (leftCellText !== '') {
       sheet.getRange(rowNum, 1).setFontWeight('bold');
    }
  }

  // Gitterlinien (Gridlines) im Info-Blatt ausblenden für einen cleanen Look
  sheet.setHiddenGridlines(true);
  
  // Benachrichtigung
  SpreadsheetApp.getActiveSpreadsheet().toast('Info-Blatt wurde erfolgreich generiert!', 'Erfolg', 3);
}