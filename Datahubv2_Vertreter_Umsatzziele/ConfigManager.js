/**
 * ConfigManager.js - Verwaltung aller Ziel-Konfigurationsblätter
 */

/**
 * Liest Artikel direkt aus dem Blatt "Artikel" (Fallback)
 */
function getArticlesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ARTIKEL);
  const articles = [];

  if (!sheet) return articles;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const artNr = String(data[i][0]).trim();
    const bez = String(data[i][1]).trim();
    const katalog = String(data[i][2]).trim();

    if (artNr) {
      articles.push({
        fldArtNr: artNr,
        fldKuBez1: bez,
        fldKatalog: katalog
      });
    }
  }

  return articles;
}

function syncAndGetTargetsConfig(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);

  const headers = [
    "Vertreter-ID", 
    "Name", 
    "Durchschnitt 12M (€)", 
    "Monatsziel (€)"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_CONFIG);
    sheet.appendRow(headers);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
  }

  const existingData = sheet.getDataRange().getValues();
  const userConfigMap = new Map();

  for (let i = 1; i < existingData.length; i++) {
    const id = String(existingData[i][0]).trim();
    if (!id) continue;
    
    const valRaw = existingData[i][3];
    const userZiel = (valRaw !== "" && valRaw !== null && valRaw !== undefined) ? parseFloat(valRaw) : null;

    userConfigMap.set(id, userZiel);
  }

  const rows = [];
  const targets = {};

  Object.keys(CONFIG.SALES_REPS).forEach(function(id) {
    const repName = CONFIG.SALES_REPS[id].name;
    const repStats = stats[id] || { umsatz12M: 0 };
    
    const avg12M = repStats.umsatz12M / 12.0;
    const savedZiel = userConfigMap.get(id);

    const finalMonatsziel = (savedZiel !== null && savedZiel !== undefined && !isNaN(savedZiel)) 
      ? savedZiel 
      : Math.round(avg12M);

    rows.push([
      id,
      repName,
      avg12M,
      finalMonatsziel
    ]);

    targets[id] = {
      avg12M: avg12M,
      monatsziel: finalMonatsziel,
      jahresziel: finalMonatsziel * 12
    };
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 3, rows.length, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 4, rows.length, 1).setBackground(CONFIG.COLOR_TARGET_EDITABLE).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }

  return targets;
}

function syncAndGetProductTargetsConfig(passedArticles, stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_PRODUKTZIELE);

  // Nimmt übergebene Artikel entgegen oder zieht sie per Fallback aus dem Blatt "Artikel"
  const articles = (passedArticles && Array.isArray(passedArticles) && passedArticles.length > 0) 
    ? passedArticles 
    : getArticlesFromSheet();

  const headers = [
    "Vertreter-ID", 
    "Vertreter-Name", 
    "Artikel-Nr.", 
    "Artikel-Bezeichnung", 
    "Durchschnitt 12M (Stk.)", 
    "Ziel Monat (Stk.)"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PRODUKTZIELE);
    sheet.appendRow(headers);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
  }

  const existingData = sheet.getDataRange().getValues();
  const userTargetMap = new Map();

  for (let i = 1; i < existingData.length; i++) {
    const vtrNr = String(existingData[i][0]).trim();
    const artNr = String(existingData[i][2]).trim();
    const valRaw = existingData[i][5];

    if (vtrNr && artNr) {
      const userZiel = (valRaw !== "" && valRaw !== null && valRaw !== undefined) ? parseFloat(valRaw) : null;
      userTargetMap.set(`${vtrNr}_${artNr}`, userZiel);
    }
  }

  const newRows = [];
  const productTargets = {};

  Object.keys(CONFIG.SALES_REPS).forEach(function(vtrNr) {
    const repName = CONFIG.SALES_REPS[vtrNr].name;
    const repArticles = (stats[vtrNr] && stats[vtrNr].articles) ? stats[vtrNr].articles : {};

    if (!productTargets[vtrNr]) productTargets[vtrNr] = {};

    articles.forEach(function(art) {
      const artNr = String(art.fldArtNr).trim();
      const key = `${vtrNr}_${artNr}`;

      const total12MStk = (repArticles[artNr] && repArticles[artNr].stueck12M) ? repArticles[artNr].stueck12M : 0;
      const avg12MStk = total12MStk / 12.0;

      const savedZiel = userTargetMap.get(key);
      let zielMonat = (savedZiel !== null && savedZiel !== undefined && !isNaN(savedZiel)) 
        ? savedZiel 
        : Math.round(avg12MStk);

      newRows.push([
        vtrNr,
        repName,
        artNr,
        String(art.fldKuBez1 || ""),
        avg12MStk,
        zielMonat
      ]);

      productTargets[vtrNr][artNr] = {
        bezeichnung: String(art.fldKuBez1 || ""),
        avg12M: avg12MStk,
        zielMonat: zielMonat,
        zielJahr: zielMonat * 12
      };
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
    sheet.getRange(2, 1, newRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 3, newRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 5, newRows.length, 1).setNumberFormat("#,##0.0");
    sheet.getRange(2, 6, newRows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 6, newRows.length, 1).setBackground(CONFIG.COLOR_TARGET_EDITABLE).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }

  return productTargets;
}

function syncAndGetKeyAccountsConfig(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_KEYACCOUNTS);

  const headers = [
    "Vertreter-ID", 
    "Vertreter-Name", 
    "Kunden-AdrNr", 
    "Kunden-Name", 
    "Umsatz 12M (€)",
    "Durchschnitt 12M (€)", 
    "Ziel Monat (€)",
    "Ist MTD (€)",
    "Zielerreichung Monat (%)"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_KEYACCOUNTS);
    sheet.appendRow(headers);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
  }

  const existingData = sheet.getDataRange().getValues();
  const kaConfigMap = new Map();

  for (let i = 1; i < existingData.length; i++) {
    const vtrNr = String(existingData[i][0]).trim();
    const adrNr = String(existingData[i][2]).trim();
    const kName = String(existingData[i][3] || "").trim();
    const valRaw = existingData[i][6];

    if (vtrNr && adrNr) {
      const userZiel = (valRaw !== "" && valRaw !== null && valRaw !== undefined) ? parseFloat(valRaw) : null;
      kaConfigMap.set(`${vtrNr}_${adrNr}`, {
        kundenName: kName,
        zielMonat: userZiel
      });
    }
  }

  const newRows = [];
  const keyAccountTargets = {};

  Object.keys(CONFIG.SALES_REPS).forEach(function(vtrNr) {
    const repName = CONFIG.SALES_REPS[vtrNr].name;
    const repKAs = (stats[vtrNr] && stats[vtrNr].keyAccounts) ? stats[vtrNr].keyAccounts : {};

    if (!keyAccountTargets[vtrNr]) keyAccountTargets[vtrNr] = {};

    const allAdrNrs = new Set();

    Object.keys(repKAs).forEach(function(adrNr) {
      if ((repKAs[adrNr].umsatz12M || 0) >= CONFIG.KEYACCOUNT_MIN_UMSATZ) {
        allAdrNrs.add(adrNr);
      }
    });

    kaConfigMap.forEach(function(val, key) {
      if (key.startsWith(`${vtrNr}_`)) {
        allAdrNrs.add(key.split("_")[1]);
      }
    });

    allAdrNrs.forEach(function(adrNr) {
      const kaStat = repKAs[adrNr] || { umsatzMTD: 0, umsatzYTD: 0, umsatz12M: 0, kundenName: "" };
      const key = `${vtrNr}_${adrNr}`;
      const savedConfig = kaConfigMap.get(key) || {};

      const total12M = kaStat.umsatz12M || 0;
      const avg12M = total12M / 12.0;

      const genericNames = ["firma", "company", "herr", "frau", "gmbh", "ag", "kunden-name"];
      let kundenName = kaStat.kundenName;
      if (!kundenName || genericNames.includes(kundenName.toLowerCase())) {
        kundenName = savedConfig.kundenName || `Kunde ${adrNr}`;
      }
      
      const savedZiel = savedConfig.zielMonat;
      let zielMonat = (savedZiel !== null && savedZiel !== undefined && !isNaN(savedZiel)) 
        ? savedZiel 
        : Math.round(avg12M);

      const istMTD = kaStat.umsatzMTD || 0;
      const zielerreichungMonat = zielMonat > 0 ? (istMTD / zielMonat) : 0;

      newRows.push([
        vtrNr,
        repName,
        adrNr,
        kundenName,
        total12M,
        avg12M,
        zielMonat,
        istMTD,
        zielerreichungMonat
      ]);

      keyAccountTargets[vtrNr][adrNr] = {
        kundenName: kundenName,
        umsatz12M: total12M,
        avg12M: avg12M,
        zielMonat: zielMonat,
        zielJahr: zielMonat * 12
      };
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
    sheet.getRange(2, 1, newRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 3, newRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 5, newRows.length, 4).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 9, newRows.length, 1).setNumberFormat("0.0%");

    sheet.getRange(2, 7, newRows.length, 1).setBackground(CONFIG.COLOR_TARGET_EDITABLE).setFontWeight("bold");

    sheet.autoResizeColumns(1, headers.length);
  }

  return keyAccountTargets;
}