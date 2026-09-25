// ==========================================
// 4. ANALYZER.GS - Produkt-, Vertreter- & Ländervergleiche
// ==========================================

function analyzeAdvancedSalesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet || rawSheet.getLastRow() < 2) return;

  const qBounds = getQuarterBoundaries();
  const qLabels = qBounds.quarters.map(q => q.label);
  const last12M_QSet = new Set(qLabels.slice(4));

  const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, CONFIG.HEADERS_RAW.length).getValues();

  const productGlobals = {};
  const repCustomerQuarterMatrix = {};
  const repCustomersMap = {}; // vtrNr -> adrNr -> { kundenName, land, isReseller }
  const allCountriesSet = new Set();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const vtrNr = String(row[1]).trim();
    const repName = String(row[2]).trim();
    const adrNr = String(row[3]).trim();
    const kundenName = row[4];
    const land = String(row[5]).trim() || "Unbekannt";
    const artNr = String(row[8]).trim();
    const artName = String(row[9]).trim();
    const menge = Number(row[11]) || 0;
    const umsatz = Number(row[12]) || 0;
    const qLabel = String(row[13]).trim();
    const isReseller = row[14] === true;

    if (!artNr) continue;
    allCountriesSet.add(land);

    if (!productGlobals[artNr]) {
      productGlobals[artNr] = {
        artNr: artNr,
        artName: artName,
        totalMge24M: 0,
        totalUmsatz24M: 0,
        totalMge12M: 0,
        totalUmsatz12M: 0,
        quarters: {},
        repMge24M: {},
        repMge12M: {},
        countryMge24M: {},
        countryMge12M: {}
      };
      qLabels.forEach(lbl => productGlobals[artNr].quarters[lbl] = 0);
    }

    const p = productGlobals[artNr];
    p.totalMge24M += menge;
    p.totalUmsatz24M += umsatz;
    if (p.quarters[qLabel] !== undefined) {
      p.quarters[qLabel] += umsatz;
    }

    p.repMge24M[vtrNr] = (p.repMge24M[vtrNr] || 0) + menge;
    p.countryMge24M[land] = (p.countryMge24M[land] || 0) + menge;

    if (last12M_QSet.has(qLabel)) {
      p.totalMge12M += menge;
      p.totalUmsatz12M += umsatz;
      p.repMge12M[vtrNr] = (p.repMge12M[vtrNr] || 0) + menge;
      p.countryMge12M[land] = (p.countryMge12M[land] || 0) + menge;
    }

    if (vtrNr && CONFIG.REPS[vtrNr]) {
      if (!repCustomersMap[vtrNr]) repCustomersMap[vtrNr] = {};
      if (!repCustomersMap[vtrNr][adrNr]) {
        repCustomersMap[vtrNr][adrNr] = { kundenName: kundenName, land: land, isReseller: isReseller };
      } else if (isReseller) {
        repCustomersMap[vtrNr][adrNr].isReseller = true;
      }

      if (!repCustomerQuarterMatrix[vtrNr]) repCustomerQuarterMatrix[vtrNr] = {};
      const itemKey = adrNr + "___" + artNr;

      if (!repCustomerQuarterMatrix[vtrNr][itemKey]) {
        repCustomerQuarterMatrix[vtrNr][itemKey] = {
          adrNr: adrNr, kundenName: kundenName, land: land,
          artNr: artNr, artName: artName, totalMge: 0,
          quarters: {}
        };
        qLabels.forEach(lbl => repCustomerQuarterMatrix[vtrNr][itemKey].quarters[lbl] = 0);
      }
      repCustomerQuarterMatrix[vtrNr][itemKey].totalMge += menge;
      if (repCustomerQuarterMatrix[vtrNr][itemKey].quarters[qLabel] !== undefined) {
        repCustomerQuarterMatrix[vtrNr][itemKey].quarters[qLabel] += umsatz;
      }
    }
  }

  // Erzeuge "Nie gekauft"-Einträge NUR FÜR WIEDERVERKÄUFER
  for (const vtrNr in repCustomersMap) {
    if (!repCustomerQuarterMatrix[vtrNr]) repCustomerQuarterMatrix[vtrNr] = {};
    const customers = repCustomersMap[vtrNr];

    for (const adrNr in customers) {
      const cust = customers[adrNr];

      if (!cust.isReseller) continue; // Nur Wiederverkäufer

      for (const artNr in productGlobals) {
        const prod = productGlobals[artNr];
        const itemKey = adrNr + "___" + artNr;

        if (!repCustomerQuarterMatrix[vtrNr][itemKey]) {
          repCustomerQuarterMatrix[vtrNr][itemKey] = {
            adrNr: adrNr,
            kundenName: cust.kundenName,
            land: cust.land,
            artNr: artNr,
            artName: prod.artName,
            totalMge: 0,
            quarters: {}
          };
          qLabels.forEach(lbl => repCustomerQuarterMatrix[vtrNr][itemKey].quarters[lbl] = 0);
        }
      }
    }
  }

  buildProductAnalysisSheet(ss, productGlobals, qLabels, qBounds.periodString24M);
  buildRepComparisonSheet(ss, productGlobals, qBounds.periodString12M);
  buildCountryComparisonSheet(ss, productGlobals, Array.from(allCountriesSet), qBounds.periodString12M);
  buildRepQuarterSheets(ss, repCustomerQuarterMatrix, qLabels, qBounds.periodString24M);
}

/**
 * 1. Gesamtübersicht: Produktanalyse mit Vorjahresquartals-Farbvergleich (Q5 vs Q1, Q6 vs Q2, etc.)
 */
function buildProductAnalysisSheet(ss, productGlobals, qLabels, periodString) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCT_ANALYSIS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.PRODUCT_ANALYSIS);
  else { sheet.clear(); if (sheet.getFilter()) sheet.getFilter().remove(); }

  sheet.setFrozenRows(1); // FIXIERUNG HEADER
  sheet.getRange("A:A").setNumberFormat("@");

  const labelPrev12M = `Umsatz ${qLabels[0]} - ${qLabels[3]} (€)`;
  const labelLast12M = `Umsatz ${qLabels[4]} - ${qLabels[7]} (€)`;

  const headers = [
    "Artikel-Nr.", "Artikelname", "Betrachteter Zeitraum", "Gesamtmenge (Stk)", "Gesamtumsatz (€)",
    qLabels[0], qLabels[1], qLabels[2], qLabels[3],
    qLabels[4], qLabels[5], qLabels[6], qLabels[7],
    labelPrev12M, labelLast12M, "Veränderung (12M vs. 12M) (%)",
    "Top-Vertreter (% Anteil)", "Top-Land (% Anteil)", "Gesamttendenz"
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");

  const recordsToProcess = [];

  for (const artNr in productGlobals) {
    const p = productGlobals[artNr];
    const q = p.quarters;

    const prevYear = q[qLabels[0]] + q[qLabels[1]] + q[qLabels[2]] + q[qLabels[3]];
    const currYear = q[qLabels[4]] + q[qLabels[5]] + q[qLabels[6]] + q[qLabels[7]];

    let yoyPercent = 0;
    if (prevYear > 0) yoyPercent = ((currYear - prevYear) / prevYear);
    else if (currYear > 0) yoyPercent = 1.0;

    let topRepKey = "Keine";
    let topRepMge = 0;
    for (const r in p.repMge24M) {
      if (p.repMge24M[r] > topRepMge) {
        topRepMge = p.repMge24M[r];
        topRepKey = r;
      }
    }
    const topRepName = CONFIG.REPS[topRepKey]?.name || topRepKey;
    const topRepShare = p.totalMge24M > 0 ? (topRepMge / p.totalMge24M) : 0;
    const topRepStr = topRepKey !== "Keine" ? `${topRepName} (${(topRepShare * 100).toFixed(1)}%)` : "-";

    let topCountry = "Keine";
    let topCountryMge = 0;
    for (const c in p.countryMge24M) {
      if (p.countryMge24M[c] > topCountryMge) {
        topCountryMge = p.countryMge24M[c];
        topCountry = c;
      }
    }
    const topCountryShare = p.totalMge24M > 0 ? (topCountryMge / p.totalMge24M) : 0;
    const topCountryStr = topCountry !== "Keine" ? `${topCountry} (${(topCountryShare * 100).toFixed(1)}%)` : "-";

    let tendenz = "Stabil";
    if (currYear === 0 && prevYear > 0) tendenz = "Inaktiv / Gestoppt";
    else if (yoyPercent > 0.20) tendenz = "Starkes Wachstum";
    else if (yoyPercent > 0.02) tendenz = "Wachstum";
    else if (yoyPercent < -0.20) tendenz = "Starker Abfall";
    else if (yoyPercent < -0.02) tendenz = "Rückgang";

    const rowVal = [
      String(p.artNr), p.artName, periodString, p.totalMge24M, Math.round(p.totalUmsatz24M * 100) / 100,
      Math.round(q[qLabels[0]] * 100) / 100, Math.round(q[qLabels[1]] * 100) / 100,
      Math.round(q[qLabels[2]] * 100) / 100, Math.round(q[qLabels[3]] * 100) / 100,
      Math.round(q[qLabels[4]] * 100) / 100, Math.round(q[qLabels[5]] * 100) / 100,
      Math.round(q[qLabels[6]] * 100) / 100, Math.round(q[qLabels[7]] * 100) / 100,
      Math.round(prevYear * 100) / 100, Math.round(currYear * 100) / 100,
      yoyPercent, topRepStr, topCountryStr, tendenz
    ];

    const rowColor = getRowColorByStatus(tendenz);
    const rowBg = new Array(headers.length).fill(rowColor);

    // QUARTALS-VERGLEICH: Q1..Q4 ohne Farbe (Basis). Q5..Q8 vs Vorjahresquartal
    for (let i = 0; i < 4; i++) {
      rowBg[5 + i] = null;
      
      const currVal = q[qLabels[i + 4]] || 0;
      const prevVal = q[qLabels[i]] || 0;
      const qColIdx = 9 + i;

      if (currVal > prevVal) {
        rowBg[qColIdx] = CONFIG.COLOR_PALETTE.BEST;
      } else if (currVal < prevVal) {
        rowBg[qColIdx] = CONFIG.COLOR_PALETTE.WORST;
      } else {
        rowBg[qColIdx] = null;
      }
    }

    recordsToProcess.push({
      vals: rowVal,
      bg: rowBg,
      sortKey: p.totalUmsatz24M
    });
  }

  if (recordsToProcess.length > 0) {
    recordsToProcess.sort((a, b) => b.sortKey - a.sortKey);

    const rows = recordsToProcess.map(r => r.vals);
    const bgMatrix = recordsToProcess.map(r => r.bg);

    const range = sheet.getRange(2, 1, rows.length, headers.length);
    range.setValues(rows);
    range.setBackgrounds(bgMatrix);

    sheet.getRange(2, 4, rows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 5, rows.length, 11).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 16, rows.length, 1).setNumberFormat("+0.0%;-0.0%;0.0%");
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 2. Artikel vs. Vertreter (% Anteil + Stückzahl) – NUR LETZTE 12 MONATE
 */
function buildRepComparisonSheet(ss, productGlobals, periodString12M) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.REP_COMPARISON);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.REP_COMPARISON);
  else { sheet.clear(); if (sheet.getFilter()) sheet.getFilter().remove(); }

  sheet.setFrozenRows(1); // FIXIERUNG HEADER
  sheet.getRange("A:A").setNumberFormat("@");

  const repKeys = Object.keys(CONFIG.REPS);
  const repNames = repKeys.map(k => CONFIG.REPS[k].name);

  const headers = ["Artikel-Nr.", "Artikelname", "Betrachteter Zeitraum", "Gesamtmenge 12M (Stk)", "Gesamtumsatz 12M (€)"];
  repNames.forEach(name => {
    headers.push(`${name} (Stk)`);
    headers.push(`${name} (%)`);
  });
  headers.push("Sonstige (%)");

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#16a085").setFontColor("#ffffff").setFontWeight("bold");

  const recordsToProcess = [];

  for (const artNr in productGlobals) {
    const p = productGlobals[artNr];
    if (p.totalMge12M <= 0) continue;

    let assignedMge = 0;
    const repMges = repKeys.map(vtrNr => {
      const mge = p.repMge12M[vtrNr] || 0;
      assignedMge += mge;
      return mge;
    });

    const repShares = repMges.map(mge => mge / p.totalMge12M);
    const otherMge = Math.max(0, p.totalMge12M - assignedMge);
    const otherShare = otherMge / p.totalMge12M;

    const rowVal = [
      String(p.artNr), p.artName, periodString12M, p.totalMge12M, Math.round(p.totalUmsatz12M * 100) / 100
    ];

    for (let i = 0; i < repKeys.length; i++) {
      rowVal.push(repMges[i]);
      rowVal.push(repShares[i]);
    }
    rowVal.push(otherShare);

    const rowBg = new Array(headers.length).fill(null);

    const uniquePos = Array.from(new Set(repShares.filter(s => s > 0))).sort((a, b) => b - a);
    const rank1Val = uniquePos.length > 0 ? uniquePos[0] : null;
    const rank2Val = uniquePos.length > 1 ? uniquePos[1] : null;
    const worstVal = repShares.length > 0 ? Math.min(...repShares) : null;

    repShares.forEach((share, idx) => {
      const stkColIdx = 5 + (idx * 2);
      const pctColIdx = 5 + (idx * 2) + 1;

      if (share === rank1Val && rank1Val !== null) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.BEST;        // 1. Platz (Grün)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.BEST;
      } else if (share === rank2Val && rank2Val !== null) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.SECOND_BEST; // 2. Platz (Gelb)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.SECOND_BEST;
      } else if (share === worstVal && worstVal !== null && (rank1Val === null || worstVal < rank1Val)) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.WORST;       // Schlechtester / 0 (Rot)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.WORST;
      }
    });

    recordsToProcess.push({
      vals: rowVal,
      bg: rowBg,
      sortKey: p.totalUmsatz12M
    });
  }

  if (recordsToProcess.length > 0) {
    recordsToProcess.sort((a, b) => b.sortKey - a.sortKey);

    const rows = recordsToProcess.map(r => r.vals);
    const bgMatrix = recordsToProcess.map(r => r.bg);

    const range = sheet.getRange(2, 1, rows.length, headers.length);
    range.setValues(rows);
    range.setBackgrounds(bgMatrix);

    sheet.getRange(2, 4, rows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("#,##0.00 €");

    for (let i = 0; i < repKeys.length; i++) {
      const stkCol = 6 + (i * 2);
      const pctCol = 6 + (i * 2) + 1;
      sheet.getRange(2, stkCol, rows.length, 1).setNumberFormat("#,##0");
      sheet.getRange(2, pctCol, rows.length, 1).setNumberFormat("0.0%");
    }
    sheet.getRange(2, headers.length, rows.length, 1).setNumberFormat("0.0%");
    
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 3. Artikel vs. Länder (% Anteil + Stückzahl) – NUR LETZTE 12 MONATE
 */
function buildCountryComparisonSheet(ss, productGlobals, allCountries, periodString12M) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.COUNTRY_COMPARISON);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.COUNTRY_COMPARISON);
  else { sheet.clear(); if (sheet.getFilter()) sheet.getFilter().remove(); }

  sheet.setFrozenRows(1); // FIXIERUNG HEADER
  sheet.getRange("A:A").setNumberFormat("@");

  const countryTotals12M = {};
  allCountries.forEach(c => countryTotals12M[c] = 0);

  for (const artNr in productGlobals) {
    const p = productGlobals[artNr];
    for (const c in p.countryMge12M) {
      countryTotals12M[c] = (countryTotals12M[c] || 0) + p.countryMge12M[c];
    }
  }

  const sortedCountries = Object.keys(countryTotals12M).sort((a, b) => countryTotals12M[b] - countryTotals12M[a]);
  const mainCountries = sortedCountries.slice(0, 10);

  const headers = ["Artikel-Nr.", "Artikelname", "Betrachteter Zeitraum", "Gesamtmenge 12M (Stk)", "Gesamtumsatz 12M (€)"];
  mainCountries.forEach(land => {
    headers.push(`${land} (Stk)`);
    headers.push(`${land} (%)`);
  });
  headers.push("Sonstige Länder (%)");

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#2980b9").setFontColor("#ffffff").setFontWeight("bold");

  const recordsToProcess = [];

  for (const artNr in productGlobals) {
    const p = productGlobals[artNr];
    if (p.totalMge12M <= 0) continue;

    let mainMgeSum = 0;
    const countryMges = mainCountries.map(land => {
      const mge = p.countryMge12M[land] || 0;
      mainMgeSum += mge;
      return mge;
    });

    const countryShares = countryMges.map(mge => mge / p.totalMge12M);
    const otherMge = Math.max(0, p.totalMge12M - mainMgeSum);
    const otherShare = otherMge / p.totalMge12M;

    const rowVal = [
      String(p.artNr), p.artName, periodString12M, p.totalMge12M, Math.round(p.totalUmsatz12M * 100) / 100
    ];

    for (let i = 0; i < mainCountries.length; i++) {
      rowVal.push(countryMges[i]);
      rowVal.push(countryShares[i]);
    }
    rowVal.push(otherShare);

    const rowBg = new Array(headers.length).fill(null);

    const uniquePos = Array.from(new Set(countryShares.filter(s => s > 0))).sort((a, b) => b - a);
    const rank1Val = uniquePos.length > 0 ? uniquePos[0] : null;
    const rank2Val = uniquePos.length > 1 ? uniquePos[1] : null;
    const worstVal = countryShares.length > 0 ? Math.min(...countryShares) : null;

    countryShares.forEach((share, idx) => {
      const stkColIdx = 5 + (idx * 2);
      const pctColIdx = 5 + (idx * 2) + 1;

      if (share === rank1Val && rank1Val !== null) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.BEST;        // 1. Platz (Grün)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.BEST;
      } else if (share === rank2Val && rank2Val !== null) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.SECOND_BEST; // 2. Platz (Gelb)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.SECOND_BEST;
      } else if (share === worstVal && worstVal !== null && (rank1Val === null || worstVal < rank1Val)) {
        rowBg[stkColIdx] = CONFIG.COLOR_PALETTE.WORST;       // Schlechtester / 0 (Rot)
        rowBg[pctColIdx] = CONFIG.COLOR_PALETTE.WORST;
      }
    });

    recordsToProcess.push({
      vals: rowVal,
      bg: rowBg,
      sortKey: p.totalUmsatz12M
    });
  }

  if (recordsToProcess.length > 0) {
    recordsToProcess.sort((a, b) => b.sortKey - a.sortKey);

    const rows = recordsToProcess.map(r => r.vals);
    const bgMatrix = recordsToProcess.map(r => r.bg);

    const range = sheet.getRange(2, 1, rows.length, headers.length);
    range.setValues(rows);
    range.setBackgrounds(bgMatrix);

    sheet.getRange(2, 4, rows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("#,##0.00 €");

    for (let i = 0; i < mainCountries.length; i++) {
      const stkCol = 6 + (i * 2);
      const pctCol = 6 + (i * 2) + 1;
      sheet.getRange(2, stkCol, rows.length, 1).setNumberFormat("#,##0");
      sheet.getRange(2, pctCol, rows.length, 1).setNumberFormat("0.0%");
    }
    sheet.getRange(2, headers.length, rows.length, 1).setNumberFormat("0.0%");

    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 4. Individuelle Vertreter-Blätter (DE/EN) inkl. "Nie gekauft" NUR FÜR WIEDERVERKÄUFER
 */
function buildRepQuarterSheets(ss, repMatrix, qLabels, periodString24M) {
  for (const vtrNr in repMatrix) {
    const repConfig = CONFIG.REPS[vtrNr];
    const lang = repConfig.lang;
    let targetSheet = ss.getSheetByName(repConfig.name);
    if (!targetSheet) targetSheet = ss.insertSheet(repConfig.name);
    else { targetSheet.clear(); if (targetSheet.getFilter()) targetSheet.getFilter().remove(); }

    targetSheet.setFrozenRows(1); // FIXIERUNG HEADER
    targetSheet.getRange("A:A").setNumberFormat("@");
    targetSheet.getRange("D:D").setNumberFormat("@");

    const headers = lang === "EN" ? [
      "Customer No.", "Customer Name", "Period", "Article No.", "Article Name", "Total Qty (Pcs)",
      qLabels[0], qLabels[1], qLabels[2], qLabels[3],
      qLabels[4], qLabels[5], qLabels[6], qLabels[7],
      "Diff 12M (€)", "Trend / Status"
    ] : [
      "Kunden-Nr.", "Kundenname", "Betrachteter Zeitraum", "Artikel-Nr.", "Artikelname", "Gesamtmenge (Stk)",
      qLabels[0], qLabels[1], qLabels[2], qLabels[3],
      qLabels[4], qLabels[5], qLabels[6], qLabels[7],
      "Differenz 12M (€)", "Trend / Status"
    ];

    targetSheet.appendRow(headers);
    targetSheet.getRange(1, 1, 1, headers.length).setBackground("#3bb7c4").setFontColor("#ffffff").setFontWeight("bold");

    const recordsToProcess = [];
    const items = repMatrix[vtrNr];

    for (const key in items) {
      const item = items[key];
      const q = item.quarters;
      
      const prevYear = q[qLabels[0]] + q[qLabels[1]] + q[qLabels[2]] + q[qLabels[3]];
      const currYear = q[qLabels[4]] + q[qLabels[5]] + q[qLabels[6]] + q[qLabels[7]];
      const yoyDiff = Math.round((currYear - prevYear) * 100) / 100;

      let status = "";
      if (item.totalMge === 0 && prevYear === 0 && currYear === 0) {
        status = (lang === "EN") ? "Never purchased" : "Nie gekauft";
      } else if (lang === "EN") {
        if (q[qLabels[7]] === 0 && q[qLabels[6]] === 0 && prevYear > 0) status = "Churned / Stopped";
        else if (q[qLabels[7]] < q[qLabels[6]] && q[qLabels[6]] < q[qLabels[5]]) status = "Strong Decline";
        else if (yoyDiff < 0) status = "Decline";
        else if (yoyDiff > 0) status = "Growth";
        else status = "Stable";
      } else {
        if (q[qLabels[7]] === 0 && q[qLabels[6]] === 0 && prevYear > 0) status = "Kauf gestoppt";
        else if (q[qLabels[7]] < q[qLabels[6]] && q[qLabels[6]] < q[qLabels[5]]) status = "Starker Abfall";
        else if (yoyDiff < 0) status = "Rückgang";
        else if (yoyDiff > 0) status = "Wachstum";
        else status = "Stabil";
      }

      const rowVal = [
        String(item.adrNr), item.kundenName, periodString24M, String(item.artNr), item.artName, item.totalMge,
        q[qLabels[0]], q[qLabels[1]], q[qLabels[2]], q[qLabels[3]],
        q[qLabels[4]], q[qLabels[5]], q[qLabels[6]], q[qLabels[7]],
        yoyDiff, status
      ];

      const rowColor = getRowColorByStatus(status);

      recordsToProcess.push({
        vals: rowVal,
        bg: new Array(headers.length).fill(rowColor),
        sortKey: yoyDiff
      });
    }

    if (recordsToProcess.length > 0) {
      recordsToProcess.sort((a, b) => a.sortKey - b.sortKey);

      const rows = recordsToProcess.map(r => r.vals);
      const bgMatrix = recordsToProcess.map(r => r.bg);

      const range = targetSheet.getRange(2, 1, rows.length, headers.length);
      range.setValues(rows);
      range.setBackgrounds(bgMatrix);

      targetSheet.getRange(2, 6, rows.length, 1).setNumberFormat("#,##0");
      targetSheet.getRange(2, 7, rows.length, 9).setNumberFormat("#,##0.00 €");
      targetSheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
      targetSheet.autoResizeColumns(1, headers.length);
    }
  }
}