/**
 * Dashboard.js - Erstellung der Vertreter-Dashboards mit Key-Account-Bereich ZUERST am Anfang
 */

function renderDashboards(stats, targets, productTargets, keyAccountTargets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeProgress = getTimeProgress();

  Object.keys(CONFIG.SALES_REPS).forEach(function(vtrNr) {
    const repInfo = CONFIG.SALES_REPS[vtrNr];
    const repStats = stats[vtrNr] || { umsatzMTD: 0, umsatzYTD: 0, umsatz12M: 0, prognose: 0, articles: {}, keyAccounts: {} };
    const repTarget = targets[vtrNr] || { avg12M: 0, monatsziel: 0, jahresziel: 0 };
    const repProdTargets = productTargets[vtrNr] || {};
    const repKATargets = keyAccountTargets[vtrNr] || {};

    let sheet = ss.getSheetByName(repInfo.tab);
    if (!sheet) {
      sheet = ss.insertSheet(repInfo.tab);
    } else {
      sheet.clearContents();
    }

    const isEnglish = repInfo.lang === "EN";
    sheet.getRange("A1").setValue(isEnglish ? `Sales Dashboard - ${repInfo.name}` : `Vertriebs-Dashboard - ${repInfo.name}`)
         .setFontSize(14).setFontWeight("bold");

    // 1. GESAMTUMSATZ-KOPFZEILE
    const headers = isEnglish ? 
      ["Metric", "Monthly Target", "Soll Today (Month)", "Actual MTD", "Achievement Month Today (%)", "Yearly Target", "Soll Today (Year)", "Actual YTD", "Achievement Year Today (%)", "Pipeline"] :
      ["Metrik", "Ziel Monat (€)", "Soll Monat heute (€)", "Ist MTD (€)", "Status Monat heute (%)", "Ziel Jahr (€)", "Soll Jahr heute (€)", "Ist YTD (€)", "Status Jahr heute (%)", "Prognose Offen (€)"];

    sheet.getRange("A3:J3").setValues([headers]).setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

    const sollMonatHeute = repTarget.monatsziel * timeProgress.monthProgress;
    const sollJahrHeute = repTarget.jahresziel * timeProgress.yearProgress;

    const pctMonatHeute = sollMonatHeute > 0 ? (repStats.umsatzMTD / sollMonatHeute) : 0;
    const pctJahrHeute = sollJahrHeute > 0 ? (repStats.umsatzYTD / sollJahrHeute) : 0;

    const mainRows = [
      [
        isEnglish ? "Revenue Catalog 1" : "Umsatz Katalog 1 Gesamt",
        repTarget.monatsziel,
        sollMonatHeute,
        repStats.umsatzMTD,
        pctMonatHeute,
        repTarget.jahresziel,
        sollJahrHeute,
        repStats.umsatzYTD,
        pctJahrHeute,
        repStats.prognose
      ]
    ];

    sheet.getRange("A4:J4").setValues(mainRows);
    sheet.getRange("B4:D4").setNumberFormat("#,##0.00 €");
    sheet.getRange("E4").setNumberFormat("0.0%").setBackground(getAchievementColor(pctMonatHeute)).setFontWeight("bold");
    sheet.getRange("F4:H4").setNumberFormat("#,##0.00 €");
    sheet.getRange("I4").setNumberFormat("0.0%").setBackground(getAchievementColor(pctJahrHeute)).setFontWeight("bold");
    sheet.getRange("J4").setNumberFormat("#,##0.00 €");

    let nextStartRow = 7;

    // 2. KEY ACCOUNTS DIREKT AM ANFANG (DIREKT UNTER DEM GESAMTUMSATZ)
    const kaKeys = Object.keys(repKATargets);
    if (kaKeys.length > 0) {
      sheet.getRange(nextStartRow, 1).setValue(isEnglish ? "Key Accounts Overview" : "Key Account Auswertung").setFontSize(12).setFontWeight("bold");
      nextStartRow++;

      const kaHeader = isEnglish ?
        ["Account-Nr.", "Customer Name (Name 1 / Name 2)", "Monthly Target", "Soll Today (Month)", "Actual MTD", "Status Month Today (%)", "Yearly Target", "Soll Today (Year)", "Actual YTD", "Status Year Today (%)", "Pipeline"] :
        ["Kunden-AdrNr", "Kunden-Name (Name 1 / Name 2)", "Ziel Monat (€)", "Soll Monat heute (€)", "Ist MTD (€)", "Status Monat heute (%)", "Ziel Jahr (€)", "Soll Jahr heute (€)", "Ist YTD (€)", "Status Jahr heute (%)", "Prognose Offen (€)"];

      sheet.getRange(nextStartRow, 1, 1, 11).setValues([kaHeader]).setFontWeight("bold").setBackground("#38761d").setFontColor("#ffffff");
      nextStartRow++;

      const kaRows = [];
      const kaColorsMonat = [];
      const kaColorsJahr = [];

      let sumZielMonat = 0, sumSollMonatHeute = 0, sumIstMTD = 0;
      let sumZielJahr = 0, sumSollJahrHeute = 0, sumIstYTD = 0, sumPrognose = 0;

      kaKeys.forEach(function(adrNr) {
        const kaConfig = repKATargets[adrNr];
        const kaStat = repStats.keyAccounts[adrNr] || { umsatzMTD: 0, umsatzYTD: 0, prognose: 0 };

        const zielMonat = kaConfig.zielMonat;
        const zielJahr = kaConfig.zielJahr;

        const kSollMonatHeute = zielMonat * timeProgress.monthProgress;
        const kSollJahrHeute = zielJahr * timeProgress.yearProgress;

        const kPctMonatHeute = kSollMonatHeute > 0 ? (kaStat.umsatzMTD / kSollMonatHeute) : 0;
        const kPctJahrHeute = kSollJahrHeute > 0 ? (kaStat.umsatzYTD / kSollJahrHeute) : 0;

        sumZielMonat += zielMonat;
        sumSollMonatHeute += kSollMonatHeute;
        sumIstMTD += kaStat.umsatzMTD;
        sumZielJahr += zielJahr;
        sumSollJahrHeute += kSollJahrHeute;
        sumIstYTD += kaStat.umsatzYTD;
        sumPrognose += kaStat.prognose;

        kaRows.push([
          adrNr,
          kaConfig.kundenName,
          zielMonat,
          kSollMonatHeute,
          kaStat.umsatzMTD,
          kPctMonatHeute,
          zielJahr,
          kSollJahrHeute,
          kaStat.umsatzYTD,
          kPctJahrHeute,
          kaStat.prognose
        ]);

        kaColorsMonat.push([getAchievementColor(kPctMonatHeute)]);
        kaColorsJahr.push([getAchievementColor(kPctJahrHeute)]);
      });

      const sumPctMonatHeute = sumSollMonatHeute > 0 ? (sumIstMTD / sumSollMonatHeute) : 0;
      const sumPctJahrHeute = sumSollJahrHeute > 0 ? (sumIstYTD / sumSollJahrHeute) : 0;

      const sumRow = [
        "SUMME",
        isEnglish ? "Total Key Accounts" : "Gesamtsumme Key Accounts",
        sumZielMonat,
        sumSollMonatHeute,
        sumIstMTD,
        sumPctMonatHeute,
        sumZielJahr,
        sumSollJahrHeute,
        sumIstYTD,
        sumPctJahrHeute,
        sumPrognose
      ];

      const rowCount = kaRows.length;

      sheet.getRange(nextStartRow, 1, rowCount, 11).setValues(kaRows);
      sheet.getRange(nextStartRow, 1, rowCount, 1).setNumberFormat("@");
      sheet.getRange(nextStartRow, 3, rowCount, 3).setNumberFormat("#,##0.00 €");
      sheet.getRange(nextStartRow, 6, rowCount, 1).setNumberFormat("0.0%").setBackgrounds(kaColorsMonat).setFontWeight("bold");
      sheet.getRange(nextStartRow, 7, rowCount, 3).setNumberFormat("#,##0.00 €");
      sheet.getRange(nextStartRow, 10, rowCount, 1).setNumberFormat("0.0%").setBackgrounds(kaColorsJahr).setFontWeight("bold");
      sheet.getRange(nextStartRow, 11, rowCount, 1).setNumberFormat("#,##0.00 €");

      nextStartRow += rowCount;

      sheet.getRange(nextStartRow, 1, 1, 11).setValues([sumRow]).setFontWeight("bold").setBackground("#d9ead3");
      sheet.getRange(nextStartRow, 3, 1, 3).setNumberFormat("#,##0.00 €");
      sheet.getRange(nextStartRow, 6).setNumberFormat("0.0%").setBackground(getAchievementColor(sumPctMonatHeute));
      sheet.getRange(nextStartRow, 7, 1, 3).setNumberFormat("#,##0.00 €");
      sheet.getRange(nextStartRow, 10).setNumberFormat("0.0%").setBackground(getAchievementColor(sumPctJahrHeute));
      sheet.getRange(nextStartRow, 11).setNumberFormat("#,##0.00 €");

      nextStartRow += 3;
    }

    // 3. PRODUKTZIELE (NACH DEN KEY ACCOUNTS)
    sheet.getRange(nextStartRow, 1).setValue(isEnglish ? "Product Targets (Catalog 1)" : "Produktziel Auswertung (Katalog 1)").setFontSize(12).setFontWeight("bold");
    nextStartRow++;

    const prodHeader = isEnglish ?
      ["SKU", "Description", "Monthly Target", "Soll Today (Month)", "Actual MTD", "Status Month Today (%)", "Yearly Target", "Soll Today (Year)", "Actual YTD", "Status Year Today (%)", "Pipeline"] :
      ["Artikel-Nr.", "Bezeichnung", "Ziel Monat (Stk.)", "Soll Monat heute (Stk.)", "Ist MTD (Stk.)", "Status Monat heute (%)", "Ziel Jahr (Stk.)", "Soll Jahr heute (Stk.)", "Ist YTD (Stk.)", "Status Jahr heute (%)", "Prognose Offen (Stk.)"];

    sheet.getRange(nextStartRow, 1, 1, 11).setValues([prodHeader]).setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
    nextStartRow++;

    const prodRows = [];
    const colorsMonat = [];
    const colorsJahr = [];

    Object.keys(repProdTargets).forEach(function(artNr) {
      const pConfig = repProdTargets[artNr];
      const pStat = repStats.articles[artNr] || { stueckMTD: 0, stueckYTD: 0, stueckPrognose: 0 };

      const zielMonat = pConfig.zielMonat;
      const zielJahr = pConfig.zielJahr;

      const pSollMonatHeute = zielMonat * timeProgress.monthProgress;
      const pSollJahrHeute = zielJahr * timeProgress.yearProgress;

      const pPctMonatHeute = pSollMonatHeute > 0 ? (pStat.stueckMTD / pSollMonatHeute) : 0;
      const pPctJahrHeute = pSollJahrHeute > 0 ? (pStat.stueckYTD / pSollJahrHeute) : 0;

      prodRows.push([
        artNr,
        pConfig.bezeichnung,
        zielMonat,
        pSollMonatHeute,
        pStat.stueckMTD,
        pPctMonatHeute,
        zielJahr,
        pSollJahrHeute,
        pStat.stueckYTD,
        pPctJahrHeute,
        pStat.stueckPrognose
      ]);

      colorsMonat.push([getAchievementColor(pPctMonatHeute)]);
      colorsJahr.push([getAchievementColor(pPctJahrHeute)]);
    });

    if (prodRows.length > 0) {
      const rowCount = prodRows.length;

      sheet.getRange(nextStartRow, 1, rowCount, 11).setValues(prodRows);
      sheet.getRange(nextStartRow, 1, rowCount, 1).setNumberFormat("@");
      
      sheet.getRange(nextStartRow, 3, rowCount, 1).setNumberFormat("#,##0");
      sheet.getRange(nextStartRow, 4, rowCount, 2).setNumberFormat("#,##0.0");
      sheet.getRange(nextStartRow, 6, rowCount, 1).setNumberFormat("0.0%").setBackgrounds(colorsMonat).setFontWeight("bold");

      sheet.getRange(nextStartRow, 7, rowCount, 1).setNumberFormat("#,##0");
      sheet.getRange(nextStartRow, 8, rowCount, 2).setNumberFormat("#,##0.0");
      sheet.getRange(nextStartRow, 10, rowCount, 1).setNumberFormat("0.0%").setBackgrounds(colorsJahr).setFontWeight("bold");
      sheet.getRange(nextStartRow, 11, rowCount, 1).setNumberFormat("#,##0");
    }

    sheet.autoResizeColumns(1, 11);
  });
}

function getTimeProgress() {
  const now = new Date();
  
  const dayOfMonth = now.getDate();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const monthProgress = dayOfMonth / daysInMonth;

  const startOfYear = new Date(now.getFullYear(), 0, 1);
  const diffMs = now - startOfYear;
  const dayOfYear = Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1;
  const isLeapYear = ((now.getFullYear() % 4 === 0 && now.getFullYear() % 100 !== 0) || now.getFullYear() % 400 === 0);
  const daysInYear = isLeapYear ? 366 : 365;
  const yearProgress = dayOfYear / daysInYear;

  return { monthProgress: monthProgress, yearProgress: yearProgress };
}

function getAchievementColor(pct) {
  if (pct >= 1.0) return CONFIG.COLOR_OVER_100;
  if (pct >= 0.90) return CONFIG.COLOR_90_TO_99;
  if (pct >= 0.50) return CONFIG.COLOR_50_TO_89;
  return CONFIG.COLOR_UNDER_50;
}