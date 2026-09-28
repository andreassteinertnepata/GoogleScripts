// =========================================================================
// ZENTRALE KONFIGURATION
// =========================================================================
const CONFIG = {
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  
  // Blattnamen
  SHEET_TOC: "Inhaltsverzeichnis",
  SHEET_REP_SUMMARY: "Umsatz_Vertreter",
  SHEET_MONTHLY_TREND: "Umsatz_nach_Monaten",
  SHEET_CUST_REV_MJ: "Kunden_Umsatz_Monat_Jahr",
  SHEET_CUST_REV_TODAY: "Kunden_Umsatz_Heute",
  SHEET_CUST_REV_YESTERDAY: "Kunden_Umsatz_Gestern",
  SHEET_CUST_ORDERS_TODAY: "Kunden_Bestelleingang_Heute",
  SHEET_CUST_ORDERS_YESTERDAY: "Kunden_Bestelleingang_Gestern",
  SHEET_ART_REV_TODAY: "Artikel_Umsatz_Heute",
  SHEET_ART_REV_YESTERDAY: "Artikel_Umsatz_Gestern",
  SHEET_ART_ORDERS_TODAY: "Artikel_Bestelleingang_Heute",
  SHEET_ART_ORDERS_YESTERDAY: "Artikel_Bestelleingang_Gestern",
  SHEET_OPEN_DOCS: "Offene_Belege",
  SHEET_OPEN_SUMMARY: "Vorgänge_Offene_Summen",
  SHEET_OPEN_DETAILS: "Offene_Bestellungen_Details",

  // Internationales Vertriebsteam (ID -> Name)
  REPS: {
    "28": "Alexis Fonte (DE)",
    "36": "Imad Nassef (DE)",
    "43": "Andreas Steinert (DE)",
    "46": "Stefanie Binder (DE)",
    "56": "Robin Carter Browne (EN)",
    "59": "Aneta Nedyalkova (DE)",
    "60": "Rado Kabakov (EN)"
  },

  // Archiv-Vorgangsarten für Umsatz (Rechnungen & Korrekturen)
  ARCHIVE_TYPES: [
    "70", "105", "109", "110", "113", "115", 
    "122", "123", "129", "154", "155", "156"
  ],

  // Aktive Vorgangsarten für offene Bestellungen (Spalte C = "Vorgänge")
  PROCESS_TYPES: {
    PAID: ["30", "107", "108", "114", "127", "152", "163", "167", "185", "188", "189", "198", "200"],
    UNPAID: ["106", "112", "118", "120", "164"]
  }
};

// =========================================================================
// ONOPEN & SYSTEM-STATUS
// =========================================================================
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastUpdate = PropertiesService.getDocumentProperties().getProperty("LAST_UPDATE") || "Noch kein Update erfolgt";

  safeToast(ss, "Letztes Datahub-Update: " + lastUpdate, "System-Status");

  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Nepata Datahub")
      .addItem("Alle Berichte aktualisieren", "main")
      .addSeparator()
      .addItem("Letztes Update: " + lastUpdate, "showUpdateInfo")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen UI nicht verfügbar: " + e.toString());
  }
}

function showUpdateInfo() {
  const lastUpdate = PropertiesService.getDocumentProperties().getProperty("LAST_UPDATE") || "Bisher kein automatischer Abruf dokumentiert.";
  SpreadsheetApp.getUi().alert(
    "Nepata Datahub Synchronisation",
    "Die Berichte wurden zuletzt am " + lastUpdate + " Uhr erfolgreich aus dem Datahub aktualisiert.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =========================================================================
// HAUPTFUNKTION MIT ZEITSTEMPEL-SPEICHERUNG
// =========================================================================
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  safeToast(ss, "Starte automatischen Datenabruf...", "Datahub Sync");

  try {
    createMonthlyDailyRevenueSheet(ss);
    createRepRevenueSummarySheet(ss);
    createMonthlyRevenueTrendSheet(ss);
    createCustomerRevenueMonthYearSheet(ss);
    createCustomerRevenueTodaySheet(ss);
    createCustomerRevenueYesterdaySheet(ss);
    createCustomerOrderIntakeTodaySheet(ss);
    createCustomerOrderIntakeYesterdaySheet(ss);
    createArticleRevenueTodaySheet(ss);
    createArticleRevenueYesterdaySheet(ss);
    createArticleOrderIntakeTodaySheet(ss);
    createArticleOrderIntakeYesterdaySheet(ss);
    createOpenDocumentsSheet(ss);
    createOpenOrderSummarySheet(ss);
    createOpenOrderDetailsSheet(ss);

    // Navigations-Blatt mit Klick-Links an 1. Position generieren
    createTableOfContentsSheet(ss);

    // Globale Formatierung anwenden
    applyGlobalFormatting(ss);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
    PropertiesService.getDocumentProperties().setProperty("LAST_UPDATE", timestamp);

    onOpen();

    safeToast(ss, "Alle Berichte wurden erfolgreich aktualisiert (" + timestamp + ")!", "Fertig");
  } catch (err) {
    Logger.log("Kritischer Fehler in main(): " + err.toString());
  }
}

// =========================================================================
// INHALTSVERZEICHNIS (INTERNE DIREKTLINKS - DEUTSCHE FORMELSYNTAX)
// =========================================================================
function createTableOfContentsSheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_TOC);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  const categories = {
    "Umsatz_Vertreter": "Vertrieb & Team",
    "Umsatz_nach_Monaten": "Verlauf & Trends",
    "Kunden_Umsatz_Monat_Jahr": "Kundenanalysen",
    "Kunden_Umsatz_Heute": "Kundenanalysen",
    "Kunden_Umsatz_Gestern": "Kundenanalysen",
    "Kunden_Bestelleingang_Heute": "Bestelleingang",
    "Kunden_Bestelleingang_Gestern": "Bestelleingang",
    "Artikel_Umsatz_Heute": "Artikelanalysen",
    "Artikel_Umsatz_Gestern": "Artikelanalysen",
    "Artikel_Bestelleingang_Heute": "Bestelleingang",
    "Artikel_Bestelleingang_Gestern": "Bestelleingang",
    "Offene_Belege": "Offene Vorgänge",
    "Vorgänge_Offene_Summen": "Offene Vorgänge",
    "Offene_Bestellungen_Details": "Offene Vorgänge"
  };

  const headers = ["Kategorie", "Tabellenblatt", "Direktlink", "Anzahl Datensätze"];
  const rows = [headers];

  const sheets = ss.getSheets();
  sheets.forEach(s => {
    const sName = s.getName();
    if (sName === CONFIG.SHEET_TOC) return;

    const sId = s.getSheetId();
    const cat = categories[sName] || (sName.startsWith("Umsatz_") ? "Monatsübersichten" : "Sonstige Berichte");
    
    // Semikolon (;) als Formeltrennzeichen für deutsche Google Sheets
    const linkFormula = `=HYPERLINK("#gid=${sId}"; "➔ Zum Blatt '${sName}' öffnen")`;
    
    const lastRow = s.getLastRow();
    const dataCount = lastRow > 1 ? lastRow - 1 : 0;

    rows.push([cat, sName, linkFormula, dataCount]);
  });

  writeToSheet(sheet, rows);

  sheet.getRange("A:B").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 4, rows.length - 1, 1).setNumberFormat("#,##0");
  }

  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");

  sheet.autoResizeColumns(1, headers.length);
}

// =========================================================================
// HILFSFUNKTION: VERTRETER-AUFLÖSUNG (BELEG OR ADRESSE)
// =========================================================================
function resolveRepInfo(node) {
  const docVtr = String(node.fldVtrNr || "").trim();
  const adrVtr = String(node.rowAdresse?.fldVtrNr || "").trim();
  
  const repId = (docVtr && docVtr !== "0" && docVtr !== "00") ? docVtr : ((adrVtr && adrVtr !== "0" && adrVtr !== "00") ? adrVtr : "");
  const repName = CONFIG.REPS[repId] || "Nicht zugeordnet / Extern";
  
  return { repId: repId, repName: repName };
}

// =========================================================================
// 1. TAGESUMSATZ AKTUELLER MONAT {Umsatz_Monat_Jahr}
// =========================================================================
function createMonthlyDailyRevenueSheet(ss) {
  const now = new Date();
  const monthTitle = "Umsatz_" + Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM_yyyy");
  
  let sheet = prepareSheet(ss, monthTitle);

  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const startISO = startOfMonth.toISOString();

  const query = `
    query GetMonthlyArchive($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  const dailyTotals = {};

  for (let d = 1; d <= endOfMonth.getDate(); d++) {
    const dayStr = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), d), Session.getScriptTimeZone(), "yyyy-MM-dd");
    dailyTotals[dayStr] = { gross: 0, creditNotes: 0, net: 0, grossProfit: 0, count: 0 };
  }

  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startISO, types: formattedTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      if (!node.fldDat) return;

      const docDateStr = Utilities.formatDate(new Date(node.fldDat), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (!dailyTotals[docDateStr]) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const lineRevenue = mge * eprNt;
        const lineProfit = lineRevenue - (mge * eekRoh);

        dailyTotals[docDateStr].net += lineRevenue;
        dailyTotals[docDateStr].grossProfit += lineProfit;

        if (isCreditNote) {
          dailyTotals[docDateStr].creditNotes += lineRevenue;
        } else {
          dailyTotals[docDateStr].gross += lineRevenue;
        }
        dailyTotals[docDateStr].count++;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const daysDE = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
  const headers = ["Datum", "Wochentag", "Bruttoumsatz (€)", "Gutschriften (€)", "Nettoumsatz (€)", "Rohertrag (€)", "Marge (%)", "Positionsanzahl"];
  const rows = [headers];

  Object.keys(dailyTotals).sort().forEach(dateKey => {
    const d = dailyTotals[dateKey];
    const margin = d.net !== 0 ? d.grossProfit / d.net : 0;
    const dateObj = new Date(dateKey + "T00:00:00");
    const dayName = daysDE[dateObj.getDay()];

    rows.push([dateKey, dayName, d.gross, d.creditNotes, d.net, d.grossProfit, margin, d.count]);
  });

  writeToSheet(sheet, rows);
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(2, 3, rows.length - 1, 4).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 8, rows.length - 1, 1).setNumberFormat("#,##0");
  }
}

// =========================================================================
// 2. UMSATZ, ROHERTRAG & MARGE PRO VERTRETER (Umsatz_Vertreter)
// =========================================================================
function createRepRevenueSummarySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_REP_SUMMARY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const currentQuarterMonth = Math.floor(now.getMonth() / 3) * 3;
  const startOfQuarter = new Date(now.getFullYear(), currentQuarterMonth, 1);
  const startOfYear = new Date(now.getFullYear(), 0, 1);
  
  const start30Days = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const start60Days = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
  const start90Days = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

  const minStartDate = new Date(Math.min(startOfYear.getTime(), start90Days.getTime(), startOfYesterday.getTime()));
  const minStartISO = minStartDate.toISOString();

  const stats = {};
  Object.keys(CONFIG.REPS).forEach(repId => {
    stats[repId] = { 
      name: CONFIG.REPS[repId], 
      dayRev: 0, dayProfit: 0, 
      yestRev: 0, yestProfit: 0, 
      monthRev: 0, monthProfit: 0, 
      quarterRev: 0, quarterProfit: 0, 
      d30Rev: 0, d30Profit: 0,
      d60Rev: 0, d60Profit: 0,
      d90Rev: 0, d90Profit: 0,
      yearRev: 0, yearProfit: 0 
    };
  });

  stats["UNASSIGNED"] = { 
    name: "Nicht zugeordnet / Sonstige", 
    dayRev: 0, dayProfit: 0, yestRev: 0, yestProfit: 0, monthRev: 0, monthProfit: 0, 
    quarterRev: 0, quarterProfit: 0, d30Rev: 0, d30Profit: 0,
    d60Rev: 0, d60Profit: 0, d90Rev: 0, d90Profit: 0, yearRev: 0, yearProfit: 0 
  };
  stats["TOTAL"] = { 
    name: "GESAMTUMSATZ", 
    dayRev: 0, dayProfit: 0, yestRev: 0, yestProfit: 0, monthRev: 0, monthProfit: 0, 
    quarterRev: 0, quarterProfit: 0, d30Rev: 0, d30Profit: 0,
    d60Rev: 0, d60Profit: 0, d90Rev: 0, d90Profit: 0, yearRev: 0, yearProfit: 0 
  };

  const query = `
    query GetRepRevenueFull($cursor: String, $minStartISO: DateTime!, $types: [FilterValue!]!) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $minStartISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              fldVtrNr
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { cursor: cursor, minStartISO: minStartISO, types: formattedTypes } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    if (edges.length === 0) break;

    edges.forEach(edge => {
      const node = edge.node || {};
      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const repInfo = resolveRepInfo(node);
      const repKey = stats[repInfo.repId] ? repInfo.repId : "UNASSIGNED";
      const isCreditNote = belegNr.startsWith("123") || artCode === "123";

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        if (docTime >= startOfDay.getTime()) {
          stats[repKey].dayRev += rev;
          stats[repKey].dayProfit += profit;
          stats["TOTAL"].dayRev += rev;
          stats["TOTAL"].dayProfit += profit;
        }

        if (docTime >= startOfYesterday.getTime() && docTime < startOfDay.getTime()) {
          stats[repKey].yestRev += rev;
          stats[repKey].yestProfit += profit;
          stats["TOTAL"].yestRev += rev;
          stats["TOTAL"].yestProfit += profit;
        }

        if (docTime >= startOfMonth.getTime()) {
          stats[repKey].monthRev += rev;
          stats[repKey].monthProfit += profit;
          stats["TOTAL"].monthRev += rev;
          stats["TOTAL"].monthProfit += profit;
        }

        if (docTime >= startOfQuarter.getTime()) {
          stats[repKey].quarterRev += rev;
          stats[repKey].quarterProfit += profit;
          stats["TOTAL"].quarterRev += rev;
          stats["TOTAL"].quarterProfit += profit;
        }

        if (docTime >= start30Days.getTime()) {
          stats[repKey].d30Rev += rev;
          stats[repKey].d30Profit += profit;
          stats["TOTAL"].d30Rev += rev;
          stats["TOTAL"].d30Profit += profit;
        }
        if (docTime >= start60Days.getTime()) {
          stats[repKey].d60Rev += rev;
          stats[repKey].d60Profit += profit;
          stats["TOTAL"].d60Rev += rev;
          stats["TOTAL"].d60Profit += profit;
        }
        if (docTime >= start90Days.getTime()) {
          stats[repKey].d90Rev += rev;
          stats[repKey].d90Profit += profit;
          stats["TOTAL"].d90Rev += rev;
          stats["TOTAL"].d90Profit += profit;
        }

        if (docTime >= startOfYear.getTime()) {
          stats[repKey].yearRev += rev;
          stats[repKey].yearProfit += profit;
          stats["TOTAL"].yearRev += rev;
          stats["TOTAL"].yearProfit += profit;
        }
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Vertreter-ID", "Vertreter Name", 
    "Umsatz Heute (€)", "Rohertrag Heute (€)", "Marge Heute (%)",
    "Umsatz Gestern (€)", "Rohertrag Gestern (€)", "Marge Gestern (%)",
    "Umsatz Monat (€)", "Rohertrag Monat (€)", "Marge Monat (%)",
    "Umsatz Quartal (€)", "Rohertrag Quartal (€)", "Marge Quartal (%)",
    "Umsatz 30 Tage (€)", "Rohertrag 30 Tage (€)", "Marge 30 Tage (%)",
    "Umsatz 60 Tage (€)", "Rohertrag 60 Tage (€)", "Marge 60 Tage (%)",
    "Umsatz 90 Tage (€)", "Rohertrag 90 Tage (€)", "Marge 90 Tage (%)",
    "Umsatz Jahr (€)", "Rohertrag Jahr (€)", "Marge Jahr (%)"
  ];
  const rows = [headers];

  const calcM = (p, r) => (r !== 0 ? p / r : 0);

  Object.keys(CONFIG.REPS).forEach(repId => {
    const s = stats[repId];
    rows.push([
      repId, s.name, 
      s.dayRev, s.dayProfit, calcM(s.dayProfit, s.dayRev),
      s.yestRev, s.yestProfit, calcM(s.yestProfit, s.yestRev),
      s.monthRev, s.monthProfit, calcM(s.monthProfit, s.monthRev),
      s.quarterRev, s.quarterProfit, calcM(s.quarterProfit, s.quarterRev),
      s.d30Rev, s.d30Profit, calcM(s.d30Profit, s.d30Rev),
      s.d60Rev, s.d60Profit, calcM(s.d60Profit, s.d60Rev),
      s.d90Rev, s.d90Profit, calcM(s.d90Profit, s.d90Rev),
      s.yearRev, s.yearProfit, calcM(s.yearProfit, s.yearRev)
    ]);
  });

  const u = stats["UNASSIGNED"];
  rows.push(["", u.name, u.dayRev, u.dayProfit, calcM(u.dayProfit, u.dayRev), u.yestRev, u.yestProfit, calcM(u.yestProfit, u.yestRev), u.monthRev, u.monthProfit, calcM(u.monthProfit, u.monthRev), u.quarterRev, u.quarterProfit, calcM(u.quarterProfit, u.quarterRev), u.d30Rev, u.d30Profit, calcM(u.d30Profit, u.d30Rev), u.d60Rev, u.d60Profit, calcM(u.d60Profit, u.d60Rev), u.d90Rev, u.d90Profit, calcM(u.d90Profit, u.d90Rev), u.yearRev, u.yearProfit, calcM(u.yearProfit, u.yearRev)]);

  const t = stats["TOTAL"];
  rows.push(["", t.name, t.dayRev, t.dayProfit, calcM(t.dayProfit, t.dayRev), t.yestRev, t.yestProfit, calcM(t.yestProfit, t.yestRev), t.monthRev, t.monthProfit, calcM(t.monthProfit, t.monthRev), t.quarterRev, t.quarterProfit, calcM(t.quarterProfit, t.quarterRev), t.d30Rev, t.d30Profit, calcM(t.d30Profit, t.d30Rev), t.d60Rev, t.d60Profit, calcM(t.d60Profit, t.d60Rev), t.d90Rev, t.d90Profit, calcM(t.d90Profit, t.d90Rev), t.yearRev, t.yearProfit, calcM(t.yearProfit, t.yearRev)]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  
  if (rows.length > 1) {
    for (let c = 3; c <= 26; c += 3) {
      sheet.getRange(2, c, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
      sheet.getRange(2, c + 2, rows.length - 1, 1).setNumberFormat("0.00%");
    }
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 3. UMSATZ NACH MONATEN (Umsatz_nach_Monaten)
// =========================================================================
function createMonthlyRevenueTrendSheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_MONTHLY_TREND);

  const currentYear = new Date().getFullYear();
  const startOfYear2024ISO = "2024-01-01T00:00:00Z";

  const query = `
    query GetMonthlyTrend($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  const monthlyData = {};

  for (let yr = 2024; yr <= currentYear; yr++) {
    monthlyData[yr] = {};
    for (let m = 1; m <= 12; m++) {
      monthlyData[yr][m] = { rev: 0, profit: 0 };
    }
  }

  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfYear2024ISO, types: formattedTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      if (!node.fldDat) return;

      const docDate = new Date(node.fldDat);
      const yr = docDate.getFullYear();
      const m = docDate.getMonth() + 1;

      if (!monthlyData[yr] || !monthlyData[yr][m]) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const price = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * price;
        const profit = rev - (mge * eekRoh);

        monthlyData[yr][m].rev += rev;
        monthlyData[yr][m].profit += profit;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = ["Monat Code", "Monatsname"];
  for (let yr = 2024; yr <= currentYear; yr++) {
    headers.push(`Umsatz ${yr} (€)`, `Rohertrag ${yr} (€)`, `Marge ${yr} (%)`);
  }

  const monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
  const rows = [headers];

  for (let m = 1; m <= 12; m++) {
    const monthStr = m < 10 ? "0" + m : "" + m;
    const row = [monthStr, monthNames[m - 1]];

    for (let yr = 2024; yr <= currentYear; yr++) {
      const r = monthlyData[yr][m].rev;
      const p = monthlyData[yr][m].profit;
      const margin = r !== 0 ? p / r : 0;
      row.push(r, p, margin);
    }
    rows.push(row);
  }

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    for (let c = 3; c <= headers.length; c += 3) {
      sheet.getRange(2, c, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
      sheet.getRange(2, c + 2, rows.length - 1, 1).setNumberFormat("0.00%");
    }
  }
}

// =========================================================================
// 4. KUNDEN-UMSATZ: MONAT & JAHR (Kunden_Umsatz_Monat_Jahr)
// =========================================================================
function createCustomerRevenueMonthYearSheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_CUST_REV_MJ);

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const startOfYear = new Date(now.getFullYear(), 0, 1);
  const startOfYearISO = startOfYear.toISOString();

  const customerMap = {};

  const query = `
    query GetCustArchiveMJ($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldReNa1
              fldReNa2
              fldVtrNr
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedArchiveTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfYearISO, types: formattedArchiveTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");
      const repInfo = resolveRepInfo(node);

      if (!customerMap[custNr]) {
        customerMap[custNr] = {
          custNr: custNr,
          custName: custName,
          repId: repInfo.repId,
          repName: repInfo.repName,
          revMonth: 0, profitMonth: 0,
          revYear: 0, profitYear: 0
        };
      } else if (!customerMap[custNr].repId && repInfo.repId) {
        customerMap[custNr].repId = repInfo.repId;
        customerMap[custNr].repName = repInfo.repName;
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        customerMap[custNr].revYear += rev;
        customerMap[custNr].profitYear += profit;

        if (docTime >= startOfMonth.getTime()) {
          customerMap[custNr].revMonth += rev;
          customerMap[custNr].profitMonth += profit;
        }
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Kunden-Nr.", "Kundenname", "Vertreter-ID", "Vertreter Name",
    "Umsatz Dieser Monat (€)", "Rohertrag Monat (€)", "Marge Monat (%)",
    "Umsatz Dieses Jahr (€)", "Rohertrag Jahr (€)", "Marge Jahr (%)"
  ];
  const rows = [headers];

  Object.values(customerMap)
    .sort((a, b) => b.revMonth - a.revMonth)
    .forEach(c => {
      const marginM = c.revMonth !== 0 ? c.profitMonth / c.revMonth : 0;
      const marginY = c.revYear !== 0 ? c.profitYear / c.revYear : 0;
      rows.push([
        c.custNr, c.custName, c.repId, c.repName,
        c.revMonth, c.profitMonth, marginM,
        c.revYear, c.profitYear, marginY
      ]);
    });

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 8, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 10, rows.length - 1, 1).setNumberFormat("0.00%");
  }
}

// =========================================================================
// 5. KUNDEN-UMSATZ: HEUTE (Kunden_Umsatz_Heute)
// =========================================================================
function createCustomerRevenueTodaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_CUST_REV_TODAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfDayISO = startOfDay.toISOString();

  const customerMap = {};
  let sumTodayRev = 0, sumTodayProfit = 0;

  const query = `
    query GetCustArchiveToday($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldReNa1
              fldReNa2
              fldVtrNr
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedArchiveTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfDayISO, types: formattedArchiveTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      if (docTime < startOfDay.getTime()) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");
      const repInfo = resolveRepInfo(node);

      if (!customerMap[custNr]) {
        customerMap[custNr] = {
          custNr: custNr,
          custName: custName,
          repId: repInfo.repId,
          repName: repInfo.repName,
          revToday: 0, profitToday: 0
        };
      } else if (!customerMap[custNr].repId && repInfo.repId) {
        customerMap[custNr].repId = repInfo.repId;
        customerMap[custNr].repName = repInfo.repName;
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        customerMap[custNr].revToday += rev;
        customerMap[custNr].profitToday += profit;

        sumTodayRev += rev;
        sumTodayProfit += profit;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Kunden-Nr.", "Kundenname", "Vertreter-ID", "Vertreter Name",
    "Umsatz Heute (€)", "Rohertrag Heute (€)", "Marge Heute (%)"
  ];
  const rows = [headers];

  Object.values(customerMap)
    .sort((a, b) => b.revToday - a.revToday)
    .forEach(c => {
      const margin = c.revToday !== 0 ? c.profitToday / c.revToday : 0;
      rows.push([
        c.custNr, c.custName, c.repId, c.repName,
        c.revToday, c.profitToday, margin
      ]);
    });

  const totalMarginToday = sumTodayRev !== 0 ? sumTodayProfit / sumTodayRev : 0;
  rows.push(["", "SUMME HEUTE", "", "", sumTodayRev, sumTodayProfit, totalMarginToday]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 6. KUNDEN-UMSATZ: GESTERN (Kunden_Umsatz_Gestern)
// =========================================================================
function createCustomerRevenueYesterdaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_CUST_REV_YESTERDAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const startOfYesterdayISO = startOfYesterday.toISOString();

  const customerMap = {};
  let sumYestRev = 0, sumYestProfit = 0;

  const query = `
    query GetCustArchiveYesterday($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldReNa1
              fldReNa2
              fldVtrNr
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
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

  const formattedArchiveTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfYesterdayISO, types: formattedArchiveTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      if (docTime < startOfYesterday.getTime() || docTime >= startOfDay.getTime()) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");
      const repInfo = resolveRepInfo(node);

      if (!customerMap[custNr]) {
        customerMap[custNr] = {
          custNr: custNr,
          custName: custName,
          repId: repInfo.repId,
          repName: repInfo.repName,
          revYesterday: 0, profitYesterday: 0
        };
      } else if (!customerMap[custNr].repId && repInfo.repId) {
        customerMap[custNr].repId = repInfo.repId;
        customerMap[custNr].repName = repInfo.repName;
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        customerMap[custNr].revYesterday += rev;
        customerMap[custNr].profitYesterday += profit;

        sumYestRev += rev;
        sumYestProfit += profit;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Kunden-Nr.", "Kundenname", "Vertreter-ID", "Vertreter Name",
    "Umsatz Gestern (€)", "Rohertrag Gestern (€)", "Marge Gestern (%)"
  ];
  const rows = [headers];

  Object.values(customerMap)
    .sort((a, b) => b.revYesterday - a.revYesterday)
    .forEach(c => {
      const margin = c.revYesterday !== 0 ? c.profitYesterday / c.revYesterday : 0;
      rows.push([
        c.custNr, c.custName, c.repId, c.repName,
        c.revYesterday, c.profitYesterday, margin
      ]);
    });

  const totalMarginYest = sumYestRev !== 0 ? sumYestProfit / sumYestRev : 0;
  rows.push(["", "SUMME GESTERN", "", "", sumYestRev, sumYestProfit, totalMarginYest]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 7. KUNDEN-BESTELLEINGANG: HEUTE (Kunden_Bestelleingang_Heute)
// =========================================================================
function createCustomerOrderIntakeTodaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_CUST_ORDERS_TODAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedVorgangCodes = allVorgangCodes.map(code => ({ string: code }));

  const customerMap = {};
  let sumTodayRev = 0, sumTodayProfit = 0;

  const query = `
    query GetCustOrdersToday($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldAdrNr
              fldReNa1
              fldReNa2
              fldVtrNr
              fldDat
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
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

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedVorgangCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const docTime = node.fldDat ? new Date(node.fldDat).getTime() : 0;
      if (docTime < startOfDay.getTime()) return;

      const repInfo = resolveRepInfo(node);

      if (!customerMap[custNr]) {
        customerMap[custNr] = {
          custNr: custNr,
          custName: custName,
          repId: repInfo.repId,
          repName: repInfo.repName,
          ordersToday: 0, profitToday: 0
        };
      } else if (!customerMap[custNr].repId && repInfo.repId) {
        customerMap[custNr].repId = repInfo.repId;
        customerMap[custNr].repName = repInfo.repName;
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        const mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        const val = mge * eprNt;
        const profit = val - (mge * eekRoh);

        customerMap[custNr].ordersToday += val;
        customerMap[custNr].profitToday += profit;

        sumTodayRev += val;
        sumTodayProfit += profit;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Kunden-Nr.", "Kundenname", "Vertreter-ID", "Vertreter Name",
    "Bestelleingang Heute (€)", "Rohertrag Bestelleingang Heute (€)", "Marge Heute (%)"
  ];
  const rows = [headers];

  Object.values(customerMap)
    .sort((a, b) => b.ordersToday - a.ordersToday)
    .forEach(c => {
      const margin = c.ordersToday !== 0 ? c.profitToday / c.ordersToday : 0;
      rows.push([
        c.custNr, c.custName, c.repId, c.repName,
        c.ordersToday, c.profitToday, margin
      ]);
    });

  const totalMarginToday = sumTodayRev !== 0 ? sumTodayProfit / sumTodayRev : 0;
  rows.push(["", "SUMME HEUTE", "", "", sumTodayRev, sumTodayProfit, totalMarginToday]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 8. KUNDEN-BESTELLEINGANG: GESTERN (Kunden_Bestelleingang_Gestern)
// =========================================================================
function createCustomerOrderIntakeYesterdaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_CUST_ORDERS_YESTERDAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedVorgangCodes = allVorgangCodes.map(code => ({ string: code }));

  const customerMap = {};
  let sumYestRev = 0, sumYestProfit = 0;

  const query = `
    query GetCustOrdersYesterday($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldAdrNr
              fldReNa1
              fldReNa2
              fldVtrNr
              fldDat
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
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

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedVorgangCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const docTime = node.fldDat ? new Date(node.fldDat).getTime() : 0;
      if (docTime < startOfYesterday.getTime() || docTime >= startOfDay.getTime()) return;

      const repInfo = resolveRepInfo(node);

      if (!customerMap[custNr]) {
        customerMap[custNr] = {
          custNr: custNr,
          custName: custName,
          repId: repInfo.repId,
          repName: repInfo.repName,
          ordersYesterday: 0, profitYesterday: 0
        };
      } else if (!customerMap[custNr].repId && repInfo.repId) {
        customerMap[custNr].repId = repInfo.repId;
        customerMap[custNr].repName = repInfo.repName;
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        const mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        const val = mge * eprNt;
        const profit = val - (mge * eekRoh);

        customerMap[custNr].ordersYesterday += val;
        customerMap[custNr].profitYesterday += profit;

        sumYestRev += val;
        sumYestProfit += profit;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Kunden-Nr.", "Kundenname", "Vertreter-ID", "Vertreter Name",
    "Bestelleingang Gestern (€)", "Rohertrag Bestelleingang Gestern (€)", "Marge Gestern (%)"
  ];
  const rows = [headers];

  Object.values(customerMap)
    .sort((a, b) => b.ordersYesterday - a.ordersYesterday)
    .forEach(c => {
      const margin = c.ordersYesterday !== 0 ? c.profitYesterday / c.ordersYesterday : 0;
      rows.push([
        c.custNr, c.custName, c.repId, c.repName,
        c.ordersYesterday, c.profitYesterday, margin
      ]);
    });

  const totalMarginYest = sumYestRev !== 0 ? sumYestProfit / sumYestRev : 0;
  rows.push(["", "SUMME GESTERN", "", "", sumYestRev, sumYestProfit, totalMarginYest]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat("0.00%");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 9. ARTIKEL-UMSATZ: HEUTE (Artikel_Umsatz_Heute)
// =========================================================================
function createArticleRevenueTodaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_ART_REV_TODAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfDayISO = startOfDay.toISOString();

  const articleMap = {};

  const query = `
    query GetArtArchiveToday($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
                rowArtikel { fldKuBez1 }
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

  const formattedArchiveTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  let sumTodayRev = 0, sumTodayProfit = 0, sumTodayQty = 0, sumTodayCount = 0;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfDayISO, types: formattedArchiveTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      if (docTime < startOfDay.getTime()) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;
        let artNr = String(pos.fldArtNr || "").trim();
        if (!artNr) artNr = "SONSTIGE";

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        if (!articleMap[artNr]) {
          articleMap[artNr] = {
            artNr: artNr,
            artBez: String(pos.rowArtikel?.fldKuBez1 || (artNr === "SONSTIGE" ? "Sonderposition / Sonstige" : artNr)),
            revToday: 0, profitToday: 0, qtyToday: 0, countToday: 0
          };
        }

        const aObj = articleMap[artNr];
        aObj.revToday += rev;
        aObj.profitToday += profit;
        aObj.qtyToday += mge;
        aObj.countToday++;

        sumTodayRev += rev;
        sumTodayProfit += profit;
        sumTodayQty += mge;
        sumTodayCount++;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Artikel-Nr.", "Artikelbezeichnung", 
    "Umsatz Heute (€)", "Rohertrag Heute (€)", "Marge Heute (%)", "Menge Heute", "Positionsanzahl Heute"
  ];
  const rows = [headers];

  Object.values(articleMap)
    .sort((a, b) => b.revToday - a.revToday)
    .forEach(a => {
      const margin = a.revToday !== 0 ? a.profitToday / a.revToday : 0;
      rows.push([
        a.artNr, a.artBez,
        a.revToday, a.profitToday, margin, a.qtyToday, a.countToday
      ]);
    });

  const totalMarginToday = sumTodayRev !== 0 ? sumTodayProfit / sumTodayRev : 0;
  rows.push(["", "SUMME HEUTE", sumTodayRev, sumTodayProfit, totalMarginToday, sumTodayQty, sumTodayCount]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 6, rows.length - 1, 2).setNumberFormat("#,##0");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 10. ARTIKEL-UMSATZ: GESTERN (Artikel_Umsatz_Gestern)
// =========================================================================
function createArticleRevenueYesterdaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_ART_REV_YESTERDAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const startOfYesterdayISO = startOfYesterday.toISOString();

  const articleMap = {};

  const query = `
    query GetArtArchiveYesterday($startISO: DateTime!, $types: [FilterValue!]!, $cursor: String) {
      tblVorgangArchiv {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              { ge: [{ field: fldDat }, { value: { datetime: $startISO } }] },
              { in: { field: fldArt, values: $types } }
            ]
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              fldDat
              fldErstDat
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                fldEEkRoh
                fldAbrPosKz
                rowArtikel { fldKuBez1 }
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

  const formattedArchiveTypes = CONFIG.ARCHIVE_TYPES.map(type => ({ string: type }));
  let hasNextPage = true;
  let cursor = null;

  let sumYestRev = 0, sumYestProfit = 0, sumYestQty = 0, sumYestCount = 0;

  while (hasNextPage) {
    const payload = { query: query, variables: { startISO: startOfYesterdayISO, types: formattedArchiveTypes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgangArchiv?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const docDateStr = node.fldDat || node.fldErstDat;
      if (!docDateStr) return;

      const docTime = new Date(docDateStr).getTime();
      if (docTime < startOfYesterday.getTime() || docTime >= startOfDay.getTime()) return;

      const belegNr = String(node.fldBelegNr || "");
      const artCode = String(node.fldArt || "");
      const isCreditNote = (belegNr.startsWith("123") || artCode === "123");

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        if (pos.fldAbrPosKz !== true) return;
        let artNr = String(pos.fldArtNr || "").trim();
        if (!artNr) artNr = "SONSTIGE";

        let mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        if (isCreditNote) {
          mge = -Math.abs(mge);
        }

        const rev = mge * eprNt;
        const profit = rev - (mge * eekRoh);

        if (!articleMap[artNr]) {
          articleMap[artNr] = {
            artNr: artNr,
            artBez: String(pos.rowArtikel?.fldKuBez1 || (artNr === "SONSTIGE" ? "Sonderposition / Sonstige" : artNr)),
            revYesterday: 0, profitYesterday: 0, qtyYesterday: 0, countYesterday: 0
          };
        }

        const aObj = articleMap[artNr];
        aObj.revYesterday += rev;
        aObj.profitYesterday += profit;
        aObj.qtyYesterday += mge;
        aObj.countYesterday++;

        sumYestRev += rev;
        sumYestProfit += profit;
        sumYestQty += mge;
        sumYestCount++;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Artikel-Nr.", "Artikelbezeichnung", 
    "Umsatz Gestern (€)", "Rohertrag Gestern (€)", "Marge Gestern (%)", "Menge Gestern", "Positionsanzahl Gestern"
  ];
  const rows = [headers];

  Object.values(articleMap)
    .sort((a, b) => b.revYesterday - a.revYesterday)
    .forEach(a => {
      const margin = a.revYesterday !== 0 ? a.profitYesterday / a.revYesterday : 0;
      rows.push([
        a.artNr, a.artBez,
        a.revYesterday, a.profitYesterday, margin, a.qtyYesterday, a.countYesterday
      ]);
    });

  const totalMarginYest = sumYestRev !== 0 ? sumYestProfit / sumYestRev : 0;
  rows.push(["", "SUMME GESTERN", sumYestRev, sumYestProfit, totalMarginYest, sumYestQty, sumYestCount]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 6, rows.length - 1, 2).setNumberFormat("#,##0");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 11. ARTIKEL-BESTELLEINGANG: HEUTE (Artikel_Bestelleingang_Heute)
// =========================================================================
function createArticleOrderIntakeTodaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_ART_ORDERS_TODAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedVorgangCodes = allVorgangCodes.map(code => ({ string: code }));

  const articleMap = {};
  let sumTodayRev = 0, sumTodayProfit = 0, sumTodayQty = 0;

  const query = `
    query GetArtOrdersToday($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldDat
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                fldEEkRoh
                rowArtikel { fldKuBez1 }
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

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedVorgangCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const docTime = node.fldDat ? new Date(node.fldDat).getTime() : 0;
      if (docTime < startOfDay.getTime()) return;

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        let artNr = String(pos.fldArtNr || "").trim();
        if (!artNr) artNr = "SONSTIGE";

        const mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        const val = mge * eprNt;
        const profit = val - (mge * eekRoh);

        if (!articleMap[artNr]) {
          articleMap[artNr] = {
            artNr: artNr,
            artBez: String(pos.rowArtikel?.fldKuBez1 || (artNr === "SONSTIGE" ? "Sonderposition / Sonstige" : artNr)),
            ordersToday: 0, profitToday: 0, qtyToday: 0
          };
        }

        const aObj = articleMap[artNr];
        aObj.ordersToday += val;
        aObj.profitToday += profit;
        aObj.qtyToday += mge;

        sumTodayRev += val;
        sumTodayProfit += profit;
        sumTodayQty += mge;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Artikel-Nr.", "Artikelbezeichnung", 
    "Bestelleingang Heute (€)", "Rohertrag Heute (€)", "Marge Heute (%)", "Menge Heute"
  ];
  const rows = [headers];

  Object.values(articleMap)
    .sort((a, b) => b.ordersToday - a.ordersToday)
    .forEach(a => {
      const margin = a.ordersToday !== 0 ? a.profitToday / a.ordersToday : 0;
      rows.push([
        a.artNr, a.artBez,
        a.ordersToday, a.profitToday, margin, a.qtyToday
      ]);
    });

  const totalMarginToday = sumTodayRev !== 0 ? sumTodayProfit / sumTodayRev : 0;
  rows.push(["", "SUMME HEUTE", sumTodayRev, sumTodayProfit, totalMarginToday, sumTodayQty]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat("#,##0");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 12. ARTIKEL-BESTELLEINGANG: GESTERN (Artikel_Bestelleingang_Gestern)
// =========================================================================
function createArticleOrderIntakeYesterdaySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_ART_ORDERS_YESTERDAY);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedVorgangCodes = allVorgangCodes.map(code => ({ string: code }));

  const articleMap = {};
  let sumYestRev = 0, sumYestProfit = 0, sumYestQty = 0;

  const query = `
    query GetArtOrdersYesterday($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldDat
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                fldEEkRoh
                rowArtikel { fldKuBez1 }
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

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedVorgangCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const docTime = node.fldDat ? new Date(node.fldDat).getTime() : 0;
      if (docTime < startOfYesterday.getTime() || docTime >= startOfDay.getTime()) return;

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        let artNr = String(pos.fldArtNr || "").trim();
        if (!artNr) artNr = "SONSTIGE";

        const mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        const val = mge * eprNt;
        const profit = val - (mge * eekRoh);

        if (!articleMap[artNr]) {
          articleMap[artNr] = {
            artNr: artNr,
            artBez: String(pos.rowArtikel?.fldKuBez1 || (artNr === "SONSTIGE" ? "Sonderposition / Sonstige" : artNr)),
            ordersYesterday: 0, profitYesterday: 0, qtyYesterday: 0
          };
        }

        const aObj = articleMap[artNr];
        aObj.ordersYesterday += val;
        aObj.profitYesterday += profit;
        aObj.qtyYesterday += mge;

        sumYestRev += val;
        sumYestProfit += profit;
        sumYestQty += mge;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Artikel-Nr.", "Artikelbezeichnung", 
    "Bestelleingang Gestern (€)", "Rohertrag Gestern (€)", "Marge Gestern (%)", "Menge Gestern"
  ];
  const rows = [headers];

  Object.values(articleMap)
    .sort((a, b) => b.ordersYesterday - a.ordersYesterday)
    .forEach(a => {
      const margin = a.ordersYesterday !== 0 ? a.profitYesterday / a.ordersYesterday : 0;
      rows.push([
        a.artNr, a.artBez,
        a.ordersYesterday, a.profitYesterday, margin, a.qtyYesterday
      ]);
    });

  const totalMarginYest = sumYestRev !== 0 ? sumYestProfit / sumYestRev : 0;
  rows.push(["", "SUMME GESTERN", sumYestRev, sumYestProfit, totalMarginYest, sumYestQty]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat("#,##0");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 13. OFFENE BELEGE GRUPPIERT NACH VORGANGSART (Offene_Belege - 6 SPALTEN)
// =========================================================================
function createOpenDocumentsSheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_OPEN_DOCS);

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedCodes = allVorgangCodes.map(code => ({ string: code }));

  const query = `
    query GetOpenBelege($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldBelegNr
              fldArt
              lblArt
              rowsPositions {
                fldMge
                fldEPrNt
                fldEEkRoh
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

  const summaryMap = {};
  let sumPaidNet = 0, sumPaidProfit = 0;
  let sumUnpaidNet = 0, sumUnpaidProfit = 0;

  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const artCode = String(node.fldArt || "").trim();
      if (!artCode) return;

      const lblArt = String(node.lblArt || artCode).trim();
      const isPaid = CONFIG.PROCESS_TYPES.PAID.includes(artCode);
      const zahlungStatus = isPaid ? "bezahlt" : "unbezahlt";
      const key = artCode + "_" + zahlungStatus;

      if (!summaryMap[key]) {
        summaryMap[key] = {
          code: artCode,
          bezeichnung: lblArt,
          zahlung: zahlungStatus,
          isPaid: isPaid,
          netTotal: 0,
          profitTotal: 0
        };
      }

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        const mge = pos.fldMge || 0;
        const eprNt = pos.fldEPrNt || 0;
        const eekRoh = pos.fldEEkRoh || 0;

        const val = mge * eprNt;
        const profit = val - (mge * eekRoh);

        summaryMap[key].netTotal += val;
        summaryMap[key].profitTotal += profit;

        if (isPaid) {
          sumPaidNet += val;
          sumPaidProfit += profit;
        } else {
          sumUnpaidNet += val;
          sumUnpaidProfit += profit;
        }
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = [
    "Code", "Bezeichnung (Vorgangsart)", "Zahlung", 
    "Gesamtbetrag Netto (€)", "Rohertrag (€)", "Marge (%)"
  ];
  const rows = [headers];

  Object.values(summaryMap)
    .sort((a, b) => b.netTotal - a.netTotal)
    .forEach(item => {
      const margin = item.netTotal !== 0 ? item.profitTotal / item.netTotal : 0;
      rows.push([
        item.code,
        item.bezeichnung,
        item.zahlung,
        item.netTotal,
        item.profitTotal,
        margin
      ]);
    });

  const paidMargin = sumPaidNet !== 0 ? sumPaidProfit / sumPaidNet : 0;
  const unpaidMargin = sumUnpaidNet !== 0 ? sumUnpaidProfit / sumUnpaidNet : 0;
  const grandTotalNet = sumPaidNet + sumUnpaidNet;
  const grandTotalProfit = sumPaidProfit + sumUnpaidProfit;
  const grandTotalMargin = grandTotalNet !== 0 ? grandTotalProfit / grandTotalNet : 0;

  rows.push(["", "SUMME BEZAHLT", "bezahlt", sumPaidNet, sumPaidProfit, paidMargin]);
  rows.push(["", "SUMME UNBEZAHLT", "unbezahlt", sumUnpaidNet, sumUnpaidProfit, unpaidMargin]);
  rows.push(["", "GESAMTSUMME OFFENE BELEGE", "", grandTotalNet, grandTotalProfit, grandTotalMargin]);

  writeToSheet(sheet, rows);

  sheet.getRange("A:A").setNumberFormat("@");

  if (rows.length > 1) {
    sheet.getRange(2, 4, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
    sheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat("0.00%");
  }

  const totalRows = rows.length;
  if (totalRows >= 4) {
    sheet.getRange(totalRows - 2, 1, 1, headers.length)
         .setBackground("#34495e")
         .setFontColor("#ffffff")
         .setFontWeight("bold");

    sheet.getRange(totalRows - 1, 1, 1, headers.length)
         .setBackground("#34495e")
         .setFontColor("#ffffff")
         .setFontWeight("bold");

    sheet.getRange(totalRows, 1, 1, headers.length)
         .setBackground("#1f2a38")
         .setFontColor("#ffffff")
         .setFontWeight("bold");
  }
}

// =========================================================================
// 14. VORGÄNGE OFFENE SUMMEN (Vorgänge_Offene_Summen)
// =========================================================================
function createOpenOrderSummarySheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_OPEN_SUMMARY);

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedCodes = allVorgangCodes.map(code => ({ string: code }));

  const query = `
    query GetOpenSummary($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldArt
              fldVtrNr
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldMge
                fldEPrNt
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

  const repSummary = {};
  for (let repId in CONFIG.REPS) {
    repSummary[repId] = { name: CONFIG.REPS[repId], paidAmount: 0, unpaidAmount: 0, totalAmount: 0 };
  }
  repSummary["UNASSIGNED"] = { name: "Nicht zugeordnet / Sonstige", paidAmount: 0, unpaidAmount: 0, totalAmount: 0 };

  let totalPaid = 0;
  let totalUnpaid = 0;
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node;
      const artCode = String(node.fldArt || "");
      const isPaid = CONFIG.PROCESS_TYPES.PAID.includes(artCode);
      const repInfo = resolveRepInfo(node);
      const repKey = repSummary[repInfo.repId] ? repInfo.repId : "UNASSIGNED";

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        const lineTotal = (pos.fldMge || 0) * (pos.fldEPrNt || 0);
        if (isPaid) {
          repSummary[repKey].paidAmount += lineTotal;
          totalPaid += lineTotal;
        } else {
          repSummary[repKey].unpaidAmount += lineTotal;
          totalUnpaid += lineTotal;
        }
        repSummary[repKey].totalAmount += lineTotal;
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  const headers = ["Vertreter-ID", "Vertreter Name", "Bezahlte Bestellungen (€)", "Unbezahlte Bestellungen (€)", "Offener Gesamtwert (€)"];
  const rows = [headers];

  Object.keys(repSummary).forEach(key => {
    const item = repSummary[key];
    if (item.totalAmount > 0 || key !== "UNASSIGNED") {
      rows.push([key === "UNASSIGNED" ? "" : key, item.name, item.paidAmount, item.unpaidAmount, item.totalAmount]);
    }
  });

  rows.push(["", "GESAMTSUMME OFFENE BESTELLUNGEN", totalPaid, totalUnpaid, totalPaid + totalUnpaid]);

  writeToSheet(sheet, rows);
  sheet.getRange("A:A").setNumberFormat("@");
  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 3).setNumberFormat("#,##0.00 €");
  }

  sheet.getRange(rows.length, 1, 1, headers.length)
       .setBackground("#1f2a38")
       .setFontColor("#ffffff")
       .setFontWeight("bold");
}

// =========================================================================
// 15. OFFENE BESTELLUNGEN DETAILS (Offene_Bestellungen_Details)
// =========================================================================
function createOpenOrderDetailsSheet(ss) {
  let sheet = prepareSheet(ss, CONFIG.SHEET_OPEN_DETAILS);

  const allVorgangCodes = [...CONFIG.PROCESS_TYPES.PAID, ...CONFIG.PROCESS_TYPES.UNPAID];
  const formattedCodes = allVorgangCodes.map(code => ({ string: code }));

  const query = `
    query GetOpenOrders($codes: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: { in: { field: fldArt, values: $codes } }
        ) {
          edges {
            node {
              fldBelegNr
              fldAuftrNr
              fldArt
              lblArt
              fldDat
              fldLiefDat
              fldVtrNr
              fldAdrNr
              fldReNa1
              fldReNa2
              rowAdresse { fldVtrNr }
              rowsPositions {
                fldArtNr
                fldMge
                fldOMge
                fldEPrNt
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
    "Belegdatum", "Lieferdatum", "Beleg-Nr.", "Auftrags-Nr.", "Vorgangs-Code", 
    "Vorgangsart", "Zahlungsstatus", "Vertreter-ID", "Vertreter Name", "Kunden-Nr.", 
    "Kundenname", "Artikel-Nr.", "Artikelbezeichnung", "Menge", 
    "Offene Menge (Bestand)", "Einzelpreis Netto (€)", "Gesamtpreis Netto (€)"
  ];
  
  const rows = [headers];
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const payload = { query: query, variables: { codes: formattedCodes, cursor: cursor } };
    const json = callGraphQL(payload);
    if (!json || !json.data) break;

    const conRead = json.data.tblVorgang.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const node = edge.node || {};
      const artCode = String(node.fldArt || "");
      const isPaid = CONFIG.PROCESS_TYPES.PAID.includes(artCode);
      const paymentStatus = isPaid ? "Bezahlt" : "Unbezahlt";
      
      const repInfo = resolveRepInfo(node);
      let custNr = String(node.fldAdrNr || "").trim();
      let custName = String(node.fldReNa2 || node.fldReNa1 || "").trim();

      if (!custNr) custNr = "OHNE_KUNDENNR";
      if (!custName) custName = "Ohne Kundennummer (Gast/Laufkunde)";

      const positions = node.rowsPositions || [];
      positions.forEach(pos => {
        let artNr = String(pos.fldArtNr || "").trim();
        if (!artNr) artNr = "SONSTIGE";

        const qty = pos.fldMge || 0;
        const oQty = pos.fldOMge || 0;
        const unitPrice = pos.fldEPrNt || 0;
        const totalPrice = qty * unitPrice;

        rows.push([
          node.fldDat ? new Date(node.fldDat) : "",
          node.fldLiefDat ? new Date(node.fldLiefDat) : "",
          node.fldBelegNr || "",
          node.fldAuftrNr || "",
          artCode,
          node.lblArt || artCode,
          paymentStatus,
          repInfo.repId,
          repInfo.repName,
          custNr,
          custName,
          artNr,
          pos.rowArtikel?.fldKuBez1 || (artNr === "SONSTIGE" ? "Sonderposition / Sonstige" : artNr),
          qty,
          oQty,
          unitPrice,
          totalPrice
        ]);
      });
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  writeToSheet(sheet, rows);
  
  sheet.getRange("C:F").setNumberFormat("@");
  sheet.getRange("H:H").setNumberFormat("@");
  sheet.getRange("J:J").setNumberFormat("@");
  sheet.getRange("L:L").setNumberFormat("@");

  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 2).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(2, 14, rows.length - 1, 2).setNumberFormat("#,##0");
    sheet.getRange(2, 16, rows.length - 1, 2).setNumberFormat("#,##0.00 €");
  }
}

// =========================================================================
// HILFSFUNKTIONEN ZUM SICHEREN SCHREIBEN & FORMATIEREN
// =========================================================================

function prepareSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.clear();
  sheet.clearFormats();
  return sheet;
}

function writeToSheet(sheet, data) {
  if (!data || data.length === 0) return;

  let maxCols = 0;
  for (let i = 0; i < data.length; i++) {
    if (Array.isArray(data[i]) && data[i].length > maxCols) {
      maxCols = data[i].length;
    }
  }

  if (maxCols === 0) return;

  const paddedData = data.map(row => {
    const r = Array.isArray(row) ? row.slice() : [row];
    while (r.length < maxCols) {
      r.push("");
    }
    return r;
  });

  const numRows = paddedData.length;
  sheet.getRange(1, 1, numRows, maxCols).setValues(paddedData);
}

function applyGlobalFormatting(ss) {
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) return;

    sheet.setFrozenRows(1);

    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
    fullRange.createFilter();

    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground("#2c3e50")
               .setFontColor("#ffffff")
               .setFontWeight("bold")
               .setFontSize(10);

    fullRange.setBorder(true, true, true, true, true, true, "#d3d3d3", SpreadsheetApp.BorderStyle.SOLID);

    sheet.autoResizeColumns(1, lastCol);
    for (let col = 1; col <= lastCol; col++) {
      let currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, Math.min(Math.max(currentWidth + 12, 100), 300));
    }
  });
}

function safeToast(ss, message, title) {
  try {
    ss.toast(message, title, 5);
  } catch (e) {
    Logger.log(`[${title}] ${message}`);
  }
}

function callGraphQL(payload) {
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "X-API-Token": CONFIG.API_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("GraphQL API-Fehler: " + e.toString());
    return null;
  }
}