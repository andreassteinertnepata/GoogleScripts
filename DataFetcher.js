/**
 * DataFetcher.js - GraphQL Abfragen & kaufmännische Verarbeitungslogik (MTD, YTD, 12M)
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

function getActiveVorgangsarten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_VORGANGSARTEN);
  
  const result = { archiv: [], vorgaenge: [] };

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const code = String(data[i][0]).trim();
      const typ = String(data[i][2]).trim().toLowerCase();
      
      if (!code) continue;

      if (CONFIG.EXCLUDED_BELEGARTEN.includes(code)) continue;
      
      if (typ === "archiv" || typ === "x" || typ === "ja" || data[i][2] === true) {
        result.archiv.push(code);
      } else if (typ.includes("vorgäng") || typ.includes("vorgaeng")) {
        result.vorgaenge.push(code);
      }
    }
  }

  if (result.archiv.length === 0) {
    result.archiv = CONFIG.DEFAULT_ARCHIV_TYPES.filter(function(c) { return !CONFIG.EXCLUDED_BELEGARTEN.includes(c); });
  }
  if (result.vorgaenge.length === 0) {
    result.vorgaenge = CONFIG.DEFAULT_VORGAENGE_TYPES.filter(function(c) { return !CONFIG.EXCLUDED_BELEGARTEN.includes(c); });
  }

  return result;
}

function fetchSalesData(startTime) {
  const now = new Date();
  const dateFrom12M = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  const dateFromIso = dateFrom12M.toISOString();

  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const startOfYear = new Date(now.getFullYear(), 0, 1, 0, 0, 0);

  const vorgangsarten = getActiveVorgangsarten();

  const stats = {
    _totalArticles: {}
  };

  Object.keys(CONFIG.SALES_REPS).forEach(function(id) {
    stats[id] = {
      umsatzMTD: 0,
      umsatzYTD: 0,
      umsatz12M: 0,
      prognose: 0,
      articles: {},
      keyAccounts: {}
    };
  });

  fetchDataForTable("tblVorgangArchiv", dateFromIso, vorgangsarten.archiv, stats, true, startTime, startOfMonth, startOfYear);

  if (Date.now() - startTime < CONFIG.MAX_EXECUTION_TIME_MS) {
    fetchDataForTable("tblVorgang", dateFromIso, vorgangsarten.vorgaenge, stats, false, startTime, startOfMonth, startOfYear);
  }

  return stats;
}

function fetchDataForTable(tableName, dateFromIso, allowedTypes, stats, isArchiv, startTime, startOfMonth, startOfYear) {
  let hasNextPage = true;
  let cursor = null;

  if (!allowedTypes || allowedTypes.length === 0) return;

  const formattedTypes = allowedTypes.map(function(art) { return { string: art }; });

  const query = `
    query GetSalesData($cursor: String, $dateFrom: DateTime!, $vorgangsArten: [FilterValue!]!) {
      ${tableName} {
        conRead(
          first: ${CONFIG.PAGE_SIZE},
          after: $cursor,
          fastFilter: {
            and: [
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

  while (hasNextPage) {
    if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME_MS) break;

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { cursor: cursor, dateFrom: dateFromIso, vorgangsArten: formattedTypes }
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.errors) break;

    const tableData = (json.data && json.data[tableName]) ? json.data[tableName] : null;
    const conRead = (tableData && tableData.conRead) ? tableData.conRead : null;

    if (!conRead || !conRead.edges) break;

    conRead.edges.forEach(function(edge) {
      processBelegNode(edge.node, allowedTypes, stats, isArchiv, startOfMonth, startOfYear);
    });

    hasNextPage = conRead.pageInfo ? conRead.pageInfo.hasNextPage : false;
    cursor = conRead.pageInfo ? conRead.pageInfo.endCursor : null;
  }
}

function processBelegNode(node, allowedTypes, stats, isArchiv, startOfMonth, startOfYear) {
  if (node.fldSel14 === true) return;
  if (node.fldStorniertKz === true) return;

  const artCode = String(node.fldArt || "").trim();

  if (CONFIG.EXCLUDED_BELEGARTEN.includes(artCode)) return;

  const vtrNr = node.fldVtrNr ? String(node.fldVtrNr).trim() : null;
  const isAssignedRep = Boolean(vtrNr && stats[vtrNr]);

  const belegDatum = parseBelegDatum(node.fldDat) || parseBelegDatum(node.fldErstDat);
  const isMTD = Boolean(belegDatum && belegDatum.getTime() >= startOfMonth.getTime());
  const isYTD = Boolean(belegDatum && belegDatum.getTime() >= startOfYear.getTime());

  const isKorrektur = CONFIG.CORRECTION_TYPES.includes(artCode) ||
                      (node.fldBelegNr && String(node.fldBelegNr).startsWith("123"));

  let belegUmsatzGesamt = 0;

  if (node.rowsPositions && Array.isArray(node.rowsPositions)) {
    node.rowsPositions.forEach(function(pos) {
      if (pos.fldAbrPosKz === false) return;

      const artNr = String(pos.fldArtNr || (pos.rowArtikel ? pos.rowArtikel.fldArtNr : "") || "").trim();
      if (artNr && CONFIG.EXCLUDED_ARTICLES.includes(artNr)) return;

      const mge = pos.fldMge || 0;
      const preis = pos.fldEPrNt || 0;

      // ZEILENUMSATZ-BERECHNUNG:
      // Bei Korrekturbelegen/Gutschriften wird die gesamte Zeilensumme invertiert,
      // damit stornierende Belege den Umsatz exakt abziehen!
      let lineUmsatz = mge * preis;
      if (isKorrektur) {
        lineUmsatz = -lineUmsatz;
      }

      // 1. KUNDEN- & VERTRETERUMSATZ: Alle Abrechnungspositionen der Rechnung (inkl. Rabatte)
      belegUmsatzGesamt += lineUmsatz;

      if (isAssignedRep) {
        if (isArchiv) {
          stats[vtrNr].umsatz12M += lineUmsatz;
          if (isYTD) stats[vtrNr].umsatzYTD += lineUmsatz;
          if (isMTD) stats[vtrNr].umsatzMTD += lineUmsatz;
        } else {
          stats[vtrNr].prognose += lineUmsatz;
        }
      }

      // 2. ARTIKEL- & PRODUKTZIELE (STRIKT KATALOG 1)
      const katalog = pos.rowArtikel ? String(pos.rowArtikel.fldKatalog || "") : "";
      const isKatalog1 = (katalog === CONFIG.TARGET_CATALOG);

      if (artNr && isKatalog1) {
        if (!stats._totalArticles[artNr]) {
          stats._totalArticles[artNr] = { stueckMTD: 0, stueckYTD: 0, stueck12M: 0 };
        }
        if (isArchiv) {
          stats._totalArticles[artNr].stueck12M += mge;
          if (isYTD) stats._totalArticles[artNr].stueckYTD += mge;
          if (isMTD) stats._totalArticles[artNr].stueckMTD += mge;
        }

        if (isAssignedRep) {
          if (!stats[vtrNr].articles[artNr]) {
            stats[vtrNr].articles[artNr] = { stueckMTD: 0, stueckYTD: 0, stueck12M: 0, stueckPrognose: 0 };
          }
          if (isArchiv) {
            stats[vtrNr].articles[artNr].stueck12M += mge;
            if (isYTD) stats[vtrNr].articles[artNr].stueckYTD += mge;
            if (isMTD) stats[vtrNr].articles[artNr].stueckMTD += mge;
          } else {
            stats[vtrNr].articles[artNr].stueckPrognose += mge;
          }
        }
      }
    });
  }

  // Key Accounts Tracking mit fldReNa2 / fldReNa3 (Firmenname)
  if (isAssignedRep) {
    const adrNr = node.fldAdrNr ? String(node.fldAdrNr).trim() : null;
    if (adrNr) {
      const na1 = String(node.fldReNa1 || "").trim();
      const na2 = String(node.fldReNa2 || "").trim();
      const na3 = String(node.fldReNa3 || "").trim();

      const genericNames = ["firma", "company", "herr", "frau", "gmbh", "ag", "kunden-name"];
      let companyName = [na2, na3].filter(Boolean).join(" ").trim();
      
      if (!companyName || genericNames.includes(companyName.toLowerCase())) {
        if (na1 && !genericNames.includes(na1.toLowerCase())) {
          companyName = na1;
        } else if (na1 && na2) {
          companyName = `${na1} ${na2}`;
        } else {
          companyName = companyName || na1 || `Kunde ${adrNr}`;
        }
      }

      if (!stats[vtrNr].keyAccounts[adrNr]) {
        stats[vtrNr].keyAccounts[adrNr] = { 
          umsatzMTD: 0, 
          umsatzYTD: 0, 
          umsatz12M: 0, 
          prognose: 0, 
          kundenName: companyName
        };
      } else if (companyName && stats[vtrNr].keyAccounts[adrNr].kundenName.toLowerCase().startsWith("company")) {
        stats[vtrNr].keyAccounts[adrNr].kundenName = companyName;
      }

      if (isArchiv) {
        stats[vtrNr].keyAccounts[adrNr].umsatz12M += belegUmsatzGesamt;
        if (isYTD) stats[vtrNr].umsatzYTD += belegUmsatzGesamt;
        if (isMTD) stats[vtrNr].umsatzMTD += belegUmsatzGesamt;
      } else {
        stats[vtrNr].keyAccounts[adrNr].prognose += belegUmsatzGesamt;
      }
    }
  }
}