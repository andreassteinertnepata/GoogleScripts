// ==========================================
// FILE: DatahubService.gs
// ==========================================
function fetchAlteVorkasseVorgaenge(startTime) {
  const formattedVorgangsArten = CONFIG.VORGANGSARTEN.map(art => ({ string: art }));

  // Stichtag: Heute minus 3 Monate
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);

  const query = `
    query GetAlteVorkasseVorgaengeDetails($vorgangsArten: [FilterValue!]!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            in: { field: fldArt, values: $vorgangsArten }
          }
        ) {
          edges {
            node {
              fldBelegNr
              fldAuftrNr
              fldArt
              fldVtrNr
              fldAdrNr
              fldDat
              fldLiefDat
              fldGspKz
              fldReNa2
              fldReNa3
              fldStorniertKz
              fldErstDat
              rowsPositions {
                fldArtNr
                fldMge
                fldEPrNt
                rowArtikel {
                  fldKuBez1
                  fldKuBez3
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
  const belegeNachVertreter = {};
  const detailsRows = [];
  let totalFoundDocs = 0;

  safeToast("Lade Belege & Positionen aus dem Datahub...", "Bitte warten");

  while (hasNextPage) {
    if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME_MS) {
      Logger.log("Zeitlimit erreicht. Breche Pagination ab.");
      safeToast("Zeitlimit erreicht, verarbeite bisherige Daten.", "Hinweis");
      break;
    }

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: {
          vorgangsArten: formattedVorgangsArten,
          cursor: cursor
        }
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("GraphQL Fehler: " + JSON.stringify(json.errors || response.getContentText()));
        safeToast("Serverfehler bei der Datahub-Abfrage.", "Fehler");
        return { belegeNachVertreter, detailsRows, totalFoundDocs: 0 };
      }

      const conRead = json.data.tblVorgang?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        const node = edge.node;
        if (!node) return;

        // 1. Stornierte Belege ignorieren (JS Filter wg. NULL-Falle)
        if (node.fldStorniertKz === true) return;

        // 2. Belegdatum (oder Erfassungsdatum)
        const rawDate = node.fldDat || node.fldErstDat;
        if (!rawDate) return;
        const docDate = new Date(rawDate);

        // 3. Datumsfilter (> 3 Monate alt)
        if (docDate >= threeMonthsAgo) return;

        // 4. Vertreternummer ermitteln & normalisieren
        let rawVtr = String(node.fldVtrNr || "").trim();
        let vtrNr = rawVtr !== "" && rawVtr !== "0" ? parseInt(rawVtr, 10).toString() : "FALLBACK";

        const belegNr = String(node.fldBelegNr || "Ohne Nummer");
        const liefDate = node.fldLiefDat ? new Date(node.fldLiefDat) : null;

        // 5. Positionen und Gesamtsummen aufbereiten
        const positionen = node.rowsPositions || [];
        let docTotalMge = 0;
        let docTotalNetto = 0;
        const posList = [];

        if (positionen.length > 0) {
          positionen.forEach(pos => {
            const mge = pos.fldMge || 0;
            const eprNt = pos.fldEPrNt || 0;
            const lineNetto = mge * eprNt;

            docTotalMge += mge;
            docTotalNetto += lineNetto;

            const artNr = String(pos.fldArtNr || "");
            const kuBez1 = String(pos.rowArtikel?.fldKuBez1 || "");
            const kuBez3 = String(pos.rowArtikel?.fldKuBez3 || "");

            posList.push({
              artNr: artNr,
              kuBez1: kuBez1,
              kuBez3: kuBez3,
              mge: mge,
              eprNt: eprNt,
              lineNetto: lineNetto
            });

            detailsRows.push([
              belegNr,
              String(node.fldAuftrNr || ""),
              String(node.fldArt || ""),
              String(node.fldVtrNr || ""),
              String(node.fldAdrNr || ""),
              docDate,
              liefDate || "",
              node.fldGspKz || false,
              String(node.fldReNa2 || ""),
              String(node.fldReNa3 || ""),
              artNr,
              kuBez1,
              kuBez3,
              mge
            ]);
          });
        } else {
          detailsRows.push([
            belegNr,
            String(node.fldAuftrNr || ""),
            String(node.fldArt || ""),
            String(node.fldVtrNr || ""),
            String(node.fldAdrNr || ""),
            docDate,
            liefDate || "",
            node.fldGspKz || false,
            String(node.fldReNa2 || ""),
            String(node.fldReNa3 || ""),
            "", "", "", 0
          ]);
        }

        const docObj = {
          belegNr: belegNr,
          auftrNr: String(node.fldAuftrNr || ""),
          art: String(node.fldArt || ""),
          vtrNr: String(node.fldVtrNr || ""),
          adrNr: String(node.fldAdrNr || ""),
          docDate: docDate,
          liefDate: liefDate,
          reNa2: String(node.fldReNa2 || ""),
          reNa3: String(node.fldReNa3 || ""),
          totalMge: docTotalMge,
          totalNetto: docTotalNetto,
          positions: posList
        };

        if (!belegeNachVertreter[vtrNr]) {
          belegeNachVertreter[vtrNr] = [];
        }

        if (!belegeNachVertreter[vtrNr].some(d => d.belegNr === belegNr)) {
          belegeNachVertreter[vtrNr].push(docObj);
          totalFoundDocs++;
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Skriptfehler: " + e.message);
      safeToast("Netzwerkfehler: " + e.message, "Fehler");
      return { belegeNachVertreter, detailsRows, totalFoundDocs: 0 };
    }
  }

  return { belegeNachVertreter, detailsRows, totalFoundDocs };
}