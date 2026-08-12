function fetchVorgaenge() {
  const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
  const API_TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
  
  // ZIEL-BLATT & REPRESENTATIVE-NR.
  const ZIEL_BLATT = "Open Orders"; // bzw. "Vorgaenge"


  // EXTERNE TAGEBUCH- & OVERVIEW-DATEI FÜR VORGANGSARTEN
  const EXTERNAL_SPREADSHEET_ID = "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ";
  const KATALOG_BLATT = "Vorgangsarten";

  // 26 Spalten-Überschriften
  const HEADERS_DE = [
    "fldBelegNr", "fldAuftrNr", "fldArt", "Vorgangsart_Klartext", "fldVtrNr", "fldAdrNr", "fldDat", "fldLiefDat", "fldGspKz",
    "fldReNa2", "fldReNa3", "fldReLandBez", 
    "fldLiNa2", "fldLiNa3", "fldLiStr", "fldLiPLZ", "fldLiOrt", "fldLiLandBez",
    "fldArtNr", "fldKuBez1", "fldKuBez3", "fldMge", "fldEPrNt", "Position Gesamtpreis Netto", "fldAusLagNr", "fldAbrPosKz"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Vorgangsarten aus dem EXTERNEN Blatt einlesen
  const vorgangsartenMap = new Map();
  
  try {
    const externalSs = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
    const katalogSheet = externalSs.getSheetByName(KATALOG_BLATT);
    
    if (katalogSheet) {
      const lastRowKatalog = katalogSheet.getLastRow();
      if (lastRowKatalog >= 2) {
        // Spalte A = Code, Spalte B = Klartext, Spalte C = Filter/Kategorie
        const katalogDaten = katalogSheet.getRange(2, 1, lastRowKatalog - 1, 3).getValues();
        
        katalogDaten.forEach(row => {
          const code = String(row[0] || "").trim();
          const klartext = String(row[1] || "").trim();
          const kategorie = String(row[2] || "").trim(); // Spalte C prüfen
          
          // REGEL: Nur übernehmen, wenn ein Code da ist UND Spalte C NICHT LEER ist!
          if (code !== "" && kategorie !== "") {
            vorgangsartenMap.set(code, klartext);
          }
        });
      }
    } else {
      Logger.log(`Achtung: Blatt '${KATALOG_BLATT}' im externen Sheet nicht gefunden.`);
    }
  } catch (e) {
    Logger.log("Fehler beim Zugriff auf das externe Vorgangsarten-Sheet: " + e.message);
  }

  // 2. Gefilterte Codes als Array für GraphQL formatieren
  let vorgangsArten = Array.from(vorgangsartenMap.keys());
  
  if (vorgangsArten.length === 0) {
    ss.toast("Keine gültigen Vorgangsarten mit gefüllter Spalte C gefunden.", "Hinweis");
    return;
  }
  
  const formattedVorgangsArten = vorgangsArten.map(art => ({ string: art }));

  // ==========================================
  // 3. FIX: Zielblatt vorbereiten & Pivot schützen
  // ==========================================
  let zielSheet = ss.getSheetByName(ZIEL_BLATT);
  let isNewSheet = false;
  
  if (!zielSheet) {
    zielSheet = ss.insertSheet(ZIEL_BLATT);
    isNewSheet = true;
  } 

  const currentFilter = zielSheet.getFilter();
  if (currentFilter) currentFilter.remove();
  
  // Nur Daten ab Zeile 2 löschen (schützt die Header in Zeile 1)
  const maxRows = zielSheet.getMaxRows();
  const maxCols = zielSheet.getMaxColumns();
  if (maxRows > 1) {
    zielSheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    zielSheet.getRange(2, 1, maxRows - 1, maxCols).clearFormat();
  }

  // Header nur schreiben, wenn das Blatt komplett neu ist
  if (isNewSheet) {
    zielSheet.getRange(1, 1, 1, HEADERS_DE.length).setValues([HEADERS_DE])
             .setBackground("#2c3e50")
             .setFontColor("#ffffff")
             .setFontWeight("bold");
  }

  // 4. GraphQL Query mit serverseitigem Filter für fldVtrNr & fldArt
  const query = `
    query GetVorgaenge($vorgangsArten: [FilterValue!]!, $vtrNr: String!, $cursor: String) {
      tblVorgang {
        conRead(
          first: 100,
          after: $cursor,
          fastFilter: {
            and: [
              {
                in: {
                  field: fldArt,
                  values: $vorgangsArten
                }
              },
              {
                eq: [
                  { field: fldVtrNr },
                  { value: { string: $vtrNr } }
                ]
              }
            ]
          }
        ) {
          edges {
            node {
              fldAdrNr
              fldArt
              fldAuftrNr
              fldBelegNr
              fldDat
              fldGspKz
              fldLiLandBez
              fldLiNa2
              fldLiNa3
              fldLiOrt
              fldLiPLZ
              fldLiStr
              fldLiefDat
              fldReLandBez
              fldReNa2
              fldReNa3
              fldVtrNr
              rowsPositions {
                fldAbrPosKz
                fldArtNr
                fldAusLagNr
                fldEPrNt
                fldMge
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

  let alleVorgaenge = [];
  let hasNextPage = true;
  let cursor = null;

  while (hasNextPage) {
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({
        query: query,
        variables: { 
          vorgangsArten: formattedVorgangsArten,
          vtrNr: CONFIG.VERTRETER_NR,
          cursor: cursor 
        }
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      const json = JSON.parse(response.getContentText());

      if (json.errors || !json.data) {
        Logger.log("FEHLER GraphQL: " + JSON.stringify(json.errors || response.getContentText()));
        ss.toast("Fehler bei der API-Abfrage.", "Fehler");
        return;
      }

      const conRead = json.data.tblVorgang?.conRead || {};
      const edges = conRead.edges || [];

      edges.forEach(edge => {
        if (edge.node) {
          alleVorgaenge.push(edge.node);
        }
      });

      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;

    } catch (e) {
      Logger.log("Fehler: " + e.message);
      ss.toast("Netzwerk- oder Skriptfehler: " + e.message, "Fehler");
      return;
    }
  }

  if (alleVorgaenge.length === 0) {
    ss.toast(`Keine passenden Vorgänge für Vertreter ${CONFIG.VERTRETER_NR} gefunden.`, "Info", 5);
    return;
  }

  // Das Array 'zeilen' speichert jetzt nur noch Daten, KEINE Header mehr!
  const zeilen = [];

  // 5. Daten aufbereiten
  alleVorgaenge.forEach(vorgang => {
    const artCode = String(vorgang.fldArt || "").trim();
    const artKlartext = vorgangsartenMap.get(artCode) || "";

    const kopfDaten = [
      String(vorgang.fldBelegNr || ""),
      String(vorgang.fldAuftrNr || ""),
      artCode,
      artKlartext,
      String(vorgang.fldVtrNr || ""),
      String(vorgang.fldAdrNr || ""),
      vorgang.fldDat ? new Date(vorgang.fldDat) : "",
      vorgang.fldLiefDat ? new Date(vorgang.fldLiefDat) : "",
      vorgang.fldGspKz || false,
      String(vorgang.fldReNa2 || ""),
      String(vorgang.fldReNa3 || ""),
      String(vorgang.fldReLandBez || ""),
      String(vorgang.fldLiNa2 || ""),
      String(vorgang.fldLiNa3 || ""),
      String(vorgang.fldLiStr || ""),
      String(vorgang.fldLiPLZ || ""),
      String(vorgang.fldLiOrt || ""),
      String(vorgang.fldLiLandBez || "")
    ];

    const positionen = vorgang.rowsPositions || [];

    if (positionen.length > 0) {
      positionen.forEach(pos => {
        const menge = pos.fldMge || 0;
        const einzelpreisNetto = pos.fldEPrNt || 0;
        const gesamtpreisNetto = menge * einzelpreisNetto;

        zeilen.push([
          ...kopfDaten,
          String(pos.fldArtNr || ""),
          String(pos.rowArtikel?.fldKuBez1 || ""),
          String(pos.rowArtikel?.fldKuBez3 || ""),
          menge,
          einzelpreisNetto,
          gesamtpreisNetto,
          String(pos.fldAusLagNr || ""),
          String(pos.fldAbrPosKz || "")
        ]);
      });
    } else {
      zeilen.push([
        ...kopfDaten,
        "", "", "", 0, 0, 0, "", ""
      ]);
    }
  });

  // 6. In Sheet schreiben & formatieren
  if (zeilen.length > 0) {
    zielSheet.getRange("A:F").setNumberFormat("@");
    zielSheet.getRange("J:R").setNumberFormat("@");
    zielSheet.getRange("S:S").setNumberFormat("@");
    zielSheet.getRange("T:U").setNumberFormat("@");
    zielSheet.getRange("Y:Z").setNumberFormat("@");

    // Zeilen ab Reihe 2 einfügen (unter den Headern)
    zielSheet.getRange(2, 1, zeilen.length, HEADERS_DE.length).setValues(zeilen);

    // Formatierungen aufbauen
    zielSheet.getRange(2, 7, zeilen.length, 2).setNumberFormat("dd.mm.yyyy");
    zielSheet.getRange(2, 22, zeilen.length, 1).setNumberFormat("#,##0");
    zielSheet.getRange(2, 23, zeilen.length, 2).setNumberFormat("#,##0.00 €");

    // Rahmen und Filter über den gesamten Bereich (Header + Daten)
    const fullRange = zielSheet.getRange(1, 1, zeilen.length + 1, HEADERS_DE.length);
    fullRange.setBorder(true, true, true, true, true, true);
    fullRange.createFilter();

    zielSheet.autoResizeColumns(1, HEADERS_DE.length);
  }
}