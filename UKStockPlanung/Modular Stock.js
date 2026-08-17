function fetchStuecklisten() {
  const BLATT_NAME = "Modular Stock";

  // --- WERTESAMMLUNG AUS CONFIG.GS ---
  const API_URL = CONFIG.NEPATA_GRAPHQL_URL;
  const API_TOKEN = CONFIG.NEPATA_API_TOKEN;
  const ERLAUBTE_LAGER = CONFIG.ERLAUBTE_LAGER; 
  const ARTIKEL_KATALOGE = CONFIG.ARTIKEL_KATALOGE;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);
  if (!sheet) sheet = ss.insertSheet(BLATT_NAME);
  
  sheet.clear(); 
  

  // Dynamische Kataloge aus Config in das GraphQL-JSON-Format umwandeln
  const katalogValuesJson = ARTIKEL_KATALOGE.map(k => `{ string: "${k}" }`).join(", ");

  const query = `
    query GetStuecklisten($cursor: String) {
      tblArtikel {
        conRead(
          first: 100, 
          after: $cursor,
          fastFilter: {
            and: [
              {
                in: {
                  field: fldKatalog,
                  values: [ ${katalogValuesJson} ]
                }
              },
              {
                ne: [
                  { field: fldArtikelArt },
                  { value: { string: "0" } }
                ]
              }
            ]
          }
        ) {
          edges {
            node {
              fldArtNr
              fldKuBez1
              fldKuBez3
              fldKatalog
              rowsStueckliste {
                fldArtNr
                fldMge
                rowArtikel {
                  fldKuBez1
                  fldKuBez3
                  rowsArtikelLager {
                    fldLagNr
                    fldMge
                    fldKBstMge
                  }
                  rowsLieferantenbestelleingang {
                    fldArtNr
                    fldLagNr
                    fldLiefDat
                    fldOMge
                  }
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
  let zeilen = [];
  let pageCount = 0;
  
  const startTime = Date.now(); 
  const MAX_TIME_MS = 5 * 60 * 1000; 
  const heute = new Date();

  // Englische Header
  const headers = [
    "Main Item SKU", 
    "Main Item Description (DE)", 
    "Main Item Description (EN)", 
    "Catalog",
    "Max. Available Sets (Immediate)", 
    "Available Date (Main Item)", 
    "Qty in Set",
    "Component SKU", 
    "Component Description (DE)", 
    "Component Description (EN)", 
    "Stock Qty (Component)",
    "Customer Orders (Component)",
    "Free Stock (Component)",
    "Possible Sets (from Component)",
    "Next Delivery Date (Component)", 
    "Incoming Qty Next Delivery (Component)"
  ];
  zeilen.push(headers);

  while (hasNextPage) {
    pageCount++;
    Logger.log("Modular Stock: Loading page " + pageCount + "... (" + (zeilen.length - 1) + " rows found so far)");

    if (Date.now() - startTime > MAX_TIME_MS) {
      Logger.log("⚠️ Time limit almost reached! Stopping Modular Stock fetch.");
      break; 
    }

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": API_TOKEN },
      payload: JSON.stringify({ query: query, variables: { cursor: cursor } }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`Modular Stock Fetch HTTP ${responseCode}: ${responseText}`);
    }

    const json = JSON.parse(responseText);

    if (json.errors) {
      throw new Error(`GraphQL Errors in Modular Stock: ${JSON.stringify(json.errors)}`);
    }

    const conRead = json.data?.tblArtikel?.conRead || {};
    const edges = conRead.edges || [];

    edges.forEach(edge => {
      const hauptArtikel = edge.node;
      const komponenten = hauptArtikel.rowsStueckliste || [];

      if (komponenten.length > 0) {
        let maxBaubareSetsHauptartikel = Infinity;
        let komponetenAnalyse = [];

        // STEP 1: Component analysis (Lager-Filter aus CONFIG)
        komponenten.forEach(komp => {
          const kompDetails = komp.rowArtikel || {};
          const lagerListe = kompDetails.rowsArtikelLager || [];
          const bestellEingaenge = kompDetails.rowsLieferantenbestelleingang || [];

          let lagerMenge = 0;
          let kundenBestellMenge = 0;

          lagerListe.forEach(l => {
            if (ERLAUBTE_LAGER.includes(l.fldLagNr)) {
              lagerMenge += (l.fldMge || 0);
              kundenBestellMenge += (l.fldKBstMge || 0);
            }
          });

          const mengeImSet = komp.fldMge || 0;
          const freierBestand = lagerMenge - kundenBestellMenge;

          let moeglicheSetsProKomp = 0;
          if (mengeImSet > 0) {
            moeglicheSetsProKomp = Math.floor(Math.max(0, freierBestand) / mengeImSet);
          }

          if (moeglicheSetsProKomp < maxBaubareSetsHauptartikel) {
            maxBaubareSetsHauptartikel = moeglicheSetsProKomp;
          }

          // Next delivery date for this component
          let kompNaechstesDatum = null;
          let kompOffeneMenge = 0;

          if (bestellEingaenge.length > 0) {
            const valideEingaenge = bestellEingaenge
              .filter(b => 
                b.fldLiefDat && 
                (b.fldOMge || 0) > 0 && 
                (b.fldLagNr === undefined || b.fldLagNr === null || ERLAUBTE_LAGER.includes(b.fldLagNr))
              )
              .map(b => ({
                datum: new Date(b.fldLiefDat),
                menge: b.fldOMge || 0
              }))
              .sort((a, b) => a.datum - b.datum);

            if (valideEingaenge.length > 0) {
              kompNaechstesDatum = valideEingaenge[0].datum;
              kompOffeneMenge = valideEingaenge[0].menge;
            }
          }

          komponetenAnalyse.push({
            komp: komp,
            kompDetails: kompDetails,
            lagerMenge: lagerMenge,
            kundenBestellMenge: kundenBestellMenge,
            freierBestand: freierBestand,
            mengeImSet: mengeImSet,
            moeglicheSetsProKomp: moeglicheSetsProKomp,
            kompNaechstesDatum: kompNaechstesDatum,
            kompOffeneMenge: kompOffeneMenge
          });
        });

        if (maxBaubareSetsHauptartikel === Infinity) {
          maxBaubareSetsHauptartikel = 0;
        }

        // STEP 2: Main Item availability date
        let gesamtLieferdatumHauptartikel = null;

        if (maxBaubareSetsHauptartikel >= 1) {
          gesamtLieferdatumHauptartikel = heute;
        } else {
          let lieferTermineFehlteile = [];
          let fehlteilOhneBestellung = false;

          komponetenAnalyse.forEach(k => {
            if (k.freierBestand < k.mengeImSet) {
              if (k.kompNaechstesDatum) {
                lieferTermineFehlteile.push(k.kompNaechstesDatum.getTime());
              } else {
                fehlteilOhneBestellung = true;
              }
            }
          });

          if (fehlteilOhneBestellung) {
            gesamtLieferdatumHauptartikel = "Pending (PO missing)";
          } else if (lieferTermineFehlteile.length > 0) {
            const maxTimestamp = Math.max(...lieferTermineFehlteile);
            gesamtLieferdatumHauptartikel = new Date(maxTimestamp);
          } else {
            gesamtLieferdatumHauptartikel = "Out of Stock";
          }
        }

        // STEP 3: Write rows
        komponetenAnalyse.forEach(k => {
          zeilen.push([
            hauptArtikel.fldArtNr || "",
            hauptArtikel.fldKuBez1 || "",
            hauptArtikel.fldKuBez3 || "", 
            hauptArtikel.fldKatalog || "",
            maxBaubareSetsHauptartikel,
            gesamtLieferdatumHauptartikel, 
            k.mengeImSet, 
            k.komp.fldArtNr || "", 
            k.kompDetails.fldKuBez1 || "", 
            k.kompDetails.fldKuBez3 || "", 
            k.lagerMenge,          
            k.kundenBestellMenge,
            k.freierBestand,
            k.moeglicheSetsProKomp,
            k.kompNaechstesDatum ? k.kompNaechstesDatum : "", 
            k.kompOffeneMenge                               
          ]);
        });
      }
    });

    hasNextPage = conRead.pageInfo?.hasNextPage || false;
    cursor = conRead.pageInfo?.endCursor || null;
  }

  // --- WRITE & NUMBER FORMATS ONLY ---
  if (zeilen.length === 1) {
    sheet.getRange(1, 1).setValue("No matching BOM items found in these catalogs.");
    return;
  }

  const totalRows = zeilen.length;
  const totalCols = headers.length;

  sheet.getRange("A:A").setNumberFormat("@"); 
  sheet.getRange("D:D").setNumberFormat("@"); 
  sheet.getRange("H:H").setNumberFormat("@"); 

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(zeilen);

  if (totalRows > 1) {
    sheet.getRange(2, 5, totalRows - 1, 1).setNumberFormat("#,##0"); 
    sheet.getRange(2, 6, totalRows - 1, 1).setNumberFormat("yyyy-mm-dd;@"); 
    sheet.getRange(2, 7, totalRows - 1, 1).setNumberFormat("#,##0.00"); 
    sheet.getRange(2, 11, totalRows - 1, 3).setNumberFormat("#,##0.00"); 
    sheet.getRange(2, 14, totalRows - 1, 1).setNumberFormat("#,##0"); 
    sheet.getRange(2, 15, totalRows - 1, 1).setNumberFormat("yyyy-mm-dd;@"); 
    sheet.getRange(2, 16, totalRows - 1, 1).setNumberFormat("#,##0.00");    
  }

  // HINWEIS: Rahmen, Hintergrundfarben und Spaltenbreiten übernimmt am Ende 
  // deine 'formatGoogleSheet()'-Funktion aus der main()!
}