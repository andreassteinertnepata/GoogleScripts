function fetchOffenePosten() {
  const HEADERS = [
    "Vertreter-Nr.", "Kreditlimit", "Adress-Nr. (Kunde)", "Status",   // A, B, C, D
    "Name 2", "Name 3", "Land",                                       // E, F, G
    "fldBelegNr", "fldAuftrNr", "OP-Text",                             // H, I, J
    "Erstelldatum", "Netto Tage", "Bezahlt-Betrag",                   // K, L, M
    "OP-Saldo Betrag", "Mahnstufe", "Mahndatum",                      // N, O, P
    "Fällig?", "Fällig seit (Tagen)"                                  // Q, R
  ];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_OUTSTANDING);
  let isNewSheet = false;
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_OUTSTANDING);
    isNewSheet = true;
  }
  
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
   
  }
  
  if (isNewSheet) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
         .setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  }
  
  let hasNextPage = true;
  let cursor = null; 
  let allRows = [];
  
  const heute = new Date();
  heute.setHours(0, 0, 0, 0);
  
  while (hasNextPage) {
    const graphqlQuery = JSON.stringify({
      query: `
        query GetOffenePosten($cursor: String) {
          tblOffenerPosten {
            conRead(first: 100, after: $cursor) {
              edges {
                node {
                  rowAdresse {
                    fldVtrNr
                    fldKredLimit
                    fldAdrNr
                    fldStatus
                    rowsAnschriften {
                      fldNa2
                      fldNa3
                      fldLandBez
                    }
                  }
                  fldBelegNr
                  fldAuftrNr
                  fldText
                  fldErstDat
                  fldNettoTg
                  fldBezBet
                  fldOPSaldoBet
                  fldMahnSt
                  fldMahnDat
                }
              }
              pageInfo { hasNextPage endCursor }
            }
          }
        }
      `,
      variables: { cursor: cursor }
    });
    
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Token": CONFIG.API_TOKEN },
      payload: graphqlQuery,
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const json = JSON.parse(response.getContentText());
      if (json.errors) break;
      
      const conRead = json.data.tblOffenerPosten.conRead;
      conRead.edges.forEach(function(edge) {
        const node = edge.node;
        const adr = node.rowAdresse || {}; 
        const adrNr = adr.fldAdrNr || "";
        const status = adr.fldStatus || "";
        const vtrNr = adr.fldVtrNr ? adr.fldVtrNr.toString().trim() : "";
        
        const anschriften = adr.rowsAnschriften || [];
        const primaerAnschrift = anschriften.length > 0 ? anschriften[0] : {};
        
        const istLieferant = (status.trim() === "Lieferant");
        const istInAusschluss = CONFIG.AUSGESCHLOSSENE_ADRESSEN.includes(adrNr);
        const istZielVertreter = (vtrNr === CONFIG.VERTRETER_NR);
        
        if (!istLieferant && !istInAusschluss && istZielVertreter) {
          const erstDatStr = node.fldErstDat || ""; 
          const nettoTg = parseInt(node.fldNettoTg, 10) || 0; 
          let faelligkeitStr = ""; 
          let faelligSeitTage = ""; 
          
          if (erstDatStr) {
            const erstelldatum = new Date(erstDatStr);
            if (!isNaN(erstelldatum.getTime())) {
              const faelligkeitsDatum = new Date(erstelldatum);
              faelligkeitsDatum.setDate(faelligkeitsDatum.getDate() + nettoTg);
              faelligkeitStr = Utilities.formatDate(faelligkeitsDatum, Session.getScriptTimeZone(), "yyyy-MM-dd");
              
              faelligkeitsDatum.setHours(0, 0, 0, 0);
              if (faelligkeitsDatum < heute) {
                faelligSeitTage = Math.floor((heute.getTime() - faelligkeitsDatum.getTime()) / (1000 * 60 * 60 * 24));
              }
            }
          }
          
          allRows.push([
            adr.fldVtrNr || "",                  // Spalte A
            adr.fldKredLimit || 0,               // Spalte B
            adrNr,                               // Spalte C
            status,                              // Spalte D
            primaerAnschrift.fldNa2 || "",       // Spalte E
            primaerAnschrift.fldNa3 || "",       // Spalte F
            primaerAnschrift.fldLandBez || "",   // Spalte G
            node.fldBelegNr || "",               // Spalte H
            node.fldAuftrNr || "",               // Spalte I
            node.fldText || "",                  // Spalte J
            erstDatStr,                          // Spalte K
            nettoTg,                             // Spalte L
            node.fldBezBet || 0,                 // Spalte M
            node.fldOPSaldoBet || 0,             // Spalte N
            node.fldMahnSt || 0,                 // Spalte O
            node.fldMahnDat || "",               // Spalte P
            faelligkeitStr,                      // Spalte Q
            faelligSeitTage                      // Spalte R
          ]);
        }
      });
      
      hasNextPage = conRead.pageInfo.hasNextPage;
      cursor = conRead.pageInfo.endCursor;
    } catch (e) { break; }
  }
  
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, HEADERS.length).setValues(allRows);
    
    // Formate erzwingen:
    sheet.getRange("H:I").setNumberFormat("@");                             // Beleg- & Auftrags-Nr.
    sheet.getRange(2, 13, allRows.length, 2).setNumberFormat("#,##0.00 €"); // Beträge M & N (Bezahlt & OP-Saldo)
    sheet.getRange(2, 17, allRows.length, 1).setNumberFormat("yyyy-mm-dd"); // Spalte Q: Fällig?
    sheet.getRange(2, 18, allRows.length, 1).setNumberFormat("0");          // Spalte R: Fällig seit
  }
}