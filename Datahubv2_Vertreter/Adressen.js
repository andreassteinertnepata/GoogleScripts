// --- KONFIGURATION ---
const API_URL = "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql";
const TOKEN = "e12Bfv!@Ss#asrpPFjucm8a8";
const BLATT_NAME = "Adressen";

// GraphQL-Query ohne fastFilter: Holt absolut alle Adressen, sortiert nach der neuesten Änderung!
// GraphQL-Query mit Vorfilter: Schließt alle Adressen ohne "AbwArtDatGrp" direkt auf Serverebene aus!
const GRAPHQL_QUERY = `
query GetAdressen($cursor: String) {
  tblAdresse {
    conRead(
      first: 100, 
      after: $cursor, 
      orderBy: [{ field: fldLtzAend, desc: true }], # Holt die neuesten Änderungen ZUERST
      fastFilter: {
        isNotNull: { field: fldAbwArtDatGrp }      # Lässt nur Zeilen durch, die hier einen Wert haben
      }
    ) {
      edges {
        node {
          fldAdrNr
          fldKredLimit
          fldVtrNr
          fldAbwArtDatGrp
          rowsAnschriften {
            fldNa2
            fldNa3
            fldLandBez
            fldEMail1
            fldEMail2
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

function importGraphQLData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BLATT_NAME);
  
  // 1. Blatt bereinigen
  if (sheet) {
    sheet.clear(); 
  } else {
    sheet = ss.insertSheet(BLATT_NAME);
  }
  
  // Header-Zeile schreiben (Spalten A bis I)
  const headers = [
    "AdrNr", "KredLimit", "VtrNr", "AbwArtDatGrp", 
    "Name2", "Name3", "LandBez", "EMail1", "EMail2"
  ];
  sheet.appendRow(headers);
  
  // Schickes Design für den Header
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#34495e")
       .setFontColor("#ffffff")
       .setFontWeight("bold");

  let allRows = [];
  let hasNextPage = true;
  let cursor = null; 
  let pageCount = 0;
  const MAX_PAGES = 500; // Schutzbremse erhöht, da ohne Filter mehr Seiten geladen werden
  
  // Eindeutigkeitsspeicher für Kundennummern
  const verarbeiteteKunden = new Set();
  
  Logger.log("Starte API-Abruf für alle Adressen (ohne Filter)...");
  
  // 2. Daten seitenweise per API-Call abrufen
  while (hasNextPage && pageCount < MAX_PAGES) {
    pageCount++;
    Logger.log("Lade Seite " + pageCount + "... (Aktuelle Zeilen im Speicher: " + allRows.length + ")");
    
    const payload = {
      query: GRAPHQL_QUERY,
      variables: {
        cursor: cursor
      }
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { 
        "X-API-Token": TOKEN 
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      
      if (response.getResponseCode() !== 200) {
        Logger.log("Fehler bei der API: " + response.getContentText());
        break;
      }
      
      const json = JSON.parse(response.getContentText());
      
      if (json.errors) {
        Logger.log("GraphQL Fehler: " + JSON.stringify(json.errors));
        break;
      }
      
      const conRead = json.data?.tblAdresse?.conRead || {};
      const edges = conRead.edges || [];
      
      if (edges.length === 0) {
        hasNextPage = false;
        break;
      }
      
      // 3. Struktur entpacken (Echte Eindeutigkeit erzwingen)
      for (let i = 0; i < edges.length; i++) {
        const node = edges[i].node || {};
        const kundenNummer = node.fldAdrNr ? node.fldAdrNr.toString().trim() : "";
        
        // Wenn wir diese Kundennummer in diesem Durchlauf schon importiert haben,
        // überspringen wir jede weitere (ältere) Version konsequent!
        if (verarbeiteteKunden.has(kundenNummer) || kundenNummer === "") {
          continue;
        }
        
        // Da fastFilter weg ist, fldAbwArtDatGrp auf "null" oder leeren Text prüfen, 
        // aber nicht mehr die Zeile komplett überspringen, sondern Standardwert "" setzen.
        let abwGruppe = node.fldAbwArtDatGrp;
        if (!abwGruppe || abwGruppe.toString().trim() === "null") {
          abwGruppe = "";
        }

        const baseData = [
          kundenNummer,
          node.fldKredLimit || 0,
          node.fldVtrNr || "",
          abwGruppe
        ];
        
        const anschriftenList = node.rowsAnschriften || [];
        
        if (anschriftenList.length === 0) {
          allRows.push(baseData.concat(["", "", "", "", ""]));
          verarbeiteteKunden.add(kundenNummer); // Als erledigt markieren
        } else {
          // Wir nehmen nur die allererste (primäre) Anschrift der Adresse
          const inner = anschriftenList[0] || {};
          const flatRow = baseData.concat([
            inner.fldNa2 || "",
            inner.fldNa3 || "",
            inner.fldLandBez || "",
            inner.fldEMail1 || "",
            inner.fldEMail2 || ""
          ]);
          allRows.push(flatRow);
          verarbeiteteKunden.add(kundenNummer); // Als erledigt markieren
        }
      }
      
      hasNextPage = conRead.pageInfo?.hasNextPage || false;
      cursor = conRead.pageInfo?.endCursor || null;
      
    } catch (e) {
      Logger.log("Fehler bei der Abfrage: " + e.toString());
      break;
    }
  }
  
  // 4. Daten gesammelt ins Google Sheet schreiben
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
    
    // 5. FORMEL FÜR KUNDENART IN SPALTE J (Spalte 10) SCHREIBEN
    const kundenartCell = sheet.getRange("J1");
    kundenartCell.setFormula(`={"Kundenart"; ARRAYFORMULA(IF(D2:D=""; ""; CHOOSE(D2:D+1; "YOW"; "A"; "B"; "C")))}`);
    kundenartCell.setBackground("#34495e")
                 .setFontColor("#ffffff")
                 .setFontWeight("bold");

    // 6. FORMEL FÜR VERTRETER IN SPALTE K (Spalte 11) SCHREIBEN
    const formulaCell = sheet.getRange("K1");
    formulaCell.setFormula(`={"Vertreter"; ARRAYFORMULA(IF(C2:C=""; ""; VLOOKUP(C2:C; IMPORTRANGE("https://docs.google.com/spreadsheets/d/1fM-PheFfirTWpd2dtEIwFUqOka-N3tRHzadI0LuB9qs/edit?gid=0#gid=0"; "Vertreter!A:D"); 2; FALSE)))}`);
    formulaCell.setBackground("#34495e")
               .setFontColor("#ffffff")
               .setFontWeight("bold");

    // Spaltenbreiten automatisch anpassen (Spalten 1 bis 11)
    sheet.autoResizeColumns(1, 11);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Erfolgreich " + allRows.length + " Adressnummern ohne Filter geladen!", 
      "Import beendet", 
      5
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Keine passenden Adressdaten gefunden.", "Hinweis", 5);
  }
}