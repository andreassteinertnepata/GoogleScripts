// ==========================================
// FILE: LogService.gs
// ==========================================
function initLogSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TAB_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TAB_LOG);
  } else {
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    sheet.clear();
  }

  const headers = [
    "Zeitstempel", "Vertreter-Nr", "Vertreter / Empfänger", "E-Mail (An)", "Anzahl Belege", "Belegnummern"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.getRange("B:B").setNumberFormat("@");

  return sheet;
}

function writeLogEntry(ss, timestamp, vtrNr, repName, toEmail, belegListe) {
  let sheet = ss.getSheetByName(CONFIG.TAB_LOG);
  if (!sheet) {
    sheet = initLogSheet(ss);
  }

  sheet.appendRow([
    timestamp,
    vtrNr !== "FALLBACK" ? vtrNr : "0",
    repName,
    toEmail,
    belegListe.length,
    belegListe.join(", ")
  ]);

  sheet.autoResizeColumns(1, 6);
}