// ==========================================
// FILE: DetailsService.gs
// ==========================================

/**
 * Leert das Blatt "Auftragsdetails" zu Beginn eines Durchlaufs vollständig.
 */
function initDetailsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TAB_DETAILS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TAB_DETAILS);
  } else {
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    sheet.clear(); // Leert alle Inhalte und Formatierungen vollständig
  }
  return sheet;
}

/**
 * Schreibt die ermittelten Positionsdetails in das vorbereitete Tabellenblatt.
 */
function writeDetailsSheet(ss, detailsRows) {
  const sheet = initDetailsSheet(ss);

  const headers = [
    "fldBelegNr", "fldAuftrNr", "fldArt", "fldVtrNr", "fldAdrNr",
    "fldDat", "fldLiefDat", "fldGspKz", "fldReNa2", "fldReNa3",
    "fldArtNr", "fldKuBez1", "fldKuBez3", "fldMge"
  ];

  if (detailsRows.length === 0) {
    sheet.getRange(1, 1).setValue("Keine offenen Vorkasse-Vorgänge (>3 Monate) vorhanden.");
    return;
  }

  const outputData = [headers, ...detailsRows];
  const totalRows = outputData.length;
  const totalCols = headers.length;

  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(outputData);

  // Formate & Styles
  sheet.getRange("A:E").setNumberFormat("@");
  sheet.getRange("I:M").setNumberFormat("@");
  sheet.getRange(2, 6, totalRows - 1, 2).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(2, 14, totalRows - 1, 1).setNumberFormat("#,##0.00");

  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#2c3e50");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  range.setBorder(true, true, true, true, true, true);
  
  if (!sheet.getFilter()) {
    range.createFilter();
  }
  
  sheet.autoResizeColumns(1, totalCols);
}