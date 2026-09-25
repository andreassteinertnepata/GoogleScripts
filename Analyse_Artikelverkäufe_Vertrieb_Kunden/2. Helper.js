// ==========================================
// 2. HELPER.GS - Hilfsfunktionen & Dynamische Quartale
// ==========================================

function getActiveVorgangsarten() {
  const configSs = SpreadsheetApp.openByUrl(CONFIG.SHEETS.VORGANGSARTEN_URL);
  const sheet = configSs.getSheetByName(CONFIG.SHEETS.VORGANGSARTEN_TAB);
  
  if (!sheet) {
    throw new Error(`Das Blatt "${CONFIG.SHEETS.VORGANGSARTEN_TAB}" wurde nicht gefunden!`);
  }
  
  const data = sheet.getDataRange().getValues();
  const validTypes = [];
  
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0] || "").trim();
    const condition = String(data[i][2] || "").trim().toLowerCase(); // Spalte C
    
    // Spalte C gefüllt UND enthält "archiv"
    if (code !== "" && condition.includes("archiv")) {
      validTypes.push({ string: code });
    }
  }
  
  if (validTypes.length === 0) {
    throw new Error("Keine gültigen Vorgangsarten für 'Archiv' in Spalte C gefunden.");
  }
  return validTypes;
}

/**
 * Berechnet 8 Kalenderquartale ausgehend vom Anfang des aktuellen Quartals vor 24 Monaten.
 */
function getQuarterBoundaries() {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  
  const currentQStartMonth = Math.floor(currentMonth / 3) * 3;
  const currentQStart = new Date(currentYear, currentQStartMonth, 1);
  
  const quarters = [];
  
  for (let i = 7; i >= 0; i--) {
    const qStart = new Date(currentQStart);
    qStart.setMonth(currentQStart.getMonth() - (i * 3));
    
    const qEnd = new Date(qStart);
    qEnd.setMonth(qStart.getMonth() + 3);
    
    const qNum = Math.floor(qStart.getMonth() / 3) + 1;
    const qYear = qStart.getFullYear();
    const label = `Q${qNum}-${qYear}`;
    
    quarters.push({
      label: label,
      start: qStart,
      end: qEnd
    });
  }
  
  const start24M_Label = quarters[0].label;
  const start12M_Label = quarters[4].label;
  const end_Label = quarters[7].label;

  return {
    quarters: quarters,
    start24M: quarters[0].start,
    isoStart24M: quarters[0].start.toISOString(),
    periodString24M: `${start24M_Label} bis ${end_Label} (24M)`,
    periodString12M: `${start12M_Label} bis ${end_Label} (12M)`
  };
}

function assignQuarterLabel(recordDate, quarterBoundaries) {
  if (!recordDate) return "Außerhalb";
  for (let q of quarterBoundaries.quarters) {
    if (recordDate >= q.start && recordDate < q.end) {
      return q.label;
    }
  }
  return "Außerhalb";
}

function getRowColorByStatus(status) {
  const p = CONFIG.COLOR_PALETTE;
  switch (status) {
    case "Starkes Wachstum":
      return p.STRONG_GROWTH;
    case "Wachstum":
    case "Growth":
      return p.GROWTH;
    case "Rückgang":
    case "Decline":
      return p.DECLINE;
    case "Starker Abfall":
    case "Strong Decline":
      return p.STRONG_DECLINE;
    case "Inaktiv / Gestoppt":
    case "Kauf gestoppt":
    case "Churned / Stopped":
      return p.STOPPED;
    case "Nie gekauft":
    case "Never purchased":
      return p.NEVER_PURCHASED;
    default:
      return null;
  }
}