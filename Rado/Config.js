// ==========================================
// GLOBALE KONFIGURATION (Config.gs)
// ==========================================
const CONFIG = {
  // --- API Zugangsdaten ---
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",

  

  // --- Benutzer & Filter ---
  VERTRETER_NR: "60",
  AUSGESCHLOSSENE_ADRESSEN: [],                   // IDs eintragen, um Kunden zu ignorieren
  ERLAUBTE_LAGER: ["1", "2", "200", "13"],        // Für fetchArticles
  ARTIKEL_KATALOGE: ["0", "1", "7","8", "13", "18"],  // Für fetchArticles

  // --- Externe Referenz-Dateien ---
  MASTER_SHEET_ID: "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ", // Vorgangsarten
  TRANS_SHEET_ID: "1urqEUeu1TdGBU-MA_lnUuGzuiWHosOix4qYaB6mrACQ",  // Übersetzungen

  // --- Externe Tabellenblatt-Namen ---
  TAB_VORGANGSARTEN: "Vorgangsarten",
  TAB_TRANS_HEADERS: "Header",
  TAB_TRANS_VALUES: "Values",
  TAB_TRANS_SHEETS: "Sheets",

  // --- Lokale Ziel-Blätter ---
  SHEET_ARCHIVE: "Archive",
  SHEET_OPEN_ORDERS: "Open Orders",
  SHEET_OUTSTANDING: "Outstanding Invoices",
  SHEET_ARTICLES: "Articles",
  SHEET_COLOR_MAPPING: "ColorMapping",

  // --- System-Blätter (werden nicht formatiert/übersetzt) ---
  IGNORED_SHEETS: ["Values", "Header", "Sheet", "ColorMapping", "Vorgangsarten", "Pivot-Invoices_Last_90_Days", "Pivot_Open_Orders", "Important Links"],

  // --- Corporate Design / Formatierung ---
  FORMAT_HEADER_BG: "#3bb7c4",      // Türkis
  FORMAT_HEADER_TEXT: "#000000",    // Schwarz
  FORMAT_ROW_EVEN: "#f4f8f9",       // Helles Grautürkis
  FORMAT_ROW_ODD: "#ffffff",        // Weiß
  FORMAT_BORDER_COLOR: "#000000",   // schwarz
  HEADER_FONT_SIZE: 11,
  COL_WIDTH_BUFFER: 20,             // Pixel für Filter-Pfeile
  MAX_COL_WIDTH: 250                // Maximale Spaltenbreite
};