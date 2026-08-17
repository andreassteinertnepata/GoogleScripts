// ==========================================
// GLOBALE KONFIGURATION (Config.gs)
// ==========================================
const CONFIG = {
  // --- API Zugangsdaten ---
  NEPATA_GRAPHQL_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  NEPATA_API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  SECABO_STOCK_URL: "https://www.secabo.com/store-api/nepata/018f1aad15167b07b14e9954409c9ef5/stock",

  // --- Benutzer & Filter ---
  VERTRETER_NR: "56",                             // Vertreter-ID für Archiv-Abfragen
  ERLAUBTE_LAGER: ["1", "13", "2", "200", "202"], // Für Lagerbestand-Summen
  ARTIKEL_KATALOGE: ["0", "1", "7", "8", "13", "18"], // Für fetchArticles

  // --- Externe Referenz-Dateien ---
  MASTER_SHEET_ID: "1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ", // Enthält Vorgangsarten
  TRANS_SHEET_ID: "1urqEUeu1TdGBU-MA_lnUuGzuiWHosOix4qYaB6mrACQ",  // Übersetzungen

  // --- Externe Tabellenblatt-Namen ---
  TAB_VORGANGSARTEN: "Vorgangsarten",
  TAB_TRANS_HEADERS: "Header",
  TAB_TRANS_VALUES: "Values",
  TAB_TRANS_SHEETS: "Sheets",

  // --- Google Drive Einstellungen (Mobile Import) ---
  MOBILE_FOLDER_ID: "1vmZH1zESYgnKOhsXz-dMW_qLmWkMU-gh",
  MOBILE_FILE_PREFIX: "Stocklist_",
  MOBILE_DESIRED_COLUMNS: ["SKU", "Name", "Stock level", "On Hand"],

  // --- Lokale Ziel-Blätter ---
  SHEET_SECABO_STOCK: "Secabo Stock",
  SHEET_MOBILE: "Mobile",
  SHEET_ARTICLES: "Articles",
  SHEET_ARCHIVE: "Archive",
  SHEET_COLOR_MAPPING: "ColorMapping",

  // --- System-Blätter (werden bei Formatierung ignoriert) ---
  IGNORED_SHEETS: ["ColorMapping", "Values", "Header", "Sheet", "Vorgangsarten", "Pivot-Invoices_Last_90_Days", "Pivot_Open_Orders", "Important Links"],

  // --- Corporate Design / Formatierung ---
  FORMAT_HEADER_BG: "#3bb7c4",      // Türkis
  FORMAT_HEADER_TEXT: "#000000",    // Schwarz
  FORMAT_ROW_EVEN: "#f4f8f9",       // Helles Grautürkis
  FORMAT_ROW_ODD: "#ffffff",        // Weiß
  FORMAT_BORDER_COLOR: "#000000",   // Schwarz
  HEADER_FONT_SIZE: 11,
  COL_WIDTH_BUFFER: 20,             // Extra-Pixel für Filter-Pfeile
  MAX_COL_WIDTH: 250                // Maximale Spaltenbreite
};