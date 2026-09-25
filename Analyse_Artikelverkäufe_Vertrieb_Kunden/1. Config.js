// ==========================================
// 1. CONFIG.GS - Zentrale Einstellungen
// ==========================================
const CONFIG = {
  API: {
    URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
    TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
    MAX_PAGES: 500
  },
  SHEETS: {
    VORGANGSARTEN_URL: "https://docs.google.com/spreadsheets/d/1xyKAfpitLrJ28xUnOIKYTX9pFk3SBa9iwyateMMIGoQ/edit",
    VORGANGSARTEN_TAB: "Vorgangsarten",
    RAW_DATA: "Archiv_Rohdaten_24M",
    PRODUCT_ANALYSIS: "Produktanalyse",
    REP_COMPARISON: "Analyse_Artikel_Vertreter",
    COUNTRY_COMPARISON: "Analyse_Artikel_Laender"
  },
  CATALOG_FILTER: "1", // Strikte Begrenzung auf Katalog 1
  REPS: {
    "28": { name: "Alexis Fonte", lang: "DE" },
    "36": { name: "Imad Nassef", lang: "DE" },
    "43": { name: "Andreas Steinert", lang: "DE" },
    "46": { name: "Stefanie Binder", lang: "DE" },
    "56": { name: "Robin Carter Browne", lang: "EN" },
    "59": { name: "Aneta Nedyalkova", lang: "DE" },
    "60": { name: "Rado Kabakov", lang: "EN" }
  },
  // Vorgangsarten für Korrekturen/Gutschriften/Erstattungen (Vorzeichenumkehr)[cite: 17]
  CREDIT_TYPES: ["123", "90", "156"],
  
  // Farbcodes für automatische Auswertungen
  COLOR_PALETTE: {
    BEST: "#c8e6c9",          // 1. Platz (Grün)
    SECOND_BEST: "#fef08a",   // 2. Platz (Gelb)
    WORST: "#f8d7da",         // Schlechtest / 0 (Rot)
    STRONG_GROWTH: "#c8e6c9", // Zeilenfarbe: Starkes Wachstum
    GROWTH: "#e8f5e9",        // Zeilenfarbe: Wachstum
    DECLINE: "#fff3cd",       // Zeilenfarbe: Rückgang
    STRONG_DECLINE: "#f8d7da",// Zeilenfarbe: Starker Abfall
    STOPPED: "#f5c6cb",       // Zeilenfarbe: Inaktiv / Kauf gestoppt
    NEVER_PURCHASED: "#e3f2fd"// Zeilenfarbe: Nie gekauft / Never purchased (Eisblau - Push-Potenzial)
  },

  HEADERS_RAW: [
    "Datum", "VtrNr", "Vertreter Name", "AdrNr", "Kundenname", "Land", 
    "BelegNr", "VorgangsArt", "ArtikelNr", "Artikelname", "Katalog", "Menge", "Netto_Umsatz", "Quartal_Label", "IsReseller"
  ]
};