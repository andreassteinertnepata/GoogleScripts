/**
 * Config.gs - Globale Konfiguration
 */

const CONFIG = {
  // --- API Zugangsdaten ---
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",

  // --- Limits & Filter ---
  MAX_EXECUTION_TIME_MS: 5.5 * 60 * 1000,
  PAGE_SIZE: 100,
  TARGET_CATALOG: "1",
  MIN_EK: 50.0,                   // Mindest-EK-Filter in Euro für Katalog 1 Artikel
  KEYACCOUNT_MIN_UMSATZ: 25000.0, // Schwellenwert: Mindestens 25.000 € Umsatz in 12M

  // --- Ausgeschlossene Artikelnummern & Belegarten ---
  EXCLUDED_ARTICLES: ["109.107.13", "109.107.15", "109.011.17", "109.107.16"],
  EXCLUDED_BELEGARTEN: ["150"],   // Belegart 150 wird niemals berücksichtigt!

  // Standard-Fallbacks für Archiv-Belegarten (falls Blatt 'Vorgangsarten' fehlt/leer ist)
  DEFAULT_ARCHIV_TYPES: ["70", "105", "109", "110", "113", "115", "122", "123", "129", "154", "155", "156"],
  DEFAULT_VORGAENGE_TYPES: ["30", "106", "107", "108", "112", "114", "118", "120", "127", "152", "163", "164", "167", "185", "188", "189", "198", "200"],

  // Gutschriften / Erstattungen (Mengen werden negiert)
  CORRECTION_TYPES: ["90", "123", "156"],

  // --- Tabellenblätter ---
  SHEET_CONFIG: "Ziele_Konfiguration",
  SHEET_PRODUKTZIELE: "Produktziele_Konfiguration",
  SHEET_KEYACCOUNTS: "KeyAccounts_Konfiguration",
  SHEET_ARTIKEL: "Artikel",
  SHEET_VORGANGSARTEN: "Vorgangsarten",

  // Farbskala für Zielerreichung
  COLOR_OVER_100: "#cfe2f3", // Hellblau (>100%)
  COLOR_90_TO_99: "#d9ead3", // Hellgrün (>=90%)
  COLOR_50_TO_89: "#fff2cc", // Gelb (>=50%)
  COLOR_UNDER_50: "#f4ccd0", // Rot (<50%)

  // Hervorhebungsfarbe für manuell pflegbare Zielspalten (Eingabefelder)
  COLOR_TARGET_EDITABLE: "#ffe599", // Helles Gelb/Orange für geschützte Eingabespalten

  // Vertriebsmitarbeiter-Zuordnung
  SALES_REPS: {
    "28": { name: "Alexis Fonte", lang: "DE", tab: "Dash_Alexis" },
    "36": { name: "Imad Nassef", lang: "DE", tab: "Dash_Imad" },
    "43": { name: "Andreas Steinert", lang: "DE", tab: "Dash_Andreas" },
    "46": { name: "Stefanie Binder", lang: "DE", tab: "Dash_Stefanie" },
    "56": { name: "Robin Carter Browne", lang: "EN", tab: "Dash_Robin" },
    "59": { name: "Aneta Nedyalkova", lang: "DE", tab: "Dash_Aneta" },
    "60": { name: "Rado Kabakov", lang: "EN", tab: "Dash_Rado" }
  }
};